from __future__ import annotations

import json

import pandas as pd
import pandas.testing as pdt
import pytest

import excel_data_toolbox.nl_agent as nl
from excel_data_toolbox.nl_agent import (
    AgentExecutionError,
    DeepSeekAPIError,
    DeepSeekClient,
    PlanValidationError,
    UnsupportedPlanError,
    build_table_catalog,
    execute_plan,
    preview_plan,
    validate_plan,
)


def _payload(*steps: dict, status: str = "ready") -> dict:
    return {
        "schema_version": 1,
        "status": status,
        "summary": "处理订单并生成可核验结果",
        "message": "计划已生成" if status == "ready" else "需要补充信息",
        "clarification_questions": [],
        "assumptions": [],
        "warnings": [],
        "steps": list(steps),
    }


def _step(
    step_id: str,
    operation: str,
    input_ids: list[str],
    output_name: str,
    params: dict,
) -> dict:
    return {
        "id": step_id,
        "operation": operation,
        "input_ids": input_ids,
        "output_name": output_name,
        "params": params,
    }


def test_catalog_is_schema_only_by_default_and_samples_are_explicit() -> None:
    frame = pd.DataFrame(
        {
            "客户姓名": ["张三绝密", "李四绝密"],
            "手机号": ["13800138000", "13900139000"],
            "金额": [12.5, 20.0],
        }
    )

    catalog = build_table_catalog({"t_customer": frame})
    payload = json.dumps(catalog, ensure_ascii=False)

    assert "张三绝密" not in payload
    assert "13800138000" not in payload
    assert "客户姓名" in payload
    assert "redacted_samples" not in payload
    assert catalog["tables"][0]["row_count"] == 2
    assert catalog["tables"][0]["columns"][0]["unique_count"] == 2

    with_samples = build_table_catalog(
        {"t_customer": frame},
        redacted_samples={
            "t_customer": [{"客户姓名": "张*", "手机号": "138****8000"}]
        },
    )
    assert with_samples["tables"][0]["redacted_samples"] == [
        {"客户姓名": "张*", "手机号": "138****8000"}
    ]
    with pytest.raises(PlanValidationError, match="最多 3 行"):
        build_table_catalog(
            {"t_customer": frame},
            redacted_samples={"t_customer": [{}, {}, {}, {}]},
        )

    poisoned_catalog = dict(catalog)
    poisoned_catalog["tables"] = [
        {**catalog["tables"][0], "raw_rows": [{"客户姓名": "不应外发"}]}
    ]
    with pytest.raises(PlanValidationError, match="字段不合法"):
        DeepSeekClient("sk-test").create_plan("检查数据", poisoned_catalog)


def test_json_parser_tolerates_prose_and_client_retries_once() -> None:
    embedded = nl._parse_model_json('以下是结果：```json\n{"intent":"data"}\n```', label="测试")
    assert embedded == {"intent": "data"}

    client = DeepSeekClient("sk-test")
    calls = []
    valid = {
        "intent": "data", "normalized_request": "汇总数据", "data_request": "汇总数据",
        "chart_request": "", "engineering_category": None, "reason": "数据处理",
    }
    responses = [
        json.dumps({"choices": [{"finish_reason": "stop", "message": {"content": "抱歉，我无法输出 JSON"}}]}).encode(),
        json.dumps({"choices": [{"finish_reason": "stop", "message": {"content": json.dumps(valid, ensure_ascii=False)}}]}, ensure_ascii=False).encode(),
    ]

    def fake_request(body, *, response_limit):
        calls.append(body)
        return responses.pop(0)

    client._request = fake_request
    catalog = build_table_catalog({"t": pd.DataFrame({"金额": [1]})})
    routed = client.classify_unified_request("请汇总当前数据", catalog)
    assert routed["intent"] == "data"
    assert len(calls) == 2
    assert "上一条响应未通过" in calls[1]["messages"][-1]["content"]


def test_plan_schema_rejects_unknown_fields_operations_and_code_channels() -> None:
    catalog = build_table_catalog({"orders": pd.DataFrame({"订单号": ["A"]})})

    unsafe = _payload(
        _step(
            "evil",
            "python",
            ["orders"],
            "结果",
            {"code": "__import__('os').system('whoami')"},
        )
    )
    with pytest.raises(UnsupportedPlanError, match="无法完成操作"):
        validate_plan(unsafe, catalog)

    path_smuggling = _payload(
        _step(
            "clean1",
            "clean",
            ["orders"],
            "结果",
            {"path": "C:/secret", "sql": "DROP TABLE x"},
        )
    )
    with pytest.raises(PlanValidationError, match="未知字段"):
        validate_plan(path_smuggling, catalog)

    unknown_top = _payload()
    unknown_top["api_key"] = "must-not-be-accepted"
    with pytest.raises(PlanValidationError, match="计划字段不合法"):
        validate_plan(unknown_top, catalog)


def test_clean_model_aliases_are_normalized_before_strict_validation() -> None:
    catalog = build_table_catalog(
        {"orders": pd.DataFrame({"日期": ["2026-08-01"], "金额": [10.0]})}
    )
    plan = validate_plan(
        _payload(
            _step(
                "clean1",
                "clean",
                ["orders"],
                "清洗结果",
                {"fill_missing": True, "date_format": "%Y-%m-%d"},
            )
        ),
        catalog,
    )

    assert plan.status == "ready"
    assert plan.steps[0].params["missing_strategy"] == "fill"
    assert plan.steps[0].params["infer_types"] is True
    assert "fill_missing" not in plan.steps[0].params
    assert "date_format" not in plan.steps[0].params
    assert any("安全转换" in warning for warning in plan.warnings)


def test_engineering_brief_rejects_dangerous_vba_and_database_writes() -> None:
    base = {
        "schema_version": 1,
        "status": "ready",
        "category": "vba",
        "normalized_request": "生成受控的月度合并宏方案",
        "scope": "只生成代码文本，不执行宏。",
        "clarification_questions": [],
        "deliverables": ["标准模块"],
        "implementation_steps": ["人工审查后导入模块"],
        "artifacts": [
            {
                "name": "安全模块",
                "language": "vba",
                "content": "Sub MergeReports()\n    MsgBox \"Ready\"\nEnd Sub",
                "usage_note": "请先在副本测试",
            }
        ],
        "test_checklist": ["源文件不覆盖"],
        "risks": ["需启用宏"],
        "human_approval_points": ["运行前确认备份"],
    }
    safe = nl._validate_engineering_brief(base, "vba")
    assert safe["status"] == "ready"

    dangerous_vba = json.loads(json.dumps(base, ensure_ascii=False))
    dangerous_vba["artifacts"][0]["content"] = 'Shell "cmd.exe /c whoami"'
    with pytest.raises(PlanValidationError, match="禁止"):
        nl._validate_engineering_brief(dangerous_vba, "vba")

    write_sql = json.loads(json.dumps(base, ensure_ascii=False))
    write_sql["category"] = "database"
    write_sql["artifacts"][0].update(
        {"language": "sql", "content": "DELETE FROM orders"}
    )
    with pytest.raises(PlanValidationError, match="只读"):
        nl._validate_engineering_brief(write_sql, "database")


def test_missing_parameters_or_columns_become_clarification() -> None:
    catalog = build_table_catalog(
        {
            "erp": pd.DataFrame({"订单号": ["A"], "应收金额": [100]}),
            "bank": pd.DataFrame({"流水号": ["B"], "实收金额": [100]}),
        }
    )
    missing_key = validate_plan(
        _payload(_step("join1", "join", ["erp", "bank"], "连接结果", {})),
        catalog,
    )
    assert missing_key.status == "clarification"
    assert "连接需要明确" in missing_key.clarification_questions[0]
    assert missing_key.steps == ()

    bad_column = validate_plan(
        _payload(
            _step(
                "sum1",
                "summary",
                ["erp"],
                "汇总",
                {"by": "不存在的区域", "aggregations": {"应收金额": "sum"}},
            )
        ),
        catalog,
    )
    assert bad_column.status == "clarification"
    assert "不存在的字段" in bad_column.clarification_questions[0]

    bad_reconcile_column = validate_plan(
        _payload(
            _step(
                "recon1",
                "reconcile",
                ["erp", "bank"],
                "对账",
                {
                    "left_amount": "不存在的应收",
                    "right_amount": "实收金额",
                    "left_key_columns": ["订单号"],
                    "right_key_columns": ["流水号"],
                },
            )
        ),
        catalog,
    )
    assert bad_reconcile_column.status == "clarification"
    assert "左表缺少字段" in bad_reconcile_column.clarification_questions[0]


def test_mask_plan_validates_keep_bounds_and_executes() -> None:
    tables = {
        "orders": pd.DataFrame(
            {"手机号": ["13800138000"], "邮箱": ["buyer@example.test"]}
        )
    }
    catalog = build_table_catalog(tables)
    plan = validate_plan(
        _payload(
            _step(
                "mask1",
                "mask",
                ["orders"],
                "脱敏交付版",
                {
                    "columns": {"手机号": "phone", "邮箱": "email"},
                    "keep_start": 1,
                    "keep_end": 1,
                },
            )
        ),
        catalog,
    )

    result = execute_plan(plan, tables, dry_run=True)

    assert plan.status == "ready"
    assert result.tables["脱敏交付版"].loc[0, "手机号"] != "13800138000"
    assert result.tables["脱敏交付版"].loc[0, "邮箱"] != "buyer@example.test"


def test_complex_chained_plan_executes_on_copies_and_returns_safe_reports() -> None:
    orders = pd.DataFrame(
        {
            "订单号": ["A001", "A001", "A002", "A003", "A004"],
            "区域": ["华东 ", "华东 ", "华北", "华北", "华南"],
            "状态": ["完成", "完成", "完成", "取消", "完成"],
            "金额": [1200.0, 1200.0, 800.0, -10.0, None],
        }
    )
    untouched = orders.copy(deep=True)
    catalog = build_table_catalog({"orders": orders})
    plan = validate_plan(
        _payload(
            _step(
                "clean1",
                "recipe",
                ["orders"],
                "订单标准化",
                {
                    "name": "生产订单清理",
                    "description": "去重、去空格、仅保留完成订单",
                    "steps": [
                        {
                            "operation": "clean",
                            "params": {
                                "trim_whitespace": True,
                                "drop_duplicates": True,
                                "missing_strategy": "keep",
                                "infer_types": False,
                            },
                        },
                        {
                            "operation": "filter",
                            "params": {
                                "conditions": [
                                    {"column": "状态", "operator": "eq", "value": "完成"}
                                ]
                            },
                        },
                    ],
                },
            ),
            _step(
                "check1",
                "validate",
                ["$clean1"],
                "订单验收",
                {
                    "rules": [
                        {
                            "rule_id": "订单唯一",
                            "rule_type": "unique",
                            "column": "订单号",
                            "severity": "error",
                            "params": {"ignore_nulls": False},
                        },
                        {
                            "rule_id": "金额非负",
                            "rule_type": "range",
                            "column": "金额",
                            "severity": "error",
                            "params": {"min": 0, "ignore_nulls": False},
                        },
                    ]
                },
            ),
            _step(
                "summary1",
                "summary",
                ["$clean1"],
                "区域销售汇总",
                {"by": "区域", "aggregations": {"金额": ["sum", "count"]}},
            ),
        ),
        catalog,
    )

    preview = preview_plan(plan, {"orders": orders})
    result = execute_plan(plan, {"orders": orders}, dry_run=True)

    assert preview.executable is True
    assert preview.step_count == 3
    pdt.assert_frame_equal(orders, untouched)
    assert result.dry_run is True
    assert result.reports["clean1"]["output_fingerprint"]["row_count"] == 3
    assert result.reports["check1"]["failure_count"] == 1
    assert any(name.startswith("区域销售汇总") for name in result.tables)
    assert "A004" not in json.dumps(result.to_dict()["reports"], ensure_ascii=False)


def test_advanced_reconciliation_plan_exposes_review_artifacts() -> None:
    erp = pd.DataFrame(
        {
            "订单号": ["O-001", "O-002", "O-003", "O-004", "O-004"],
            "记账日": ["2026-08-01", "2026-08-02", "2026-08-03", "2026-08-04", "2026-08-04"],
            "应收金额": [1000, 2500, 899.99, 400, 600],
            "客户": ["甲", "乙", "丙", "丁", "丁"],
        }
    )
    bank = pd.DataFrame(
        {
            "商户订单号": ["O-001", "O-002", "O-003", "O-004"],
            "到账日": ["2026-08-01", "2026-08-04", "2026-08-03", "2026-08-04"],
            "到账金额": [1000.00, 2500.00, 900.00, 1000.00],
            "付款方": ["甲", "乙", "丙", "丁"],
        }
    )
    catalog = build_table_catalog({"erp": erp, "bank": bank})
    plan = validate_plan(
        _payload(
            _step(
                "recon1",
                "reconcile",
                ["erp", "bank"],
                "八月资金对账",
                {
                    "left_amount": "应收金额",
                    "right_amount": "到账金额",
                    "left_date": "记账日",
                    "right_date": "到账日",
                    "left_key_columns": ["订单号"],
                    "right_key_columns": ["商户订单号"],
                    "left_secondary_columns": ["客户"],
                    "right_secondary_columns": ["付款方"],
                    "amount_tolerance": "0.02",
                    "date_tolerance_days": 2,
                    "enable_split_candidates": True,
                },
            )
        ),
        catalog,
    )

    result = execute_plan(plan, {"erp": erp, "bank": bank}, dry_run=False)

    assert result.reports["recon1"]["matched_count"] >= 1
    assert result.reports["recon1"]["duplicate_rows_count"] >= 2
    assert any(name.endswith("_review") for name in result.tables)
    assert any(name.endswith("_duplicates") for name in result.tables)
    assert len(result.tables["八月资金对账_review"]) >= 0


class _FakeResponse:
    def __init__(self, payload: dict) -> None:
        self.payload = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    def __enter__(self) -> "_FakeResponse":
        return self

    def __exit__(self, *_args: object) -> None:
        return None

    def read(self, _limit: int) -> bytes:
        return self.payload


def test_deepseek_client_uses_current_model_strict_json_and_never_sends_values(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    raw = pd.DataFrame({"客户": ["绝密客户名称"], "金额": [999999]})
    catalog = build_table_catalog({"orders": raw})
    response_plan = _payload(
        _step("quality1", "quality", ["orders"], "质量体检", {})
    )
    captured: dict[str, object] = {}

    def fake_urlopen(request: object, timeout: int) -> _FakeResponse:
        captured["request"] = request
        captured["timeout"] = timeout
        return _FakeResponse(
            {
                "choices": [
                    {
                        "finish_reason": "stop",
                        "message": {
                            "role": "assistant",
                            "content": json.dumps(response_plan, ensure_ascii=False),
                        },
                    }
                ]
            }
        )

    monkeypatch.setattr(nl.urllib_request, "urlopen", fake_urlopen)
    client = DeepSeekClient("sk-super-secret", model="deepseek-chat")
    plan = client.create_plan("请检查订单表质量", catalog)

    request = captured["request"]
    request_body = request.data.decode("utf-8")  # type: ignore[attr-defined]
    assert plan.status == "ready"
    assert client.model == "deepseek-v4-flash"
    assert "绝密客户名称" not in request_body
    assert "999999" not in request_body
    assert "sk-super-secret" not in request_body
    assert "sk-super-secret" not in repr(client)
    assert request.full_url == "https://api.deepseek.com/chat/completions"  # type: ignore[attr-defined]


@pytest.mark.parametrize(
    ("finish_reason", "content", "match"),
    [
        ("length", "{}", "长度限制"),
        ("content_filter", "{}", "内容安全"),
        ("insufficient_system_resource", "{}", "资源暂时不足"),
        ("stop", "", "空计划"),
    ],
)
def test_deepseek_client_reports_incomplete_or_empty_responses(
    monkeypatch: pytest.MonkeyPatch,
    finish_reason: str,
    content: str,
    match: str,
) -> None:
    catalog = build_table_catalog({"orders": pd.DataFrame({"A": [1]})})

    def fake_urlopen(_request: object, timeout: int) -> _FakeResponse:
        assert timeout == 60
        return _FakeResponse(
            {
                "choices": [
                    {
                        "finish_reason": finish_reason,
                        "message": {"role": "assistant", "content": content},
                    }
                ]
            }
        )

    monkeypatch.setattr(nl.urllib_request, "urlopen", fake_urlopen)
    with pytest.raises(DeepSeekAPIError, match=match):
        DeepSeekClient("sk-test").create_plan("检查", catalog)


def test_non_ready_plan_cannot_execute() -> None:
    catalog = build_table_catalog({"orders": pd.DataFrame({"A": [1]})})
    payload = _payload(status="clarification")
    payload["clarification_questions"] = ["请问要处理哪一列？"]
    payload["steps"] = []
    plan = validate_plan(payload, catalog)

    with pytest.raises(UnsupportedPlanError, match="ready"):
        execute_plan(plan, {"orders": pd.DataFrame({"A": [1]})})
