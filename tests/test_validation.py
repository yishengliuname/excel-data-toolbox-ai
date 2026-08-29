from __future__ import annotations

import json

import pandas as pd
import pandas.testing as pdt
import pytest

from excel_data_toolbox.validation import (
    FAILURE_COLUMNS,
    ValidationRule,
    validate_dataframe,
)


def test_validation_rule_roundtrip_and_rejects_unsafe_or_unknown_config() -> None:
    rule = ValidationRule(
        rule_id="订单号格式",
        rule_type="regex",
        column="订单号",
        severity="warning",
        params={"pattern": r"^ORD-\d{6}$", "mode": "fullmatch"},
        message="订单号格式不正确",
    )

    restored = ValidationRule.from_dict(rule.to_dict())

    assert restored == rule
    assert json.loads(json.dumps(restored.to_dict(), ensure_ascii=False))["rule_id"] == "订单号格式"
    with pytest.raises(ValueError, match="未知字段"):
        ValidationRule.from_dict(
            {"rule_id": "x", "rule_type": "not_null", "column": "A", "code": "eval()"}
        )
    with pytest.raises(ValueError, match="不支持的验证规则"):
        ValidationRule("x", "python", "A", params={"callable": "os.system"})
    with pytest.raises(ValueError, match="嵌套重复"):
        ValidationRule("x", "regex", "A", params={"pattern": r"^(a+)+$"})
    with pytest.raises(ValueError, match="环视"):
        ValidationRule("x", "regex", "A", params={"pattern": r"^(?=a)a$"})
    with pytest.raises(ValueError, match="severity"):
        ValidationRule("x", "not_null", "A", severity="critical")
    with pytest.raises(ValueError, match="不是有效数字"):
        ValidationRule("x", "range", "A", params={"min": "不是数字"})
    with pytest.raises(ValueError, match="min 不能大于"):
        ValidationRule("x", "range", "A", params={"min": 10, "max": 1})
    with pytest.raises(ValueError, match="min 不能大于"):
        ValidationRule(
            "x", "date", "A", params={"min": "2026-12-31", "max": "2026-01-01"}
        )


def test_all_rule_types_return_long_form_failures_without_mutation() -> None:
    frame = pd.DataFrame(
        {
            "订单号": [
                "9007199254740993",
                "9007199254740993",
                "9007199254740994",
                None,
            ],
            "金额": ["10.00", "-1", "bad", "20"],
            "状态": ["完成", "未知", "完成", "完成"],
            "日期": ["2026-01-01", "错误日期", None, "2027-01-01"],
            "备注": ["A1", "A!", "B2", None],
            "应付": ["10", "5", "3", "20"],
            "实付": ["10", "4", "5", "20"],
        }
    )
    untouched = frame.copy(deep=True)
    rules = [
        ValidationRule("订单必填", "not_null", "订单号"),
        ValidationRule("订单唯一", "unique", "订单号"),
        ValidationRule(
            "金额范围", "range", "金额", params={"min": 0, "max": 100}
        ),
        ValidationRule(
            "备注格式", "regex", "备注", params={"pattern": r"^[A-Z]\d$"}
        ),
        ValidationRule(
            "状态集合",
            "allowed_values",
            "状态",
            params={"values": ["完成", "取消"]},
        ),
        ValidationRule("金额数字", "numeric", "金额"),
        ValidationRule(
            "有效日期",
            "date",
            "日期",
            params={"min": "2026-01-01", "max": "2026-12-31"},
        ),
        ValidationRule(
            "实付不少于应付",
            "column_compare",
            "实付",
            params={
                "other_column": "应付",
                "operator": "gte",
                "value_type": "numeric",
            },
        ),
    ]

    report = validate_dataframe(frame, rules, include_values=True)

    pdt.assert_frame_equal(frame, untouched)
    assert report.passed is False
    assert report.rule_count == 8
    assert report.failed_rule_count == 8
    assert report.blocking_failure_count == report.failure_count
    assert list(report.failures.columns) == FAILURE_COLUMNS
    failures = report.failures
    assert set(failures.loc[failures["rule_id"] == "订单必填", "row_position"]) == {3}
    assert set(failures.loc[failures["rule_id"] == "订单唯一", "row_position"]) == {0, 1}
    assert set(failures.loc[failures["rule_id"] == "金额范围", "code"]) == {
        "out_of_range",
        "not_numeric",
    }
    assert failures.loc[
        (failures["rule_id"] == "订单唯一") & (failures["row_position"] == 0),
        "value_preview",
    ].iloc[0] == "9007199254740993"
    assert set(failures.loc[failures["rule_id"] == "有效日期", "row_position"]) == {1, 3}
    assert set(
        failures.loc[failures["rule_id"] == "实付不少于应付", "row_position"]
    ) == {1}
    assert report.rule_results_frame().shape[0] == 8
    assert report.to_dict()["failure_count"] == len(failures)


def test_warning_and_info_violations_do_not_block_validation() -> None:
    frame = pd.DataFrame({"状态": ["完成", "未知"], "备注": [None, None]})
    report = validate_dataframe(
        frame,
        [
            ValidationRule(
                "状态建议",
                "allowed_values",
                "状态",
                severity="warning",
                params={"values": ["完成"]},
            ),
            ValidationRule(
                "备注提示",
                "not_null",
                "备注",
                severity="info",
            ),
        ],
    )

    assert report.passed is True
    assert report.failure_count == 3
    assert report.blocking_failure_count == 0
    assert report.severity_counts == {"error": 0, "warning": 1, "info": 2}
    assert report.failed_rule_count == 2
    assert report.to_dict()["passed"] is True


def test_missing_columns_are_structured_and_disabled_rules_are_skipped() -> None:
    frame = pd.DataFrame({"已有列": [1, 2]})
    report = validate_dataframe(
        frame,
        [
            ValidationRule("缺列", "not_null", "不存在"),
            ValidationRule("跳过", "not_null", "也不存在", enabled=False),
            ValidationRule(
                "比较缺列",
                "column_compare",
                "已有列",
                params={"other_column": "另一列", "operator": "eq"},
            ),
        ],
    )

    assert report.failure_count == 2
    assert set(report.failures["code"]) == {"missing_column"}
    assert report.failures["row_position"].isna().all()
    skipped = next(result for result in report.rule_results if result.rule_id == "跳过")
    assert skipped.skipped is True
    assert skipped.passed is True
    assert skipped.checked_count == 0


def test_long_numeric_identifiers_are_checked_exactly_without_coercion() -> None:
    frame = pd.DataFrame(
        {
            "订单号": [
                "9007199254740993",
                "9007199254740994",
                "9007199254740995",
            ]
        }
    )
    untouched = frame.copy(deep=True)
    report = validate_dataframe(
        frame,
        [
            ValidationRule("数字文本", "numeric", "订单号"),
            ValidationRule(
                "精确范围",
                "range",
                "订单号",
                params={"min": 9007199254740994, "max": 9007199254740994},
            ),
            ValidationRule(
                "允许订单",
                "allowed_values",
                "订单号",
                params={"values": ["9007199254740993", "9007199254740994"]},
            ),
        ],
        include_values=True,
    )

    pdt.assert_frame_equal(frame, untouched)
    numeric = next(result for result in report.rule_results if result.rule_id == "数字文本")
    assert numeric.passed is True
    range_failures = report.failures.loc[report.failures["rule_id"] == "精确范围"]
    assert range_failures["value_preview"].tolist() == [
        "9007199254740993",
        "9007199254740995",
    ]
    allowed = report.failures.loc[report.failures["rule_id"] == "允许订单"]
    assert allowed["value_preview"].tolist() == ["9007199254740995"]


def test_null_blank_unique_and_numeric_options_are_explicit() -> None:
    frame = pd.DataFrame(
        {"编号": [None, "", "   ", "A", "A"], "数量": [None, "1", "1.5", "inf", "x"]}
    )
    report = validate_dataframe(
        frame,
        [
            ValidationRule("非空", "not_null", "编号"),
            ValidationRule(
                "空值也唯一",
                "unique",
                "编号",
                params={"ignore_nulls": False, "blank_as_null": True},
            ),
            ValidationRule(
                "必须整数",
                "numeric",
                "数量",
                params={"ignore_nulls": False, "integer_only": True},
            ),
        ],
    )

    assert set(report.failures.loc[report.failures["rule_id"] == "非空", "row_position"]) == {
        0,
        1,
        2,
    }
    assert set(
        report.failures.loc[report.failures["rule_id"] == "空值也唯一", "row_position"]
    ) == {0, 1, 2, 3, 4}
    numeric_codes = report.failures.loc[
        report.failures["rule_id"] == "必须整数", "code"
    ].tolist()
    assert numeric_codes.count("null_value") == 1
    assert numeric_codes.count("not_integer") == 1
    assert numeric_codes.count("not_numeric") == 2


def test_date_format_range_and_native_column_comparison() -> None:
    frame = pd.DataFrame(
        {
            "开始": ["2026/01/01", "2026/03/01", "错误"],
            "结束": ["2026/01/02", "2026/02/01", "2026/03/01"],
            "版本A": ["001", "002", "003"],
            "版本B": ["001", "2", "004"],
        }
    )
    report = validate_dataframe(
        frame,
        [
            ValidationRule(
                "开始日期",
                "date",
                "开始",
                params={"format": "%Y/%m/%d", "max": "2026/12/31"},
            ),
            ValidationRule(
                "日期顺序",
                "column_compare",
                "结束",
                params={"other_column": "开始", "operator": "gte", "value_type": "date"},
            ),
            ValidationRule(
                "版本严格相同",
                "column_compare",
                "版本A",
                params={"other_column": "版本B", "operator": "eq"},
            ),
        ],
    )

    assert set(report.failures.loc[report.failures["rule_id"] == "开始日期", "row_position"]) == {2}
    assert set(report.failures.loc[report.failures["rule_id"] == "日期顺序", "row_position"]) == {
        1,
        2,
    }
    assert set(
        report.failures.loc[report.failures["rule_id"] == "版本严格相同", "row_position"]
    ) == {1, 2}


def test_values_are_private_by_default_and_bounded_when_opted_in() -> None:
    secret = "客户隐私-" + "甲" * 100
    frame = pd.DataFrame({"敏感列": [secret]})
    rule = ValidationRule(
        "限制",
        "allowed_values",
        "敏感列",
        params={"values": ["允许值"]},
    )

    private = validate_dataframe(frame, [rule])
    visible = validate_dataframe(frame, [rule], include_values=True, max_value_chars=12)

    assert private.failures.loc[0, "value_preview"] is None
    assert secret not in json.dumps(private.to_dict(include_failures=True), ensure_ascii=False)
    assert visible.failures.loc[0, "value_preview"] == "客户隐私-甲甲甲甲甲甲…"
    copied = private.failures_frame()
    copied.loc[0, "message"] = "已改"
    assert private.failures.loc[0, "message"] != "已改"


def test_empty_dataframe_and_duplicate_rule_ids_have_clear_results() -> None:
    empty = pd.DataFrame({"编号": pd.Series(dtype="string")})
    report = validate_dataframe(
        empty,
        [ValidationRule("非空", "not_null", "编号")],
    )

    assert report.passed is True
    assert report.failure_count == 0
    assert report.failures.empty
    assert str(report.failures["row_position"].dtype) == "Int64"
    with pytest.raises(ValueError, match="rule_id 不能重复"):
        validate_dataframe(
            empty,
            [
                ValidationRule("重复", "not_null", "编号"),
                ValidationRule("重复", "unique", "编号"),
            ],
        )
    with pytest.raises(ValueError, match="重复列名"):
        validate_dataframe(
            pd.DataFrame([[1, 2]], columns=["A", "A"]),
            [ValidationRule("x", "not_null", "A")],
        )
