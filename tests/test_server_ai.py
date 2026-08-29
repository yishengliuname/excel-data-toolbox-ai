from __future__ import annotations

from dataclasses import replace
import json
import time
from typing import Any, Callable

import pandas as pd
import pytest

import excel_data_toolbox.server as server_module
from excel_data_toolbox.nl_agent import DeepSeekAPIError, validate_plan
from excel_data_toolbox.server import AppSession, ToolboxHandler


SECRET_KEY = "sk-unit-test-ultra-secret-123456789"


def _plan_payload(
    table_id: str,
    *,
    status: str = "ready",
) -> dict[str, Any]:
    questions = ["请明确需要处理的金额字段。"] if status == "clarification" else []
    steps: list[dict[str, Any]] = []
    if status == "ready":
        steps = [
            {
                "id": "clean1",
                "operation": "clean",
                "input_ids": [table_id],
                "output_name": "AI清洗结果",
                "params": {
                    "trim_whitespace": True,
                    "drop_duplicates": True,
                    "infer_types": False,
                },
            }
        ]
    return {
        "schema_version": 1,
        "status": status,
        "summary": "清理订单并生成独立结果表",
        "message": (
            "计划可以执行"
            if status == "ready"
            else "请补充信息" if status == "clarification" else "当前能力不支持"
        ),
        "clarification_questions": questions,
        "assumptions": [],
        "warnings": [],
        "steps": steps,
    }


@pytest.fixture
def ai_session(monkeypatch: pytest.MonkeyPatch):
    session = AppSession()
    monkeypatch.setattr(server_module, "SESSION", session)
    handler = object.__new__(ToolboxHandler)
    try:
        yield session, handler
    finally:
        session.close()


def _add_orders(session: AppSession) -> str:
    return session.add_table(
        "生产订单",
        pd.DataFrame(
            {
                "订单号": ["A001", "A001", "A002"],
                "客户": [" 甲公司 ", " 甲公司 ", "乙公司"],
                "金额": [100.0, 100.0, 250.0],
            }
        ),
        source="单元测试",
        original=True,
    )


def _install_model_plan(
    monkeypatch: pytest.MonkeyPatch,
    factory: Callable[[dict[str, Any]], dict[str, Any]],
    *,
    captured: dict[str, Any] | None = None,
) -> None:
    def fake_create_plan(self: object, prompt: str, catalog: dict[str, Any]):
        if captured is not None:
            captured["client_repr"] = repr(self)
            captured["prompt"] = prompt
            captured["catalog"] = catalog
        return validate_plan(factory(catalog), catalog)

    monkeypatch.setattr(server_module.DeepSeekClient, "create_plan", fake_create_plan)


def _request_plan(
    handler: ToolboxHandler,
    table_id: str,
    *,
    prompt: str = "清理订单中的空格和重复记录，并生成一张结果表供人工核对",
) -> dict[str, Any]:
    return handler._ai_plan(
        {
            "prompt": prompt,
            "api_key": SECRET_KEY,
            "model": "deepseek-v4-flash",
            "table_ids": [table_id],
        }
    )


def test_ai_capabilities_expose_only_current_v4_models(ai_session) -> None:
    _session, handler = ai_session

    payload = handler._ai_capabilities()
    rendered = json.dumps(payload, ensure_ascii=False)

    assert set(payload["models"]) == {"deepseek-v4-flash", "deepseek-v4-pro"}
    assert payload["default_model"] == "deepseek-v4-flash"
    assert "deepseek-chat" not in rendered
    assert "deepseek-reasoner" not in rendered
    assert payload["privacy"]["local_execution"] is True


def test_ai_table_scope_accepts_real_world_thirteen_sheet_workbook() -> None:
    table_ids = [f"table-{index:02d}" for index in range(13)]

    assert server_module._normalise_ai_table_scope(table_ids, allow_empty=False) == table_ids


def test_ai_table_scope_reports_actionable_shape_and_limit_errors() -> None:
    with pytest.raises(server_module.ApiError, match="必须是表 ID 列表"):
        server_module._normalise_ai_table_scope("table-01", allow_empty=True)
    with pytest.raises(server_module.ApiError, match="101 张表.*最多支持 100 张"):
        server_module._normalise_ai_table_scope(
            [f"table-{index:03d}" for index in range(101)],
            allow_empty=True,
        )


def test_ready_plan_requires_confirmation_commits_once_and_never_stores_key(
    ai_session,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session, handler = ai_session
    table_id = _add_orders(session)
    captured: dict[str, Any] = {}
    _install_model_plan(
        monkeypatch,
        lambda catalog: _plan_payload(catalog["tables"][0]["table_id"]),
        captured=captured,
    )

    response = _request_plan(handler, table_id)
    token = response["plan_token"]

    assert response["status"] == "ready"
    assert response["normalized_request"] == "清理订单并生成独立结果表"
    assert response["normalized_request"] == response["plan"]["summary"]
    assert isinstance(token, str) and len(token) >= 32
    assert response["preview"]["executable"] is True
    assert response["dry_run"]["dry_run"] is True
    assert SECRET_KEY not in json.dumps(response, ensure_ascii=False)
    assert SECRET_KEY not in captured["client_repr"]
    assert SECRET_KEY not in repr(session.ai_plans)
    assert "甲公司" not in json.dumps(captured["catalog"], ensure_ascii=False)

    with pytest.raises(server_module.ApiError, match="确认"):
        handler._ai_execute({"plan_token": token, "confirmed": False})
    assert token in session.ai_plans, "未确认的请求不应消耗一次性凭证"

    before_ids = set(session.tables)
    executed = handler._ai_execute({"plan_token": token, "confirmed": True})

    assert executed["status"] == "completed"
    assert executed["steps_completed"] == 1
    assert executed["tables_created"] == 1
    assert len(executed["output_tables"]) == 1
    assert len(set(session.tables) - before_ids) == 1
    output_id = executed["output_tables"][0]["id"]
    assert len(session.get(output_id).frame) == 2
    assert token not in session.ai_plans
    with pytest.raises(server_module.ApiError, match="不存在|已使用|清空"):
        handler._ai_execute({"plan_token": token, "confirmed": True})


@pytest.mark.parametrize("invalidated_by", ["task_reset", "table_change", "expiry"])
def test_ai_plan_token_rejects_changed_scope_reset_or_expiry(
    ai_session,
    monkeypatch: pytest.MonkeyPatch,
    invalidated_by: str,
) -> None:
    session, handler = ai_session
    table_id = _add_orders(session)
    _install_model_plan(
        monkeypatch,
        lambda catalog: _plan_payload(catalog["tables"][0]["table_id"]),
    )
    token = _request_plan(handler, table_id)["plan_token"]

    if invalidated_by == "task_reset":
        session.reset()
        expected = "不存在|清空"
    elif invalidated_by == "table_change":
        session.get(table_id).frame["规划后新增字段"] = 1
        expected = "范围已发生变化"
    else:
        session.ai_plans[token] = replace(
            session.ai_plans[token], expires_at=time.monotonic() - 1
        )
        expected = "过期"

    with pytest.raises(server_module.ApiError, match=expected):
        handler._ai_execute({"plan_token": token, "confirmed": True})

    assert token not in session.ai_plans


@pytest.mark.parametrize("status", ["clarification", "unsupported"])
def test_non_ready_plan_never_receives_execution_token(
    ai_session,
    monkeypatch: pytest.MonkeyPatch,
    status: str,
) -> None:
    session, handler = ai_session
    table_id = _add_orders(session)
    _install_model_plan(
        monkeypatch,
        lambda catalog: _plan_payload(
            catalog["tables"][0]["table_id"], status=status
        ),
    )

    response = _request_plan(handler, table_id)

    assert response["status"] == status
    assert response["normalized_request"] == response["plan"]["summary"]
    assert isinstance(response["normalized_request"], str)
    assert response["plan_token"] is None
    assert response["expires_in_seconds"] == 0
    assert response["preview"]["executable"] is False
    assert session.ai_plans == {}


def test_provider_error_cannot_echo_api_key(
    ai_session,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session, handler = ai_session
    table_id = _add_orders(session)

    def fail_with_secret(_self: object, _prompt: str, _catalog: dict[str, Any]):
        raise DeepSeekAPIError(f"模拟上游错误，意外包含凭证 {SECRET_KEY}")

    monkeypatch.setattr(server_module.DeepSeekClient, "create_plan", fail_with_secret)

    with pytest.raises(server_module.ApiError) as captured:
        _request_plan(handler, table_id)

    assert SECRET_KEY not in str(captured.value)
    assert SECRET_KEY not in repr(captured.value)
    assert SECRET_KEY not in repr(session.ai_plans)
