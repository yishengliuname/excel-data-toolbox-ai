from __future__ import annotations

import base64
import json
from pathlib import Path
import zipfile

import pandas as pd
import pytest

import excel_data_toolbox.server as server_module
from excel_data_toolbox.nl_agent import PlanValidationError
from excel_data_toolbox.power_bi_automation import (
    FabricPublisher,
    HttpResult,
    POWER_BI_ENV_KEYS,
    PowerBIConfig,
    build_power_bi_bundle,
)
from excel_data_toolbox.server import AppSession, ToolboxHandler


def _sales_frame() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "订单编号": ["A001", "A002", "A003", "A003", "A004"],
            "日期": pd.to_datetime(["2026-01-02", "2026-01-15", "2026-02-01", "2026-02-01", "2026-03-08"]),
            "客户": ["甲公司", "乙公司", "甲公司", "甲公司", None],
            "地区": ["华东", "华南", "华东", "华东", "华北"],
            "渠道": ["线上", "门店", "直销", "直销", "线上"],
            "产品": ["办公椅", "显示器", "办公椅", "办公椅", "升降桌"],
            "订单金额": [1000.0, 2400.0, 1800.0, 1800.0, None],
            "数量": [2, 2, 3, 3, 1],
            "成本": [600.0, 1800.0, 900.0, 900.0, 1500.0],
        }
    )


def test_bundle_builds_star_model_pbir_and_passes_all_checks(tmp_path: Path) -> None:
    bundle = build_power_bi_bundle(
        _sales_frame(),
        tmp_path,
        task_id="CASE-001",
        source_name="虚构销售演示数据",
    )

    assert bundle["validation"]["passed"] is True
    assert bundle["validation"]["summary"]["failed"] == 0
    assert bundle["model_spec"]["tables"]["FactSales"]["rows"] == 4
    assert set(bundle["model_spec"]["tables"]) == {
        "FactSales", "DimDate", "DimCustomer", "DimRegion", "DimChannel", "DimProduct"
    }
    assert len(bundle["model_spec"]["relationships"]) == 5
    assert len(bundle["model_spec"]["measures"]) == 10
    assert len(bundle["report_spec"]["pages"]) == 3
    assert sum(page["visual_count"] for page in bundle["report_spec"]["pages"]) == 13
    assert Path(bundle["zip_path"]).is_file()

    with zipfile.ZipFile(bundle["zip_path"]) as archive:
        names = archive.namelist()
        assert any(name.endswith("SalesAutomation.SemanticModel/model.bim") for name in names)
        assert any(name.endswith("SalesAutomation.Report/definition/pages/pages.json") for name in names)
        assert any(name.endswith("PowerQuery/FactSales.m") for name in names)
        assert any(name.endswith("validation_report.json") for name in names)
        combined = b"".join(archive.read(name) for name in names if not name.endswith(".csv"))
        assert b"unit-test-secret-never-log" not in combined.lower()
        assert b"sk-" not in combined.lower()


def test_fabric_publisher_creates_model_then_bound_report_and_verifies(tmp_path: Path) -> None:
    bundle = build_power_bi_bundle(
        _sales_frame(), tmp_path, task_id="CASE-002", source_name="测试销售数据"
    )
    config = PowerBIConfig(
        tenant_id="11111111-1111-4111-8111-111111111111",
        client_id="22222222-2222-4222-8222-222222222222",
        client_secret="unit-test-secret-never-log",
        workspace_id="33333333-3333-4333-8333-333333333333",
    )
    semantic_id = "44444444-4444-4444-8444-444444444444"
    report_id = "55555555-5555-4555-8555-555555555555"
    calls: list[dict[str, object]] = []

    def transport(method: str, url: str, headers: object, body: bytes | None, timeout: int) -> HttpResult:
        calls.append({"method": method, "url": url, "headers": headers, "body": body, "timeout": timeout})
        if "login.microsoftonline.com" in url:
            assert body and b"client_secret=unit-test-secret-never-log" in body
            return HttpResult(200, {}, {"access_token": "fake-access-token"})
        if method == "POST" and url.endswith("/semanticModels"):
            return HttpResult(201, {}, {"id": semantic_id, "type": "SemanticModel"})
        if method == "POST" and url.endswith("/reports"):
            payload = json.loads((body or b"{}").decode("utf-8"))
            pbir_part = next(item for item in payload["definition"]["parts"] if item["path"] == "definition.pbir")
            pbir = json.loads(base64.b64decode(pbir_part["payload"]).decode("utf-8"))
            assert pbir["datasetReference"]["byConnection"]["connectionString"] == f"semanticmodelid={semantic_id}"
            assert "byPath" not in pbir["datasetReference"]
            return HttpResult(201, {}, {"id": report_id, "type": "Report"})
        if method == "GET" and url.endswith(semantic_id):
            return HttpResult(200, {}, {"id": semantic_id, "type": "SemanticModel"})
        if method == "GET" and url.endswith(report_id):
            return HttpResult(200, {}, {"id": report_id, "type": "Report"})
        raise AssertionError(f"unexpected request: {method} {url}")

    result = FabricPublisher(config, transport=transport).publish(
        model_root=bundle["model_root"],
        report_root=bundle["report_root"],
        display_name="自动化销售案例",
    )

    assert result["status"] == "published"
    assert result["verified"] is True
    assert result["semantic_model_id"] == semantic_id
    assert result["report_id"] == report_id
    assert [call["method"] for call in calls] == ["POST", "POST", "POST", "GET", "GET"]
    public_trace = json.dumps(
        [{"method": call["method"], "url": call["url"]} for call in calls], ensure_ascii=False
    )
    assert config.client_secret not in public_trace


def test_server_power_bi_case_falls_back_compiles_and_registers_download(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    session = AppSession()
    monkeypatch.setattr(server_module, "SESSION", session)
    for name in POWER_BI_ENV_KEYS:
        monkeypatch.delenv(name, raising=False)
    table_id = session.add_table("虚构销售演示数据", _sales_frame(), source="测试", original=True)

    def invalid_model_shape(*_args: object, **_kwargs: object):
        raise PlanValidationError("模拟 DeepSeek 字段结构漂移")

    monkeypatch.setattr(server_module.DeepSeekClient, "create_engineering_brief", invalid_model_shape)
    handler = object.__new__(ToolboxHandler)
    try:
        result = handler._ai_engineering(
            {
                "category": "power_bi",
                "prompt": "加载示例数据，为当前数据设计 Power BI 星型模型、核心 DAX 指标、页面布局和验收清单，并自动发布。",
                "api_key": "sk-unit-test-secret-123456789",
                "model": "deepseek-v4-flash",
                "table_ids": [table_id],
            }
        )
        assert result["execution"] == "package_built_and_validated"
        assert result["automation"]["status"] == "credentials_required"
        assert result["automation"]["validation"]["passed"] is True
        assert result["automation"]["download_url"].startswith("/download/")
        assert len(session.downloads) == 1
        assert next(iter(session.downloads.values())).exists()
        assert "自动切换本地稳定方案" in result["message"]
    finally:
        session.close()
