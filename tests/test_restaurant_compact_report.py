from __future__ import annotations

import math
from pathlib import Path
import tempfile

import pandas as pd
from openpyxl import load_workbook

from excel_data_toolbox.core import export_tables
from excel_data_toolbox.enterprise_report import build_enterprise_diagnosis_report
from excel_data_toolbox.metric_semantics import aggregate_metric, classify_metric, classify_sheet_role
from excel_data_toolbox.nl_agent import build_table_catalog, validate_plan
from excel_data_toolbox.restaurant_report import can_build_restaurant_diagnosis_report, restaurant_diagnosis_profile
from excel_data_toolbox.server import TableEntry, _enterprise_diagnosis_plan_payload


def _benchmark() -> tuple[list[pd.DataFrame], list[str]]:
    rows = [
        ["2026-04", "S01 人民广场店", 368000, 6500, 98500, 73500, 15500, 52000, 12600, 4200, 361500, 105200, 0.291010],
        ["2026-04", "S02 大学城店", 286000, 3200, 78200, 55200, 7200, 32000, 8900, 2200, 282800, 99100, 0.350424],
        ["2026-04", "S03 高新园店", 312000, 4800, 91000, 62800, 8400, 38000, 9700, 3000, 307200, 94300, 0.306966],
        ["2026-05", "S01 人民广场店", 402000, 9800, 112000, 76000, 17500, 52000, 13400, 6800, 392200, 114500, 0.291943],
        ["2026-05", "S02 大学城店", 318000, 4200, 83500, 57000, 8500, 32000, 9600, 3000, 313800, 120200, 0.383047],
        ["2026-05", "S03 高新园店", 356000, 7200, 104000, 66800, 10800, 38000, 10800, 5200, 348800, 113200, 0.324541],
        ["2026-06", "S01 人民广场店", 438000, 16800, 132500, 79800, 26500, 52000, 15100, 12800, 421200, 102500, 0.243352],
        ["2026-06", "S02 大学城店", 348000, 5600, 92500, 60300, 12000, 32000, 11200, 4800, 342400, 129600, 0.378505],
        ["2026-06", "S03 高新园店", 418000, 13800, 137000, 73500, 21200, 38000, 13800, 11800, 404200, 108900, 0.269421],
    ]
    fact = pd.DataFrame(
        rows,
        columns=["月份", "门店", "营业额", "退款", "食材成本", "人工成本", "平台费", "租金", "水电", "营销", "净营业收入", "管理利润", "管理利润率"],
    )
    summary = pd.DataFrame(
        {
            "门店": ["S01 人民广场店", "S02 大学城店", "S03 高新园店"],
            "季度营业额": [1208000, 952000, 1086000],
            "季度管理利润": [322200, 348900, 316400],
            "季度管理利润率": [322200 / 1174900, 348900 / 939000, 316400 / 1060200],
        }
    )
    return [fact, summary], ["连锁餐饮基准.xlsx__门店月度经营数据", "连锁餐饮基准.xlsx__季度汇总"]


def test_metric_semantics_do_not_average_additive_amounts_or_profit_margin() -> None:
    frames, _ = _benchmark()
    fact = frames[0]
    platform, platform_method, _ = aggregate_metric(fact, "平台费")
    margin, margin_method, _ = aggregate_metric(fact, "管理利润率")
    assert platform == 127600
    assert "求和" in platform_method
    assert math.isclose(margin, 987500 / 3174100, rel_tol=1e-12)
    assert "加权口径" in margin_method
    assert classify_metric("季度营业额").aggregation == "sum"
    assert classify_sheet_role("季度汇总", frames[1]) == "summary"


def test_compact_restaurant_schema_routes_to_diagnosis_and_excludes_summary_sheet() -> None:
    frames, names = _benchmark()
    assert can_build_restaurant_diagnosis_report(frames, names)
    assert restaurant_diagnosis_profile(frames, names) == "compact_store_period_pnl"
    result = build_enterprise_diagnosis_report(
        frames,
        source_names=names,
        user_request="全面分析三家餐饮门店的利润、成本、外卖、损耗、人工与客户风险。",
    )
    assert result.report["profile"] == "compact_store_period_pnl"
    metrics = result.outputs["管理层诊断总览"].set_index("指标")["结果"]
    assert metrics["平台费"] == 127600
    assert math.isclose(float(metrics["管理利润率"]), 987500 / 3174100, rel_tol=1e-12)
    assert metrics["最佳经营门店"] == "S02 大学城店"
    assert metrics["收入最高门店"] == "S01 人民广场店"
    sources = result.outputs["数据源与事实域"].set_index("文件/工作表")
    assert sources.loc[names[1], "纳入计算"] == "否"
    assert sources.loc[names[1], "识别角色"] == "汇总/校验表"

    monthly = result.outputs["利润驱动分析"].set_index("月份")
    assert math.isclose(monthly.loc["2026-04", "管理利润率"], 298600 / 951500, rel_tol=1e-12)
    assert math.isclose(monthly.loc["2026-05", "管理利润率"], 347900 / 1054800, rel_tol=1e-12)
    assert math.isclose(monthly.loc["2026-06", "管理利润率"], 341000 / 1167800, rel_tol=1e-12)
    risks = result.outputs["风险行动计划"]
    assert risks["风险事项"].str.contains("收入增长未同步转化").any()
    assert risks["风险事项"].str.contains("高营收门店利润转化").any()


def test_compact_restaurant_plan_is_not_rejected_by_legacy_sheet_count_gate() -> None:
    frames, names = _benchmark()
    entries = [TableEntry(str(index), name, frame, "导入文件", True) for index, (name, frame) in enumerate(zip(names, frames))]
    payload = _enterprise_diagnosis_plan_payload(entries, "全面分析餐饮门店利润、成本、外卖、人工和客户风险")
    catalog = build_table_catalog(
        {entry.id: entry.frame for entry in entries},
        display_names={entry.id: entry.name for entry in entries},
    )
    plan = validate_plan(payload, catalog)
    assert plan.status == "ready"
    assert plan.executable
    assert plan.steps[0].operation == "enterprise_diagnosis_report"


def test_compact_restaurant_export_has_specific_dashboard_and_four_named_charts() -> None:
    frames, names = _benchmark()
    result = build_enterprise_diagnosis_report(frames, source_names=names)
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "连锁餐饮基准_经营诊断报告.xlsx"
        export_tables(result.outputs, path, include_log=False)
        workbook = load_workbook(path, read_only=False, data_only=False)
        try:
            assert workbook["管理层诊断总览"]["A1"].value == "连锁餐饮经营诊断驾驶舱"
            charts = workbook["经营诊断看板"]._charts
            assert len(charts) == 4
            titles = [chart.title.tx.rich.p[0].r[0].t for chart in charts]
            assert titles == [
                "门店营业额与管理利润（元）",
                "门店管理利润率（加权口径）",
                "月度营业额与管理利润趋势（元）",
                "月度管理利润率趋势（加权口径）",
            ]
        finally:
            workbook.close()
