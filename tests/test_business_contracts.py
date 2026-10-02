"""Synthetic adversarial cases: no new customer or industry-specific template."""
import pandas as pd
import pytest

from excel_data_toolbox.adaptive_report import build_adaptive_analysis_report
from excel_data_toolbox.analysis_compiler import compile_analysis
from excel_data_toolbox.business_contracts import (
    metric_contract, evaluate_metric, relationship_graph, candidate_grain,
)


def sales():
    return pd.DataFrame({"订单号": ["A", "B", "C", "D"],
        "日期": ["2026-01-01", "2026-01-02", "2026-02-01", "2026-02-02"],
        "地区": ["东", "西", "东", "西"], "销售额": [100, 200, 300, 400]})


@pytest.mark.parametrize("bad", [None, "not a number", float("inf"), float("-inf")])
def test_incomplete_amount_does_not_become_zero_or_partial_total(bad):
    frame = sales()
    frame["销售额"] = frame["销售额"].astype(object)
    frame.loc[1, "销售额"] = bad
    contract = metric_contract(0, "orders", frame, "销售额")
    assert contract.status == "unavailable"
    assert pd.isna(evaluate_metric(frame, contract)[0])
    report = build_adaptive_analysis_report([frame], source_names=["orders"])
    row = report.outputs["管理层通用总览"].set_index("指标").loc["核心指标：销售额"]
    assert row["结果"] == "不可计算"
    assert "指标不可计算" in report.outputs["异常数据"]["异常类型"].tolist()


def test_duplicates_are_preserved_and_not_silently_aggregated():
    raw = pd.concat([sales(), sales().iloc[[0]]], ignore_index=True)
    result = build_adaptive_analysis_report([raw], source_names=["orders"])
    assert len(result.outputs["主数据分析"]) == len(raw)
    assert len(result.outputs["事实域明细"]) == len(raw)
    assert result.outputs["语义契约"].iloc[0]["状态"] == "unavailable"
    assert result.outputs["异常数据"]["异常类型"].eq("疑似重复（保留）").sum() == 2


@pytest.mark.parametrize("header", ["采购单价", "平均销售额", "周转天数", "每人销售额"])
def test_nonadditive_unit_metrics_are_not_summed(header):
    frame = sales().rename(columns={"销售额": header})
    assert metric_contract(0, "fact", frame, header).status == "unavailable"


@pytest.mark.parametrize("unit_column,units", [
    ("币种", ["CNY", "USD", "CNY", "CNY"]),
    ("采购单位", ["箱", "kg", "kg", "kg"]),
    ("数量单位", ["kg", None, "kg", "kg"]),
])
def test_mixed_or_missing_units_require_conversion_rules(unit_column, units):
    frame = sales().assign(**{unit_column: units})
    assert metric_contract(0, "fact", frame, "销售额").status == "unavailable"


def test_multiple_time_axes_never_select_the_first_column_by_accident():
    frame = sales().assign(退款日期=["2026-03-01"] * 4)
    plan = compile_analysis([frame], source_names=["orders"], user_request="分析月度趋势")
    assert not any(a.kind == "trend" for a in plan.analyses)
    result = build_adaptive_analysis_report([frame], source_names=["orders"], user_request="月度趋势")
    assert result.outputs["时间趋势"].empty
    assert result.outputs["问题与证据"].set_index("分析问题").loc["trend", "证据状态"] == "unavailable"


def test_balance_sums_entities_at_one_latest_snapshot_not_one_last_row():
    frame = pd.DataFrame({"商品编码": ["A", "B", "A", "B"],
        "日期": ["2026-01-01", "2026-01-01", "2026-02-01", "2026-02-01"],
        "库存余额": [10, 20, 30, 40]})
    contract = metric_contract(0, "snapshots", frame, "库存余额")
    assert evaluate_metric(frame, contract)[0] == 70
    missing_latest = frame.iloc[:-1]
    assert metric_contract(0, "snapshots", missing_latest, "库存余额").status == "unavailable"


def test_weighted_ratio_needs_the_same_complete_population_and_unit_scale():
    frame = sales().assign(利润=[30, 60, 90, 40], 利润率=[.3, .3, .3, .1])
    contract = metric_contract(0, "fact", frame, "利润率")
    assert evaluate_metric(frame, contract)[0] == pytest.approx(.22)
    frame.loc[1, "利润"] = None
    assert metric_contract(0, "fact", frame, "利润率").status == "unavailable"
    frame["利润"] = [30, 60, 90, 40]
    scaled = frame.rename(columns={"利润": "利润(万元)", "销售额": "销售额(元)"})
    assert metric_contract(0, "fact", scaled, "利润率").status == "unavailable"


def test_known_money_format_and_scale_do_not_corrupt_values():
    frame = sales().rename(columns={"销售额": "销售额(万元)"})
    frame["销售额(万元)"] = ["1,000", "200", "300", "400"]
    contract = metric_contract(0, "fact", frame, "销售额(万元)")
    assert contract.unit == "万元"
    assert evaluate_metric(frame, contract)[0] == 1900


def test_same_column_name_is_not_a_relationship_or_a_profit_formula():
    income = sales()
    expenses = pd.DataFrame({"日期": ["2026-01-01", "2026-02-01"], "费用金额": [20, 30]})
    result = build_adaptive_analysis_report([income, expenses], source_names=["orders", "expenses"], user_request="公司赚不赚钱，关联分析利润")
    assert result.outputs["表关系建议"].empty  # shared date is not a business key
    assert result.outputs["问题与证据"].set_index("分析问题").loc["profitability", "证据状态"] == "unavailable"
    assert not result.outputs["管理层通用总览"]["指标"].str.contains("净利润").any()


def test_many_to_many_keys_and_bidirectional_coverage_are_explicit():
    left = pd.DataFrame({"订单号": ["A", "A", "B", "C", "C"], "销售额": [1] * 5})
    right = pd.DataFrame({"订单号": ["A", "A", "X"], "退款金额": [1] * 3})
    edge = relationship_graph([left, right], ["sales", "refunds"], ["fact", "fact"])[0]
    assert edge.cardinality == "many_to_many" and edge.status == "blocked"
    assert edge.left_row_coverage == .4
    assert edge.right_row_coverage == pytest.approx(2/3)
    assert edge.left_key_coverage == pytest.approx(1/3)


def test_secondary_fact_can_drive_ranking_and_trend_without_primary_metrics():
    unknown = sales().rename(columns={"销售额": "未定义测量"})
    unknown["说明"] = ["alpha"] * 4  # make this the largest display anchor
    secondary = sales().rename(columns={"订单号": "退款单号", "销售额": "退款金额"})
    plan = compile_analysis([unknown, secondary], source_names=["unknown", "refunds"])
    assert plan.primary_index == 0
    assert any(a.kind == "trend" and a.table_index == 1 for a in plan.analyses)
    assert not any(a.kind == "trend" and a.table_index == 0 for a in plan.analyses)
    result = build_adaptive_analysis_report([unknown, secondary], source_names=["unknown", "refunds"])
    assert set(result.outputs["时间趋势"]["来源事实表"]) == {"refunds"}
    assert result.outputs["自适应图表看板"]["趋势来源"].eq("refunds").all()


def test_top_n_share_uses_full_population_and_chart_keeps_other_categories():
    frame = pd.DataFrame({"订单号": list("ABCDE"), "日期": ["2026-01-01"] * 5,
        "产品": list("ABCDE"), "销售额": [50, 40, 30, 20, 10]})
    result = build_adaptive_analysis_report([frame], source_names=["orders"], top_n=3)
    ranking = result.outputs["分类排名"]
    assert ranking["占比"].sum() == pytest.approx(.8)
    dashboard = result.outputs["自适应图表看板"]
    assert dashboard.loc[dashboard["结构分类"].eq("其他类别"), "结构指标值"].iloc[0] == pytest.approx(30)
    assert dashboard["结构指标值"].sum() == pytest.approx(150)


def test_summary_only_input_gets_structure_audit_not_business_kpis():
    frame = pd.DataFrame({"指标": ["销售额", "利润"], "结果": [100, 20]})
    result = build_adaptive_analysis_report([frame], source_names=["历史管理层总览"])
    assert result.report["fact_count"] == 0
    assert result.outputs["语义契约"].empty
    assert not result.outputs["管理层通用总览"]["指标"].str.startswith("核心指标：").any()


def test_physical_key_is_validated_and_text_is_not_truncated():
    frame = sales()
    frame["备注"] = ["长描述" * 200] * 4
    original = frame.copy(deep=True)
    result = build_adaptive_analysis_report([frame], source_names=["orders"])
    assert len(result.outputs["主数据分析"].iloc[0]["备注"]) == 600
    assert candidate_grain(frame) == ("订单号",)
    pd.testing.assert_frame_equal(original, frame)
    duplicate = pd.concat([frame, frame.iloc[[0]]])
    assert candidate_grain(duplicate) == ()


def test_stateful_rows_require_explicit_population_rules():
    frame = sales().assign(状态=["已完成", "未完成", "支付失败", "已完成"])
    contract = metric_contract(0, "orders", frame, "销售额")
    assert contract.status == "unavailable"
    assert "状态" in str(contract.issues)
