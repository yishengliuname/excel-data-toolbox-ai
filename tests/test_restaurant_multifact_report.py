from __future__ import annotations

import math

import pandas as pd

from excel_data_toolbox.metric_semantics import aggregate_metric, classify_metric, grouped_metric
from excel_data_toolbox.restaurant_report import build_restaurant_diagnosis_report


def _multifact_fixture() -> tuple[list[pd.DataFrame], list[str]]:
    frames = [
        pd.DataFrame({"门店编码": ["S1", "S2"], "门店名称": ["中心店", "大学城店"], "区域": ["华东", "华东"]}),
        pd.DataFrame({"菜品编码": ["D1", "D2"], "菜品名称": ["招牌饭", "套餐"], "标准食材成本": [20, 30]}),
        pd.DataFrame(
            [
                ["2026-06-01", "O1", "S1", "美团", "D1", 1, 70, 10, 60, "已完成", "在线"],
                ["2026-06-01", "O1", "S1", "美团", "D2", 1, 50, 10, 40, "已完成", "在线"],
                ["2026-06-02", "O2", "S2", "堂食", "D1", 1, 50, 0, 50, "已完成", "现金"],
                ["2026-06-02", "O2", "S2", "堂食", "D1", 1, 50, 0, 50, "已完成", "现金"],
                ["2026-06-02", "O3", "S2", "堂食", "D1", 1, 999, 0, 999, "未完成", "现金"],
            ],
            columns=["营业日期", "订单号", "门店编码", "渠道/平台", "菜品编码", "数量", "菜品原价", "优惠分摊", "实付分摊", "订单状态", "支付方式"],
        ),
        pd.DataFrame(
            [
                ["2026-06-03", "R1", "O1", "仅退款", 30, "已退款", "出餐慢", "门店"],
                ["2026-06-03", "R2", "O2", "仅退款", 10, "退款申请完成但尚未到账", "顾客申请", "待定"],
            ],
            columns=["退款日期", "退款单号", "原订单号", "退款类型", "退款金额", "状态", "原因", "责任归属"],
        ),
        pd.DataFrame(
            [
                ["2026-06", "S1", "美团", 1000, 100, 50, 20, 30, 800],
                ["2026-06", "S2", "饿了么", 500, 50, 20, 10, 20, 400],
            ],
            columns=["结算月份", "门店编码", "平台", "订单结算基数", "平台佣金", "配送服务费", "活动补贴承担", "退款冲减", "实际到账"],
        ),
        pd.DataFrame({"原料编码": ["I1", "I2"], "原料名称": ["大米", "蔬菜"], "标准基础单价": [5.0, math.nan]}),
        pd.DataFrame({"菜品编码": ["D1", "D2"], "原料编码": ["I1", "I2"], "标准用量": [0.2, 0.3]}),
        pd.DataFrame(
            [
                ["2026-06-01", "P1", "S1", "I1", 10, 5, "已入库"],
                ["2026-06-01", "P1", "S1", "I1", 10, 5, "已入库"],
                ["2026-06-02", "P2", "S2", "I2", 10, 4, "已入库"],
                ["2026-06-02", "P2", "S2", "I2", 12, 4, "已入库"],
            ],
            columns=["入库日期", "入库单号", "门店编码", "原料编码", "采购数量", "采购单价", "状态"],
        ),
        pd.DataFrame(
            [
                ["2026-06", "S1", "I1", 100, "kg", 2, -1],
                ["2026-06", "S2", "I2", 80, "kg", 3, -2],
            ],
            columns=["月份", "门店编码", "原料编码", "期末结存数量", "单位", "报损数量", "盘点差异数量"],
        ),
        pd.DataFrame(
            [
                ["2026-06", "S1", 5, 100, 110, 10, 200, 50, 10, 5, 0],
                ["2026-06", "S2", 3, 80, 85, 5, 100, 20, 5, 0, 0],
            ],
            columns=["月份", "门店编码", "岗位人数", "排班工时", "实际工时", "加班工时", "工资成本", "临时工成本", "加班工资", "其他人工成本", "缺勤小时"],
        ),
        pd.DataFrame({"月份": ["2026-06", "2026-06"], "门店编码": ["S1", "S2"], "费用类别": ["租金", "租金"], "金额": [100, 50]}),
        pd.DataFrame(
            [["2026-06-04", "S1", "O1", 1, "投诉", "出餐太慢"], ["2026-06-04", "S2", "O9", 2, "差评", "服务差"]],
            columns=["日期", "门店编码", "订单号", "评分", "标签", "评价摘要"],
        ),
    ]
    sheets = [
        "门店主数据", "菜品主数据", "POS销售明细", "售后退款", "外卖平台结算", "原料主数据",
        "菜品BOM", "原料采购入库", "原料盘点损耗", "人工与工时", "固定运营费用", "顾客评价",
    ]
    names = [f"餐饮Golden.xlsx__{sheet}" for sheet in sheets]
    return frames, names


def test_multifact_restaurant_golden_calculations_do_not_double_count_costs() -> None:
    frames, names = _multifact_fixture()
    result = build_restaurant_diagnosis_report(frames, source_names=names)

    settlement = result.outputs["渠道与外卖分析"]
    assert settlement["活动补贴承担"].sum() == 30
    assert settlement["可比平台成本"].sum() == 250
    assert settlement["理论到账"].tolist() == [800, 400]
    assert settlement["到账勾稽差异"].abs().max() == 0
    assert settlement["可比经营贡献"].sum() == 1200

    summary = result.outputs["管理层诊断总览"].set_index("指标")["结果"]
    assert summary["营业实付"] == 150
    assert summary["已发生退款"] == 30
    assert summary["平台成本"] == 250
    assert summary["人工成本"] == 390

    dish = result.outputs["菜品盈利分析"].set_index("菜品编码")
    assert dish.loc["D1", "已退款金额"] == 18
    assert dish.loc["D2", "已退款金额"] == 12
    assert dish.loc["D1", "退款后贡献"] == 52
    assert dish.loc["D2", "退款后贡献"] == -2

    dashboard = result.outputs["经营诊断看板"]
    assert "KPI_平台到账率" in dashboard
    assert "KPI_已发生退款" in dashboard
    assert "KPI_报损金额" in dashboard
    assert "KPI_回款率" not in dashboard
    assert "KPI_库存金额" not in dashboard


def test_multifact_restaurant_audits_duplicates_labor_loss_and_review_refunds() -> None:
    frames, names = _multifact_fixture()
    result = build_restaurant_diagnosis_report(frames, source_names=names)

    audit = result.outputs["数据口径与验收"].set_index("检查项")["结果"]
    assert audit["POS精确重复发现"] == 1
    assert audit["POS精确重复删除"] == 1
    assert audit["POS精确重复剩余"] == 0
    assert "精确重复1条" in audit["采购重复与冲突"]
    assert "业务键冲突2条" in audit["采购重复与冲突"]
    assert "无法估值1条" == audit["损耗估值"]

    labor = result.outputs["人工效率分析"].set_index("门店编码")
    assert labor.loc["S1", "人工总成本"] == 265
    assert labor.loc["S2", "人工总成本"] == 125

    loss = result.outputs["原料采购与损耗"]
    loss = loss.loc[loss["原料编码"].isin(["I1", "I2"]) & loss["月份"].notna()].set_index("原料编码")
    assert loss.loc["I1", "报损数量"] == 2
    assert loss.loc["I1", "盘点差异数量"] == -1
    assert loss.loc["I1", "报损金额"] == 10
    assert pd.isna(loss.loc["I2", "报损金额"])
    assert loss.loc["I2", "估值状态"] == "无法估值：缺少标准采购单价"

    customer = result.outputs["客户评价与退款"]
    matched = customer.loc[customer["订单号"].eq("O1")].iloc[0]
    assert matched["退款匹配状态"] == "已匹配"
    assert matched["退款金额"] == 30
    assert matched["评分"] == 1


def test_restaurant_keeps_occurrence_and_order_attribution_months_separate() -> None:
    frames, names = _multifact_fixture()
    frames[3].loc[0, "退款日期"] = "2026-07-03"
    result = build_restaurant_diagnosis_report(frames, source_names=names)
    monthly = result.outputs["利润驱动分析"]

    occurrence = monthly.loc[monthly["时间口径"].eq("发生月视角")].set_index("月份")
    attributed = monthly.loc[monthly["时间口径"].eq("订单归属月视角")].set_index("月份")
    assert occurrence.loc["2026-06", "已退款金额"] == 0
    assert occurrence.loc["2026-07", "已退款金额"] == 30
    assert attributed.loc["2026-06", "已退款金额"] == 30
    assert pd.isna(attributed.loc["2026-06", "平台成本"])


def test_restaurant_purchase_deduplication_preserves_units() -> None:
    frames, names = _multifact_fixture()
    purchases = frames[7].copy()
    purchases["采购单位"] = ["kg", "kg", "kg", "kg"]
    purchases.loc[len(purchases)] = ["2026-06-03", "P3", "S1", "I1", 10, 5, "已入库", "箱"]
    purchases.loc[len(purchases)] = ["2026-06-03", "P3", "S1", "I1", 10, 5, "已入库", "kg"]
    frames[7] = purchases
    result = build_restaurant_diagnosis_report(frames, source_names=names)
    output = result.outputs["原料采购与损耗"]
    p3 = output.loc[output["入库单号"].eq("P3")]
    assert set(p3["采购单位"]) == {"箱", "kg"}
    assert p3["纳入口径"].all()


def test_unknown_numeric_percentage_and_balance_semantics_are_conservative() -> None:
    assert classify_metric("神秘数值").aggregation == "unknown"
    unknown = pd.DataFrame({"类别": ["A", "A"], "神秘数值": [10, 20]})
    value, method, _ = aggregate_metric(unknown, "神秘数值")
    assert math.isnan(value)
    assert "禁止" in method
    assert grouped_metric(unknown, "类别", "神秘数值")["神秘数值"].isna().all()

    balance = pd.DataFrame({"日期": ["2026-02-01", "2026-01-01"], "库存余额": [200, 100]})
    value, method, _ = aggregate_metric(balance, "库存余额")
    assert value == 200
    assert "日期" in method
    no_date = pd.DataFrame({"库存余额": [200, 100]})
    value, method, _ = aggregate_metric(no_date, "库存余额")
    assert math.isnan(value)
    assert "禁止" in method


def test_percentage_text_is_scaled_to_decimal_in_compact_restaurant() -> None:
    frame = pd.DataFrame(
        [["2026-06", "S1", 1000, 0, 200, 300, 100, 100, 50, 50, 1000, 200, "20%"]],
        columns=["月份", "门店", "营业额", "退款", "食材成本", "人工成本", "平台费", "租金", "水电", "营销", "净营业收入", "管理利润", "管理利润率"],
    )
    result = build_restaurant_diagnosis_report([frame], source_names=["餐饮.xlsx__门店月度经营数据"])
    detail = result.outputs["诊断底稿"]
    assert detail.loc[0, "源利润率"] == 0.2
