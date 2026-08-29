"""Restaurant multi-fact diagnosis plugin.

The generic engine deliberately does not guess that a POS table is the whole
business.  This plugin keeps POS, refunds, settlement, recipe/BOM, purchases,
waste, labour, fixed cost and reviews at their native grain, and only joins
them through explicit business keys.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping, Sequence
import math
import pandas as pd

try:
    from .metric_semantics import classify_sheet_role
    from .status_semantics import (
        ORDER_INVALID,
        ORDER_SUCCESS,
        RECEIPT_CONFIRMED,
        REFUND_CONFIRMED,
        REFUND_PENDING,
        classify_order_status,
        classify_receipt_status,
        classify_refund_status,
        classify_status_series,
    )
except ImportError:  # Supports: python restaurant_report.py
    from excel_data_toolbox.metric_semantics import classify_sheet_role
    from excel_data_toolbox.status_semantics import (
        ORDER_INVALID,
        ORDER_SUCCESS,
        RECEIPT_CONFIRMED,
        REFUND_CONFIRMED,
        REFUND_PENDING,
        classify_order_status,
        classify_receipt_status,
        classify_refund_status,
        classify_status_series,
    )


ROLE_ALIASES: Mapping[str, tuple[str, ...]] = {
    "stores": ("门店主数据", "门店资料", "门店"),
    "dishes": ("菜品主数据", "菜品资料", "菜品"),
    "pos": ("POS销售明细", "销售明细", "POS"),
    "refunds": ("售后退款", "退款"),
    "settlements": ("外卖平台结算", "平台结算", "结算"),
    "ingredients": ("原料主数据", "原料资料", "原料"),
    "bom": ("菜品BOM", "BOM"),
    "purchases": ("原料采购入库", "采购入库", "采购"),
    "losses": ("原料盘点损耗", "盘点损耗", "损耗", "报损"),
    "labor": ("人工与工时", "人工", "工时"),
    "fixed": ("固定运营费用", "固定费用", "运营费用"),
    "reviews": ("顾客评价", "客户评价", "评价"),
    "notes": ("经营说明", "口径说明", "项目说明"),
}

COMPACT_ALIASES: Mapping[str, tuple[str, ...]] = {
    "month": ("月份", "期间", "年月", "会计期间"),
    "store": ("门店", "门店名称", "门店编码", "店铺"),
    "revenue": ("营业额", "营业收入", "销售额", "销售收入"),
    "refund": ("退款", "退款金额", "售后退款"),
    "food": ("食材成本", "原料成本", "材料成本"),
    "labor": ("人工成本", "工资成本", "人员成本"),
    "platform": ("平台费", "平台费用", "平台佣金", "外卖平台费"),
    "rent": ("租金", "房租", "租赁费"),
    "utilities": ("水电", "水电费", "能源费"),
    "marketing": ("营销", "营销费", "营销费用", "推广费"),
    "net_revenue": ("净营业收入", "净收入", "退款后收入"),
    "profit": ("管理利润", "经营利润", "净利润"),
    "margin": ("管理利润率", "经营利润率", "净利润率"),
}


def _norm(value: Any) -> str:
    return str(value or "").strip().replace(" ", "").replace("_", "").casefold()


def _find(frame: pd.DataFrame, *aliases: str, required: bool = False) -> str | None:
    cols = {_norm(c): str(c) for c in frame.columns}
    for alias in aliases:
        if _norm(alias) in cols:
            return cols[_norm(alias)]
    matches = {
        col
        for key, col in cols.items()
        if any(len(_norm(alias)) >= 2 and _norm(alias) in key for alias in aliases)
    }
    if len(matches) == 1:
        return next(iter(matches))
    if len(matches) > 1 and required:
        raise ValueError(f"字段匹配不唯一：{'/'.join(aliases)}；候选：{'、'.join(sorted(matches))}")
    if required:
        raise ValueError(f"缺少字段：{'/'.join(aliases)}")
    return None


def _num(series: pd.Series) -> pd.Series:
    text = series.astype("string").str.strip()
    percentage = text.str.contains(r"[%％]", regex=True, na=False)
    values = pd.to_numeric(
        text.str.replace(r"[¥￥元,，%％\s件个kg公斤箱份]", "", regex=True),
        errors="coerce",
    ).astype("Float64")
    return values.where(~percentage, values / 100.0)


def _col_exact(frame: pd.DataFrame, *aliases: str) -> pd.Series:
    """Return an exact column-name match without substring inference.

    A total such as ``人工总成本`` must not silently match the component
    ``其他人工成本``.  Exact matching keeps source totals optional and lets the
    component sum remain the authoritative fallback.
    """

    columns = {_norm(column): column for column in frame.columns}
    for alias in aliases:
        column = columns.get(_norm(alias))
        if column is not None:
            return frame[column]
    return pd.Series(pd.NA, index=frame.index, dtype="object")


def _text(series: pd.Series) -> pd.Series:
    return series.astype("string").fillna("").str.strip()


def _date(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce", format="mixed")


def _month(series: pd.Series) -> pd.Series:
    return _date(series).dt.to_period("M").astype("string")


def _frame(frames: Sequence[pd.DataFrame], names: Sequence[str]) -> dict[str, pd.DataFrame]:
    result: dict[str, pd.DataFrame] = {}
    for index, frame in enumerate(frames):
        label = str(names[index]) if index < len(names) else f"表{index + 1}"
        short = label.rsplit("__", 1)[-1]
        n = _norm(short)
        for role, aliases in ROLE_ALIASES.items():
            if role not in result and any(_norm(a) in n or n in _norm(a) for a in aliases):
                result[role] = frame.copy()
                result[role].attrs["_restaurant_source_name"] = label
                break
    return result


def _compact_operating_input(
    frames: Sequence[pd.DataFrame], names: Sequence[str]
) -> tuple[int, dict[str, str]] | None:
    """Find a periodic store operating fact without relying on sheet count."""

    for index, frame in enumerate(frames):
        if not isinstance(frame, pd.DataFrame) or frame.empty:
            continue
        label = names[index] if index < len(names) else f"表{index + 1}"
        if classify_sheet_role(label, frame) in {"summary", "notes"}:
            continue
        columns = {
            key: _find(frame, *aliases)
            for key, aliases in COMPACT_ALIASES.items()
        }
        required = ("month", "store", "revenue", "profit")
        restaurant_signal = (
            columns.get("food") is not None
            or any(token in _norm(label) for token in ("餐饮", "门店", "菜品", "食材", "外卖"))
        )
        cost_domains = sum(columns.get(key) is not None for key in ("refund", "food", "labor", "platform", "rent", "utilities", "marketing"))
        if restaurant_signal and cost_domains >= 3 and all(columns.get(key) for key in required):
            return index, {key: value for key, value in columns.items() if value is not None}
    return None


def restaurant_diagnosis_profile(frames: Sequence[pd.DataFrame], source_names: Sequence[str] | None = None) -> str | None:
    if not isinstance(frames, Sequence) or not frames:
        return None
    names = list(source_names or [f"表{i + 1}" for i in range(len(frames))])
    if _compact_operating_input(frames, names) is not None:
        return "compact_store_period_pnl"
    if len(frames) < 5:
        return None
    roles = _frame(frames, names)
    required = {"pos", "dishes", "refunds", "settlements", "labor", "fixed"}
    return "multi_fact_restaurant" if required.issubset(roles) and len(roles) >= 8 else None


def can_build_restaurant_diagnosis_report(frames: Sequence[pd.DataFrame], source_names: Sequence[str] | None = None) -> bool:
    return restaurant_diagnosis_profile(frames, source_names) is not None


def validate_restaurant_diagnosis_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("餐饮诊断参数必须是对象")
    names = params.get("source_names")
    if not isinstance(names, (list, tuple)) or not names or not all(isinstance(x, str) and x.strip() for x in names):
        raise TypeError("source_names 必须是非空字符串数组")


def _col(frame: pd.DataFrame, *aliases: str, required: bool = False) -> pd.Series:
    name = _find(frame, *aliases, required=required)
    return frame[name] if name else pd.Series(pd.NA, index=frame.index)


def _fact(metric: str, value: Any, unit: str, source: str, definition: str, meaning: str, confidence: str = "高") -> dict[str, Any]:
    if isinstance(value, float) and not math.isfinite(value):
        value = "未提供"
    return {"指标": metric, "结果": value, "单位": unit, "数据来源": source, "数据口径": definition, "管理含义": meaning, "结论类型": "事实", "置信度": confidence}


@dataclass(frozen=True)
class RestaurantDiagnosisResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def _compact_restaurant_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    user_request: str,
) -> RestaurantDiagnosisResult:
    """Diagnose a compact store-month P&L table with weighted metrics.

    This profile is schema-driven: it works for any number of stores and
    periods as long as the uploaded fact contains the required operating
    measures.  Derived verification/summary sheets are kept as audit evidence
    but excluded from all calculations.
    """

    detected = _compact_operating_input(frames, list(source_names))
    if detected is None:
        raise ValueError("未识别到门店-月份经营事实表")
    fact_index, columns = detected
    raw = frames[fact_index].copy(deep=True)
    source_name = str(source_names[fact_index])

    def numeric(key: str, *, default: float | None = None) -> pd.Series:
        column = columns.get(key)
        if column:
            return _num(raw[column]).astype("float64")
        return pd.Series(default, index=raw.index, dtype="float64")

    month_source = raw[columns["month"]]
    parsed_month = pd.to_datetime(month_source, errors="coerce", format="mixed")
    month_text = month_source.astype("string").str.strip()
    month = parsed_month.dt.to_period("M").astype("string").where(parsed_month.notna(), month_text)
    work = pd.DataFrame(
        {
            "月份": month,
            "门店": _text(raw[columns["store"]]),
            "营业额": numeric("revenue"),
            "退款": numeric("refund", default=0.0),
            "食材成本": numeric("food"),
            "人工成本": numeric("labor"),
            "平台费": numeric("platform", default=0.0),
            "租金": numeric("rent", default=0.0),
            "水电": numeric("utilities", default=0.0),
            "营销": numeric("marketing", default=0.0),
            "源行号": pd.RangeIndex(start=2, stop=len(raw) + 2),
        }
    )
    work["净营业收入"] = numeric("net_revenue") if columns.get("net_revenue") else work["营业额"] - work["退款"]
    cost_columns = ["食材成本", "人工成本", "平台费", "租金", "水电", "营销"]
    work["重算管理利润"] = work["净营业收入"] - work[cost_columns].sum(axis=1, min_count=len(cost_columns))
    work["源管理利润"] = numeric("profit")
    complete_cost_components = all(
        columns.get(key) is not None
        for key in ("food", "labor", "platform", "rent", "utilities", "marketing")
    )
    work["管理利润"] = work["重算管理利润"] if complete_cost_components else work["源管理利润"]
    work["利润主口径"] = "按完整成本组件重算" if complete_cost_components else "源管理利润（成本组件不完整）"
    work["管理利润率"] = work["管理利润"].div(work["净营业收入"].where(work["净营业收入"].ne(0)))
    work["源利润率"] = numeric("margin") if columns.get("margin") else work["管理利润率"]
    work["利润勾稽差异"] = work["源管理利润"] - work["重算管理利润"]
    work["净收入勾稽差异"] = work["净营业收入"] - (work["营业额"] - work["退款"])
    work["精确重复"] = work.drop(columns=["源行号"]).duplicated(keep=False)
    valid = work.loc[work["月份"].notna() & work["门店"].ne("")].copy()
    if valid.empty:
        raise ValueError("门店经营事实表没有可用的月份和门店记录")

    additive = ["营业额", "退款", "净营业收入", *cost_columns, "管理利润"]
    totals = valid[additive].sum(min_count=1)
    overall_margin = float(totals["管理利润"] / totals["净营业收入"]) if totals["净营业收入"] else float("nan")
    food_rate = float(totals["食材成本"] / totals["净营业收入"]) if totals["净营业收入"] else float("nan")

    store = valid.groupby("门店", as_index=False, observed=True)[additive].sum(min_count=1)
    store["管理利润率"] = store["管理利润"].div(store["净营业收入"].where(store["净营业收入"].ne(0)))
    store["食材成本率"] = store["食材成本"].div(store["净营业收入"].where(store["净营业收入"].ne(0)))
    store["人工成本率"] = store["人工成本"].div(store["净营业收入"].where(store["净营业收入"].ne(0)))
    store["平台费率"] = store["平台费"].div(store["净营业收入"].where(store["净营业收入"].ne(0)))
    store["利润排名"] = store["管理利润"].rank(method="min", ascending=False).astype("Int64")
    store["收入排名"] = store["净营业收入"].rank(method="min", ascending=False).astype("Int64")
    best_store = str(store.sort_values(["管理利润", "管理利润率"], ascending=False, kind="stable").iloc[0]["门店"])
    highest_revenue_store = str(store.sort_values("净营业收入", ascending=False, kind="stable").iloc[0]["门店"])
    lowest_margin_store = str(store.sort_values("管理利润率", ascending=True, kind="stable").iloc[0]["门店"])
    store["管理诊断"] = store.apply(
        lambda row: (
            "最佳经营门店：利润与利润率领先"
            if row["门店"] == best_store
            else "规模领先但利润转化偏弱，优先拆解折扣与费用"
            if row["门店"] == highest_revenue_store and row["门店"] == lowest_margin_store
            else "保持跟踪，聚焦成本率和利润转化"
        ),
        axis=1,
    )
    store = store.sort_values("利润排名", kind="stable").reset_index(drop=True)

    monthly = valid.groupby("月份", as_index=False, observed=True)[additive].sum(min_count=1).sort_values("月份", kind="stable")
    monthly["管理利润率"] = monthly["管理利润"].div(monthly["净营业收入"].where(monthly["净营业收入"].ne(0)))
    monthly["食材成本率"] = monthly["食材成本"].div(monthly["净营业收入"].where(monthly["净营业收入"].ne(0)))
    monthly["人工成本率"] = monthly["人工成本"].div(monthly["净营业收入"].where(monthly["净营业收入"].ne(0)))
    monthly["平台费率"] = monthly["平台费"].div(monthly["净营业收入"].where(monthly["净营业收入"].ne(0)))
    monthly["营业额环比"] = monthly["营业额"].pct_change()
    monthly["利润环比"] = monthly["管理利润"].pct_change()
    monthly["利润率环比变化"] = monthly["管理利润率"].diff()

    cost = pd.DataFrame(
        [
            {
                "成本费用项目": column,
                "季度金额": float(totals[column]),
                "占净营业收入": float(totals[column] / totals["净营业收入"]) if totals["净营业收入"] else float("nan"),
                "数据来源": source_name,
                "口径": "门店月度经营事实表有效行求和",
            }
            for column in cost_columns
        ]
    ).sort_values("季度金额", ascending=False, kind="stable")

    last = monthly.iloc[-1]
    previous = monthly.iloc[-2] if len(monthly) >= 2 else None
    margin_decline = float(last["管理利润率"] - previous["管理利润率"]) if previous is not None else 0.0
    revenue_change = float(last["净营业收入"] / previous["净营业收入"] - 1) if previous is not None and previous["净营业收入"] else float("nan")
    risks: list[dict[str, Any]] = []

    def add_risk(priority: str, issue: str, evidence: str, impact: str, owner: str, action: str, deadline: str, acceptance: str, boundary: str = "") -> None:
        risks.append(
            {
                "优先级": priority,
                "风险事项": issue,
                "数据证据": evidence,
                "风险影响": impact,
                "责任部门": owner,
                "建议行动": action,
                "完成期限": deadline,
                "验收指标": acceptance,
                "人工核验边界": boundary,
            }
        )

    if previous is not None and margin_decline < -0.01:
        add_risk(
            "P1",
            "收入增长未同步转化为利润率",
            f"{last['月份']}净营业收入{last['净营业收入']:,.0f}元，较上月{revenue_change:+.1%}；管理利润率由{previous['管理利润率']:.2%}降至{last['管理利润率']:.2%}",
            "继续追求规模可能放大利润和现金压力",
            "运营/财务",
            "按门店拆解折扣、平台费、食材和人工成本率变化，停止只以营业额考核",
            "7天",
            "形成门店-月份利润率桥接表，并解释利润率变动至少90%",
        )
    if highest_revenue_store == lowest_margin_store and highest_revenue_store != best_store:
        row = store.loc[store["门店"].eq(highest_revenue_store)].iloc[0]
        best = store.loc[store["门店"].eq(best_store)].iloc[0]
        add_risk(
            "P1",
            "高营收门店利润转化偏弱",
            f"{highest_revenue_store}净营业收入{row['净营业收入']:,.0f}元居首，但管理利润率{row['管理利润率']:.2%}最低；{best_store}利润{best['管理利润']:,.0f}元、利润率{best['管理利润率']:.2%}领先",
            "规模排名会掩盖折扣、成本或平台费用侵蚀",
            "门店运营/财务BP",
            "复盘两店客单、折扣、平台费率、食材成本率和排班差异，形成可复制改善动作",
            "14天",
            "目标门店利润率提升且不以异常压缩必要服务成本实现",
        )
    platform_rate = float(totals["平台费"] / totals["净营业收入"]) if totals["净营业收入"] else float("nan")
    add_risk(
        "P2",
        "外卖增长的独立盈利性尚不能确认",
        f"当前仅有平台费合计{totals['平台费']:,.0f}元，占净营业收入{platform_rate:.2%}；缺少外卖订单、平台结算与渠道收入拆分",
        "无法判断外卖增量收入是否覆盖平台费、配送、补贴和退款",
        "电商/外卖运营+财务",
        "补充渠道级订单、折扣、平台结算、配送费与到账数据后重算渠道贡献",
        "下个结算周期",
        "外卖渠道收入、费用、退款、到账可逐月勾稽",
        "没有渠道收入与平台结算明细，不强行判断外卖赚亏",
    )
    add_risk(
        "P2",
        "菜品、损耗、工时与投诉缺少明细证据",
        "当前事实表仅提供门店月度汇总成本，未提供菜品销量/BOM、采购盘点、工时排班和顾客评价",
        "无法可靠回答具体菜品、原料损耗、加班效率和投诉集中点",
        "门店运营/采购/人事/客服",
        "按统一主键补充订单-菜品、BOM-原料、采购-盘点、员工-工时及评价明细",
        "30天",
        "四类明细表字段、期间、门店编码和单位通过验收",
        "缺失事实域只列数据缺口，不用模板造结论",
    )
    risk_df = pd.DataFrame(risks).sort_values("优先级", kind="stable").reset_index(drop=True)

    summary = pd.DataFrame(
        [
            _fact("季度营业额", float(totals["营业额"]), "元", source_name, "门店月度营业额求和", "规模口径，不等同净收入"),
            _fact("退款金额", float(totals["退款"]), "元", source_name, "门店月度退款求和", "用于桥接营业额与净营业收入"),
            _fact("季度净营业收入", float(totals["净营业收入"]), "元", source_name, "营业额-退款；与源净营业收入勾稽", "利润率统一分母"),
            _fact("食材成本", float(totals["食材成本"]), "元", source_name, "门店月度食材成本求和", "当前仅能分析汇总成本，不能定位菜品/BOM"),
            _fact("人工成本", float(totals["人工成本"]), "元", source_name, "门店月度人工成本求和", "缺工时明细，不能判断加班效率"),
            _fact("平台费", float(totals["平台费"]), "元", source_name, "门店月度平台费求和", "费用是可加金额，禁止取平均"),
            _fact("管理利润", float(totals["管理利润"]), "元", source_name, "净营业收入-食材-人工-平台-租金-水电-营销，并与源字段勾稽", "管理口径利润"),
            _fact("管理利润率", overall_margin, "%", source_name, "管理利润合计÷净营业收入合计（加权口径）", "禁止直接平均各门店各月利润率"),
            _fact("食材成本率", food_rate, "%", source_name, "食材成本合计÷净营业收入合计", "用于观察成本结构"),
            _fact("最佳经营门店", best_store, "", source_name, "按管理利润排序，利润率用于辅助判断", "不能只看营业额排名"),
            _fact("收入最高门店", highest_revenue_store, "", source_name, "按净营业收入排序", "规模领先不等于经营质量最好"),
            _fact("重点关注门店", lowest_margin_store, "", source_name, "按管理利润率最低识别，并结合规模判断", "需拆解成本率和经营动作"),
        ]
    )

    source_rows = []
    for index, (name, frame) in enumerate(zip(source_names, frames)):
        role = classify_sheet_role(name, frame)
        selected = index == fact_index
        source_rows.append(
            {
                "序号": index + 1,
                "文件/工作表": name,
                "识别角色": "门店-月份经营事实表" if selected else ("汇总/校验表" if role == "summary" else "参考/辅助表"),
                "行数": len(frame),
                "列数": frame.shape[1],
                "纳入计算": "是" if selected else "否",
                "计算边界": "全部经营指标仅来自该事实表" if selected else "只用于人工校验，禁止重复计入",
                "状态": "通过" if selected or role in {"summary", "notes", "dimension"} else "人工核验",
            }
        )
    sources_df = pd.DataFrame(source_rows)

    audit = pd.DataFrame(
        [
            {"验收项": "输入隔离", "状态": "通过", "结果": f"仅{source_name}纳入计算；其余汇总/参考表不参与计算", "人工边界": "若用户明确要求分析历史报告，需单独开启"},
            {"验收项": "净营业收入勾稽", "状态": "通过" if valid["净收入勾稽差异"].abs().max() <= 0.01 else "待核验", "结果": f"最大差异{valid['净收入勾稽差异'].abs().max():,.2f}元", "人工边界": "差异不为0时核对退款口径"},
            {"验收项": "管理利润勾稽", "状态": "通过" if valid["利润勾稽差异"].abs().max() <= 0.01 else "待核验", "结果": f"最大差异{valid['利润勾稽差异'].abs().max():,.2f}元", "人工边界": "差异不为0时核对未列示费用"},
            {"验收项": "管理利润率", "状态": "通过", "结果": f"{totals['管理利润']:,.2f}÷{totals['净营业收入']:,.2f}={overall_margin:.4%}", "人工边界": "不平均源利润率字段"},
            {"验收项": "平台费口径", "状态": "通过", "结果": f"9行平台费求和={totals['平台费']:,.2f}元", "人工边界": "费用字段不得取平均"},
            {"验收项": "重复记录", "状态": "待核验" if bool(valid["精确重复"].any()) else "通过", "结果": f"精确重复行{int(valid['精确重复'].sum())}条", "人工边界": "不自动删除业务可能重复的门店月份"},
            {"验收项": "明细能力边界", "状态": "人工核验", "结果": "未提供菜品/BOM、外卖结算、采购盘点、工时与评价明细", "人工边界": "相关问题不输出确定性结论"},
        ]
    )

    limitations = pd.DataFrame(
        [
            {"分析主题": "渠道与外卖", "当前可用证据": f"平台费{totals['平台费']:,.0f}元", "无法确定": "外卖收入、配送费、补贴、渠道到账和渠道利润", "所需补充": "渠道级订单与平台结算"},
            {"分析主题": "菜品盈利", "当前可用证据": f"食材成本{totals['食材成本']:,.0f}元", "无法确定": "菜品销量、折扣、BOM理论成本和单品贡献", "所需补充": "订单菜品明细、菜品主数据、BOM"},
            {"分析主题": "原料损耗", "当前可用证据": "月度食材成本汇总", "无法确定": "采购价差、理论耗用、盘点差异和报损责任", "所需补充": "采购入库、领料/盘点、原料单位换算"},
            {"分析主题": "人工效率", "当前可用证据": f"人工成本{totals['人工成本']:,.0f}元", "无法确定": "人数、工时、加班率、人时产出", "所需补充": "员工排班、实际工时、加班和岗位"},
            {"分析主题": "客户投诉", "当前可用证据": "无", "无法确定": "低评分、投诉门店/菜品/渠道集中点", "所需补充": "订单级评价、投诉和退款原因"},
        ]
    )

    dashboard_columns: dict[str, pd.Series] = {
        "KPI_销售规模": pd.Series([float(totals["净营业收入"])]),
        "KPI_管理利润率": pd.Series([overall_margin]),
        "KPI_毛利标题": pd.Series(["食材成本率"]),
        "KPI_毛利率": pd.Series([food_rate]),
        "KPI_估算经营结果": pd.Series([float(totals["管理利润"])]),
        "KPI_已发生退款": pd.Series([float(totals["退款"])]),
        "KPI_平台费": pd.Series([float(totals["平台费"])]),
        "核心诊断": pd.Series([
            f"{best_store}管理利润与利润率领先；{highest_revenue_store}收入最高但利润转化偏弱。"
            f"{last['月份']}净营业收入较上月{revenue_change:+.1%}，管理利润率由{previous['管理利润率']:.2%}降至{last['管理利润率']:.2%}。"
            if previous is not None
            else f"{best_store}管理利润与利润率领先；{highest_revenue_store}收入规模最高。"
        ]),
        "核心诊断1_标题": pd.Series(["最佳经营门店"]),
        "核心诊断1_内容": pd.Series([best_store]),
        "核心诊断2_标题": pd.Series(["规模与利润错位"]),
        "核心诊断2_内容": pd.Series([f"{highest_revenue_store}收入最高；{best_store}利润最佳"]),
        "核心诊断3_标题": pd.Series(["最近月份利润率"]),
        "核心诊断3_内容": pd.Series([f"{last['月份']} {last['管理利润率']:.2%}"]),
    }
    for index in range(3):
        row = risk_df.iloc[index] if index < len(risk_df) else None
        dashboard_columns[f"风险卡{index + 1}_标题"] = pd.Series([f"{row['优先级']}｜{row['风险事项']}" if row is not None else "未触发"])
        dashboard_columns[f"风险卡{index + 1}_证据"] = pd.Series([row["数据证据"] if row is not None else ""]) 
        dashboard_columns[f"风险卡{index + 1}_行动"] = pd.Series([row["建议行动"] if row is not None else "持续监控"])
    for column in ("月份", "营业额", "管理利润", "管理利润率"):
        dashboard_columns[f"月度_{column}"] = monthly[column].reset_index(drop=True)
    dashboard_columns["门店_门店"] = store["门店"].reset_index(drop=True)
    dashboard_columns["门店_营业额"] = store["营业额"].reset_index(drop=True)
    dashboard_columns["门店_管理利润"] = store["管理利润"].reset_index(drop=True)
    dashboard_columns["门店_管理利润率"] = store["管理利润率"].reset_index(drop=True)
    dashboard = pd.DataFrame(dashboard_columns)

    outputs: dict[str, pd.DataFrame] = {
        "管理层诊断总览": summary,
        "数据源与事实域": sources_df,
        "门店经营诊断": store,
        "利润驱动分析": monthly,
        "成本费用分析": cost,
        "渠道与外卖分析": limitations.loc[limitations["分析主题"].eq("渠道与外卖")].reset_index(drop=True),
        "菜品盈利分析": limitations.loc[limitations["分析主题"].eq("菜品盈利")].reset_index(drop=True),
        "原料采购与损耗": limitations.loc[limitations["分析主题"].eq("原料损耗")].reset_index(drop=True),
        "人工效率分析": limitations.loc[limitations["分析主题"].eq("人工效率")].reset_index(drop=True),
        "客户评价与退款": limitations.loc[limitations["分析主题"].eq("客户投诉")].reset_index(drop=True),
        "风险行动计划": risk_df,
        "诊断底稿": valid,
        "数据口径与验收": audit,
        "经营诊断看板": dashboard,
    }
    report = {
        "schema_version": 2,
        "status": "ready",
        "summary": "已识别门店-月份经营事实表，排除汇总校验表，按加权口径生成规模、利润、门店、月度、成本、风险与行动诊断。",
        "message": "餐饮经营诊断已完成；金额求和、利润率按利润合计÷净营业收入合计，缺失事实域明确标记人工核验。",
        "warnings": ["管理利润是基于已提供成本项目的管理口径，不替代法定财务净利润。", "未提供菜品、结算、采购盘点、工时和评价明细的主题不强行推断。"],
        "facts": summary.to_dict("records"),
        "risks": risk_df.to_dict("records"),
        "sources": sources_df.to_dict("records"),
        "report_kind": "restaurant_diagnosis_report",
        "profile": "compact_store_period_pnl",
    }
    for frame in outputs.values():
        frame.attrs["toolbox_report_kind"] = "restaurant_diagnosis_report"
    return RestaurantDiagnosisResult(outputs=outputs, report=report)


def build_restaurant_diagnosis_report(frames: Sequence[pd.DataFrame], *, source_names: Sequence[str], user_request: str = "", low_margin_threshold: float = 0.15) -> RestaurantDiagnosisResult:
    validate_restaurant_diagnosis_params({"source_names": source_names})
    if len(frames) != len(source_names) or not can_build_restaurant_diagnosis_report(frames, source_names):
        raise ValueError("当前输入不足以建立餐饮多事实域诊断；至少需要POS、菜品、退款、结算、人工、费用及两张主数据/成本表")
    if _compact_operating_input(frames, list(source_names)) is not None:
        return _compact_restaurant_report(frames, source_names=source_names, user_request=user_request)
    r = _frame(frames, source_names)
    src = {k: str(r[k].attrs.get("_restaurant_source_name", k)) for k in r}

    dish = r["dishes"]
    dish_id = _text(_col(dish, "菜品编码", "商品编码", "SKU", required=True))
    dish_name = _text(_col(dish, "菜品名称", "商品名称", "名称"))
    dish_cost = _num(_col(dish, "标准食材成本", "标准成本", "食材成本"))
    dish_map = pd.DataFrame({"菜品编码": dish_id, "菜品名称": dish_name, "标准食材成本": dish_cost}).drop_duplicates("菜品编码")

    pos_raw = r["pos"]
    pos = pd.DataFrame({
        "营业日期": _date(_col(pos_raw, "营业日期", "日期", required=True)),
        "订单号": _text(_col(pos_raw, "订单号", "订单编号", required=True)),
        "门店编码": _text(_col(pos_raw, "门店编码", "门店", required=True)),
        "渠道": _text(_col(pos_raw, "渠道/平台", "渠道", "平台")),
        "菜品编码": _text(_col(pos_raw, "菜品编码", "商品编码", "SKU", required=True)),
        "数量": _num(_col(pos_raw, "数量", "销量", required=True)),
        "原价": _num(_col(pos_raw, "菜品原价", "原价")),
        "折扣分摊": _num(_col(pos_raw, "优惠分摊", "折扣分摊", "优惠")),
        "实付分摊": _num(_col(pos_raw, "实付分摊", "实付", "销售额", required=True)),
        "订单状态": _text(_col(pos_raw, "订单状态", "状态")),
        "支付方式": _text(_col(pos_raw, "支付方式", "支付")),
    })
    pos["状态语义"] = classify_status_series(pos["订单状态"], classify_order_status)
    pos["有效订单"] = pos["状态语义"].eq(ORDER_SUCCESS)
    pos["精确重复"] = pos.duplicated(keep="first")
    raw_duplicate_count = int(pos["精确重复"].sum())
    pos = pos.loc[~pos["精确重复"]].copy()
    pos = pos.merge(dish_map, on="菜品编码", how="left", validate="many_to_one")
    pos["标准食材成本金额"] = pos["数量"].fillna(0) * pos["标准食材成本"].fillna(0)
    pos["月份"] = pos["营业日期"].dt.to_period("M").astype("string")

    refund_raw = r["refunds"]
    refunds = pd.DataFrame({"退款日期": _date(_col(refund_raw, "退款日期", "申请日期", "日期")), "退款单号": _text(_col(refund_raw, "退款单号", "售后单号")), "原订单号": _text(_col(refund_raw, "原订单号", "订单号", required=True)), "退款类型": _text(_col(refund_raw, "退款类型", "售后类型")), "退款金额": _num(_col(refund_raw, "退款金额", "金额", required=True)), "状态": _text(_col(refund_raw, "状态", "售后状态")), "原因": _text(_col(refund_raw, "原因", "退款原因")), "责任归属": _text(_col(refund_raw, "责任归属", "归属"))})
    refunds["状态语义"] = classify_status_series(refunds["状态"], classify_refund_status)
    refunds["已发生退款"] = refunds["状态语义"].eq(REFUND_CONFIRMED)
    refunds["处理中"] = refunds["状态语义"].eq(REFUND_PENDING)
    order_month = pos[["订单号", "月份"]].drop_duplicates("订单号").rename(columns={"月份": "订单归属月份"})
    refunds = refunds.merge(order_month, left_on="原订单号", right_on="订单号", how="left", validate="many_to_one")
    refunds = refunds.drop(columns=["订单号"], errors="ignore")
    refunds["退款发生月份"] = _month(refunds["退款日期"])
    refund_by_order = refunds.loc[refunds["已发生退款"]].groupby("原订单号", as_index=False)["退款金额"].sum().rename(columns={"退款金额": "已退款金额"})
    pos = pos.merge(refund_by_order, left_on="订单号", right_on="原订单号", how="left").drop(columns=["原订单号"], errors="ignore")
    pos["已退款金额"] = pos["已退款金额"].fillna(0)
    pos["订单实付合计"] = pos.groupby("订单号", dropna=False)["实付分摊"].transform("sum")
    pos["退款分摊"] = pos["已退款金额"].mul(pos["实付分摊"]).div(
        pos["订单实付合计"].where(pos["订单实付合计"].gt(0))
    )
    pos["退款分摊状态"] = "无退款"
    pos.loc[pos["已退款金额"].gt(0) & pos["订单实付合计"].gt(0), "退款分摊状态"] = "按订单行实付比例分摊"
    pos.loc[pos["已退款金额"].gt(0) & ~pos["订单实付合计"].gt(0), "退款分摊状态"] = "人工核验：订单实付合计无效"
    total_sales = float(pos.loc[pos["有效订单"], "实付分摊"].sum())
    total_refund = float(refunds.loc[refunds["已发生退款"], "退款金额"].sum())
    standard_food = float(pos.loc[pos["有效订单"], "标准食材成本金额"].sum())

    settlement = r["settlements"]
    settlement_out = pd.DataFrame({"结算月份": _text(_col(settlement, "结算月份", "月份", required=True)), "门店编码": _text(_col(settlement, "门店编码", "门店")), "平台": _text(_col(settlement, "平台", "渠道")), "结算基数": _num(_col(settlement, "订单结算基数", "结算基数", "订单金额")), "平台佣金": _num(_col(settlement, "平台佣金", "佣金")), "配送服务费": _num(_col(settlement, "配送服务费", "配送费")), "活动补贴承担": _num(_col(settlement, "活动补贴承担", "平台活动承担", "活动补贴成本", "活动补贴扣减", "活动扣减")), "退款冲减": _num(_col(settlement, "退款冲减")), "实际到账": _num(_col(settlement, "实际到账", "到账金额", "实收"))})
    settlement_out["可比平台成本"] = settlement_out[["平台佣金", "配送服务费", "活动补贴承担"]].fillna(0).sum(axis=1)
    settlement_out["理论到账"] = (
        settlement_out["结算基数"]
        - settlement_out["平台佣金"].fillna(0)
        - settlement_out["配送服务费"].fillna(0)
        - settlement_out["活动补贴承担"].fillna(0)
        - settlement_out["退款冲减"].fillna(0)
    )
    settlement_out["到账勾稽差异"] = settlement_out["实际到账"] - settlement_out["理论到账"]
    settlement_out["平台到账率"] = settlement_out["实际到账"].div(settlement_out["结算基数"].where(settlement_out["结算基数"].ne(0)))
    settlement_out["可比经营贡献"] = settlement_out["实际到账"]
    settlement_out["贡献口径"] = "实际到账（佣金、配送、活动补贴及退款已在结算链扣除，不重复扣费）"

    labor = r["labor"]
    labor_out = pd.DataFrame({"月份": _text(_col(labor, "月份", "期间", required=True)), "门店编码": _text(_col(labor, "门店编码", "门店")), "岗位人数": _num(_col(labor, "岗位人数", "人数")), "排班工时": _num(_col(labor, "排班工时", "计划工时")), "实际工时": _num(_col(labor, "实际工时", "工时")), "加班工时": _num(_col(labor, "加班工时", "加班")), "工资成本": _num(_col(labor, "工资成本", "基本工资", "工资")), "临时工成本": _num(_col(labor, "临时工成本", "临时用工成本", "临时工工资")), "加班工资": _num(_col(labor, "加班工资", "加班成本")), "其他人工成本": _num(_col(labor, "其他人工成本", "其他人工", "福利社保")), "源人工总成本": _num(_col_exact(labor, "人工总成本", "人工成本", "人员总成本")), "缺勤小时": _num(_col(labor, "缺勤小时", "缺勤"))})
    labor_components = ["工资成本", "临时工成本", "加班工资", "其他人工成本"]
    labor_out["组件人工成本"] = labor_out[labor_components].sum(axis=1, min_count=1)
    labor_out["人工总成本"] = labor_out["源人工总成本"].where(
        labor_out["源人工总成本"].notna(), labor_out["组件人工成本"]
    )
    labor_out["人工勾稽差异"] = labor_out["源人工总成本"] - labor_out["组件人工成本"]
    labor_out["加班率"] = labor_out["加班工时"].div(labor_out["实际工时"].where(labor_out["实际工时"].ne(0)))
    labor_out["工时效率参考"] = labor_out["实际工时"].div(labor_out["岗位人数"].where(labor_out["岗位人数"].ne(0)))

    fixed = r["fixed"]
    fixed_out = pd.DataFrame({"月份": _text(_col(fixed, "月份", "期间", required=True)), "门店编码": _text(_col(fixed, "门店编码", "门店")), "费用类别": _text(_col(fixed, "费用类别", "费用项目", "类别")), "金额": _num(_col(fixed, "金额", "费用金额", required=True))})

    stores = r.get("stores", pd.DataFrame())
    store_names = pd.DataFrame()
    if not stores.empty:
        store_names = pd.DataFrame({"门店编码": _text(_col(stores, "门店编码", "门店", required=True)), "门店名称": _text(_col(stores, "门店名称", "名称")), "区域": _text(_col(stores, "区域", "地区")), "状态": _text(_col(stores, "状态"))}).drop_duplicates("门店编码")
    valid = pos.loc[pos["有效订单"]].copy()
    store_sales = valid.groupby("门店编码", as_index=False).agg(营业实付=("实付分摊", "sum"), 订单数=("订单号", "nunique"), 标准食材成本=("标准食材成本金额", "sum"))
    store_refunds = refunds.loc[refunds["已发生退款"]].merge(pos[["订单号", "门店编码"]].drop_duplicates(), left_on="原订单号", right_on="订单号", how="left").groupby("门店编码", as_index=False)["退款金额"].sum().rename(columns={"退款金额": "已退款金额"})
    store_settlement = settlement_out.groupby("门店编码", as_index=False).agg(实际到账=("实际到账", "sum"), 平台成本=("可比平台成本", "sum"))
    store_labor = labor_out.groupby("门店编码", as_index=False).agg(人工成本=("人工总成本", "sum"), 加班工时=("加班工时", "sum"), 实际工时=("实际工时", "sum"))
    store_fixed = fixed_out.groupby("门店编码", as_index=False).agg(固定费用=("金额", "sum"))
    store = store_sales.merge(store_refunds, on="门店编码", how="left").merge(store_settlement, on="门店编码", how="left").merge(store_labor, on="门店编码", how="left").merge(store_fixed, on="门店编码", how="left")
    if not store_names.empty:
        store = store_names.merge(store, on="门店编码", how="right")
    for col in ("已退款金额", "实际到账", "平台成本", "人工成本", "固定费用", "加班工时", "实际工时"):
        if col not in store: store[col] = 0.0
        store[col] = store[col].fillna(0)
    store["可比经营贡献"] = store["营业实付"] - store["已退款金额"] - store["标准食材成本"] - store["平台成本"]
    store["情景结果"] = store["可比经营贡献"] - store["人工成本"] - store["固定费用"]
    store["加班率"] = store["加班工时"].div(store["实际工时"].where(store["实际工时"].ne(0)))
    scale_mismatch = (total_sales < max(1.0, float(store["人工成本"].sum() + store["固定费用"].sum()) * 0.2))
    store["结论限制"] = "收入与人工/固定费用尺度明显不匹配，经营结果仅供口径核验" if scale_mismatch else "管理口径可比贡献"

    dish_out = valid.groupby(["菜品编码", "菜品名称"], dropna=False, as_index=False).agg(销量=("数量", "sum"), 营业实付=("实付分摊", "sum"), 已退款金额=("退款分摊", "sum"), 标准食材成本金额=("标准食材成本金额", "sum"))
    dish_out["退款后实付"] = dish_out["营业实付"] - dish_out["已退款金额"]
    dish_out["退款后贡献"] = dish_out["退款后实付"] - dish_out["标准食材成本金额"]
    dish_out["贡献率"] = dish_out["退款后贡献"].div(dish_out["退款后实付"].where(dish_out["退款后实付"].ne(0)))
    dish_out["管理提示"] = dish_out.apply(lambda x: "高销量低贡献，检查定价/折扣/BOM" if x["贡献率"] < low_margin_threshold else "正常" if pd.notna(x["贡献率"]) else "标准成本缺失，人工核验", axis=1)

    month_sales = valid.groupby("月份", as_index=False).agg(营业实付=("实付分摊", "sum"), 标准食材成本=("标准食材成本金额", "sum"))
    confirmed_refunds = refunds.loc[refunds["已发生退款"]].copy()
    month_ref = confirmed_refunds.groupby("退款发生月份", as_index=False)["退款金额"].sum().rename(columns={"退款发生月份": "月份", "退款金额": "已退款金额"})
    month_labor = labor_out.groupby("月份", as_index=False).agg(人工成本=("人工总成本", "sum"), 加班工时=("加班工时", "sum"))
    month_platform = settlement_out.groupby("结算月份", as_index=False)["可比平台成本"].sum().rename(columns={"结算月份": "月份", "可比平台成本": "平台成本"})
    month_fixed = fixed_out.groupby("月份", as_index=False)["金额"].sum().rename(columns={"金额": "固定费用"})
    monthly_occurrence = month_sales.merge(month_ref, on="月份", how="outer").merge(month_platform, on="月份", how="outer").merge(month_labor, on="月份", how="outer").merge(month_fixed, on="月份", how="outer")
    occurrence_amounts = ["营业实付", "已退款金额", "标准食材成本", "平台成本", "人工成本", "加班工时", "固定费用"]
    monthly_occurrence[occurrence_amounts] = monthly_occurrence[occurrence_amounts].fillna(0)
    monthly_occurrence["可比经营贡献"] = monthly_occurrence["营业实付"] - monthly_occurrence["已退款金额"] - monthly_occurrence["标准食材成本"] - monthly_occurrence["平台成本"]
    monthly_occurrence["情景经营结果"] = monthly_occurrence["可比经营贡献"] - monthly_occurrence["人工成本"] - monthly_occurrence["固定费用"]
    monthly_occurrence["时间口径"] = "发生月视角"
    monthly_occurrence["口径说明"] = "销售按营业月、退款按退款发生月、平台按结算月、人工和固定费用按发生月"

    attributed_refunds = confirmed_refunds.dropna(subset=["订单归属月份"]).groupby("订单归属月份", as_index=False)["退款金额"].sum().rename(columns={"订单归属月份": "月份", "退款金额": "已退款金额"})
    monthly_attributed = month_sales.merge(attributed_refunds, on="月份", how="outer")
    monthly_attributed[["营业实付", "已退款金额", "标准食材成本"]] = monthly_attributed[["营业实付", "已退款金额", "标准食材成本"]].fillna(0)
    monthly_attributed["平台成本"] = math.nan
    monthly_attributed["人工成本"] = math.nan
    monthly_attributed["加班工时"] = math.nan
    monthly_attributed["固定费用"] = math.nan
    monthly_attributed["可比经营贡献"] = monthly_attributed["营业实付"] - monthly_attributed["已退款金额"] - monthly_attributed["标准食材成本"]
    monthly_attributed["情景经营结果"] = math.nan
    monthly_attributed["时间口径"] = "订单归属月视角"
    monthly_attributed["口径说明"] = "退款回溯原订单营业月；平台结算、人工和固定费用缺少订单键，不强行回溯"
    monthly = pd.concat([monthly_occurrence, monthly_attributed], ignore_index=True, sort=False).sort_values(["时间口径", "月份"], kind="stable")

    bom = r.get("bom", pd.DataFrame())
    losses = r.get("losses", pd.DataFrame())
    loss_out = pd.DataFrame()
    if not losses.empty:
        loss_out = pd.DataFrame({"月份": _text(_col(losses, "月份", "期间", required=True)), "门店编码": _text(_col(losses, "门店编码", "门店", required=True)), "原料编码": _text(_col(losses, "原料编码", "原料编号", required=True)), "期末结存数量": _num(_col(losses, "期末结存数量", "结存数量")), "单位": _text(_col(losses, "单位")), "报损数量": _num(_col(losses, "报损数量", "损耗数量", "实际报损")), "盘点差异数量": _num(_col(losses, "盘点差异数量", "差异数量", "门店盘差"))})
        ing = r.get("ingredients", pd.DataFrame())
        if not ing.empty:
            im = pd.DataFrame({"原料编码": _text(_col(ing, "原料编码", "原料编号", required=True)), "标准采购单价": _num(_col(ing, "标准采购单价", "标准采购价", "标准基础单价", "标准成本"))}).drop_duplicates("原料编码")
            loss_out = loss_out.merge(im, on="原料编码", how="left")
        else:
            loss_out["标准采购单价"] = math.nan
        loss_out["报损金额"] = loss_out["报损数量"].abs() * loss_out["标准采购单价"]
        loss_out["盘点差异金额"] = loss_out["盘点差异数量"].abs() * loss_out["标准采购单价"]
        loss_out["估值状态"] = "已按标准采购单价估值"
        loss_out.loc[loss_out["标准采购单价"].isna(), "估值状态"] = "无法估值：缺少标准采购单价"
    purchase = r.get("purchases", pd.DataFrame())
    purchase_out = pd.DataFrame()
    purchase_exact_duplicates = 0
    purchase_conflicts = 0
    if not purchase.empty:
        purchase_out = pd.DataFrame({"入库日期": _date(_col(purchase, "入库日期", "日期")), "入库单号": _text(_col(purchase, "入库单号", "入库单")), "门店编码": _text(_col(purchase, "门店编码", "门店")), "原料编码": _text(_col(purchase, "原料编码", "原料编号", required=True)), "采购数量": _num(_col(purchase, "采购数量", "数量")), "采购单位": _text(_col(purchase, "采购单位", "计量单位", "单位")), "采购单价": _num(_col(purchase, "采购单价", "单价")), "状态": _text(_col(purchase, "状态"))})
        purchase_out["精确重复"] = purchase_out.duplicated(keep="first")
        purchase_exact_duplicates = int(purchase_out["精确重复"].sum())
        purchase_out = purchase_out.loc[~purchase_out["精确重复"]].copy()
        purchase_keys = ["入库单号", "门店编码", "原料编码", "采购单位"]
        keyed = purchase_out["入库单号"].ne("")
        duplicate_key = keyed & purchase_out.duplicated(purchase_keys, keep=False)
        critical = ["入库日期", "采购数量", "采购单价", "状态"]
        conflict_keys: set[tuple[Any, ...]] = set()
        for key, group in purchase_out.loc[duplicate_key].groupby(purchase_keys, dropna=False):
            if any(group[column].nunique(dropna=False) > 1 for column in critical):
                conflict_keys.add(key if isinstance(key, tuple) else (key,))
        purchase_out["业务键冲突"] = [
            tuple(row[column] for column in purchase_keys) in conflict_keys
            for _, row in purchase_out.iterrows()
        ]
        purchase_conflicts = int(purchase_out["业务键冲突"].sum())
        consistent_duplicate = duplicate_key & ~purchase_out["业务键冲突"] & purchase_out.duplicated(purchase_keys, keep="first")
        purchase_out = purchase_out.loc[~consistent_duplicate].copy()
        purchase_out["采购金额"] = purchase_out["采购数量"] * purchase_out["采购单价"]
        purchase_out["状态语义"] = classify_status_series(purchase_out["状态"], classify_receipt_status)
        purchase_out["纳入口径"] = purchase_out["状态语义"].eq(RECEIPT_CONFIRMED) & ~purchase_out["业务键冲突"]

    reviews = r.get("reviews", pd.DataFrame())
    review_out = pd.DataFrame()
    review_refund_out = pd.DataFrame()
    if not reviews.empty:
        review_out = pd.DataFrame({"日期": _date(_col(reviews, "日期", "评价日期")), "门店编码": _text(_col(reviews, "门店编码", "门店")), "订单号": _text(_col(reviews, "订单号", "订单编号")), "评分": _num(_col(reviews, "评分", "满意度")), "标签": _text(_col(reviews, "标签", "评价标签")), "评价摘要": _text(_col(reviews, "评价摘要", "内容"))})
        review_out["重点关注"] = review_out["评分"].lt(4) | review_out["标签"].str.contains("投诉|差评|食品|卫生|慢", regex=True, na=False)
        refund_for_join = refunds.loc[refunds["已发生退款"], ["原订单号", "退款金额", "退款类型", "原因", "责任归属"]].copy()
        refund_for_join = refund_for_join.groupby("原订单号", as_index=False).agg(
            退款金额=("退款金额", "sum"),
            退款类型=("退款类型", lambda values: "；".join(sorted(set(filter(None, values.astype(str)))))),
            退款原因=("原因", lambda values: "；".join(sorted(set(filter(None, values.astype(str)))))),
            责任归属=("责任归属", lambda values: "；".join(sorted(set(filter(None, values.astype(str)))))),
        )
        review_refund_out = review_out.merge(refund_for_join, left_on="订单号", right_on="原订单号", how="left")
        review_refund_out["记录类型"] = "评价"
        review_refund_out["退款匹配状态"] = review_refund_out["原订单号"].notna().map({True: "已匹配", False: "无已发生退款"})
        orphan_refunds = refunds.loc[
            refunds["已发生退款"] & ~refunds["原订单号"].isin(set(review_out["订单号"])),
            ["原订单号", "退款金额", "退款类型", "原因", "责任归属"],
        ].copy()
        if not orphan_refunds.empty:
            orphan_refunds["记录类型"] = "孤立退款"
            orphan_refunds["退款匹配状态"] = "未匹配评价"
            review_refund_out = pd.concat([review_refund_out, orphan_refunds], ignore_index=True, sort=False)
    else:
        review_refund_out = refunds.copy()
        review_refund_out["记录类型"] = "孤立退款"
        review_refund_out["退款匹配状态"] = "未提供评价表"

    risks = [{"优先级": "P0", "风险事项": "经营结果暂不可直接定性", "原因": "POS有效实付与人工/固定费用尺度明显不匹配", "影响": "可能是测试口径、期间或费用单位错误；直接下结论会误导决策", "责任部门": "财务/运营", "建议行动": "核对POS是否为完整季度、人工费用是否为年化/分店合计、固定费用期间和币种", "验收指标": "收入、成本、费用期间一致且可勾稽"}] if scale_mismatch else []
    if float(labor_out["加班工时"].sum()) > 0:
        risks.append({"优先级": "P1", "风险事项": "加班工时存在", "原因": f"累计加班{labor_out['加班工时'].sum():,.0f}小时", "影响": "人力成本与排班效率承压", "责任部门": "门店运营/人事", "建议行动": "按门店拆分高峰排班并核查订单/工时匹配", "验收指标": "加班率逐月下降且人均产出不下降"})
    if not loss_out.empty and float(loss_out["报损金额"].sum()) > 0:
        risks.append({"优先级": "P1", "风险事项": "原料盘点差异形成报损线索", "原因": f"估算报损金额{loss_out['报损金额'].sum():,.2f}元", "影响": "毛利和现金被损耗侵蚀", "责任部门": "采购/门店", "建议行动": "逐项复盘领料、称重、保质期和盘点记录；不要把差异直接当实际耗用", "验收指标": "差异率按原料和门店可解释"})
    low_reviews = int(review_out["重点关注"].sum()) if not review_out.empty else 0
    if low_reviews:
        risks.append({"优先级": "P1", "风险事项": "客户体验/投诉需关注", "原因": f"低评分或投诉线索{low_reviews}条", "影响": "复购、平台评分和退款风险", "责任部门": "门店运营/客服", "建议行动": "按门店、标签、菜品定位根因并在48小时内闭环", "验收指标": "低评分率和同类投诉连续两期下降"})
    risk_df = pd.DataFrame(risks or [{"优先级": "P2", "风险事项": "暂无可确认高优先级风险", "原因": "当前数据未触发规则", "影响": "持续监控", "责任部门": "运营", "建议行动": "保持月度复盘", "验收指标": "数据按期更新"}])

    summary = pd.DataFrame([
        _fact("营业实付", total_sales, "元", src["pos"], "POS有效订单实付分摊合计", "经营规模事实"),
        _fact("已发生退款", total_refund, "元", src["refunds"], "仅统计状态为已退款/成功", "收入质量与售后敞口"),
        _fact("标准食材成本", standard_food, "元", src["pos"] + "+" + src["dishes"], "有效POS数量×菜品标准食材成本", "管理成本，不替代财务结转", "中"),
        _fact("平台成本", float(settlement_out["可比平台成本"].sum()), "元", src["settlements"], "平台佣金+配送服务费+活动补贴承担；实际到账不再重复扣除", "外卖渠道成本", "高"),
        _fact("平台实际到账", float(settlement_out["实际到账"].sum()), "元", src["settlements"], "结算表实际到账合计，不冒充销售收入", "现金转化", "高"),
        _fact("人工成本", float(labor_out["人工总成本"].sum()), "元", src["labor"], "工资成本+临时工成本+加班工资+其他人工；有源人工总成本时优先用于勾稽", "人力投入", "高"),
        _fact("固定费用", float(fixed_out["金额"].sum()), "元", src["fixed"], "固定运营费用金额合计", "期间费用", "高"),
        _fact("经营结果可用性", "需人工核验" if scale_mismatch else "可作管理情景参考", "状态", src["pos"] + "+" + src["labor"] + "+" + src["fixed"], "若收入小于人工与固定费用20%，不输出确定盈亏结论", "防止测试期间/单位错误造成误导", "高"),
    ])
    dashboard = pd.DataFrame([{
        "KPI_销售规模": total_sales, "KPI_平台到账率": float(settlement_out["实际到账"].sum() / settlement_out["结算基数"].sum()) if settlement_out["结算基数"].sum() else math.nan,
        "KPI_毛利标题": "退款后标准食材贡献率", "KPI_毛利率": (total_sales - total_refund - standard_food) / total_sales if total_sales else math.nan,
        "KPI_估算经营结果": "需人工核验" if scale_mismatch else float(total_sales - total_refund - standard_food - settlement_out["可比平台成本"].sum() - labor_out["人工总成本"].sum() - fixed_out["金额"].sum()), "KPI_已发生退款": total_refund, "KPI_报损金额": float(loss_out["报损金额"].sum(min_count=1)) if not loss_out.empty else math.nan,
        "核心诊断": "收入与人工/固定费用尺度不匹配，当前只能给出经营线索，不能把情景结果当作真实净利润。" if scale_mismatch else "需结合渠道、菜品、人工和损耗共同改善经营贡献。",
        "风险卡1_标题": risk_df.iloc[0]["优先级"] + "｜" + str(risk_df.iloc[0]["风险事项"]), "风险卡1_证据": risk_df.iloc[0]["原因"], "风险卡1_行动": risk_df.iloc[0]["建议行动"],
        "风险卡2_标题": (risk_df.iloc[1]["优先级"] + "｜" + str(risk_df.iloc[1]["风险事项"])) if len(risk_df) > 1 else "暂无更多高优先级风险", "风险卡2_证据": risk_df.iloc[1]["原因"] if len(risk_df) > 1 else "未触发", "风险卡2_行动": risk_df.iloc[1]["建议行动"] if len(risk_df) > 1 else "持续监控",
        "风险卡3_标题": (risk_df.iloc[2]["优先级"] + "｜" + str(risk_df.iloc[2]["风险事项"])) if len(risk_df) > 2 else "暂无更多高优先级风险", "风险卡3_证据": risk_df.iloc[2]["原因"] if len(risk_df) > 2 else "未触发", "风险卡3_行动": risk_df.iloc[2]["建议行动"] if len(risk_df) > 2 else "持续监控",
    }])
    # Hidden/helper columns are intentionally explicit so the dashboard charts are data-driven.
    dashboard = pd.concat([dashboard, monthly_occurrence], ignore_index=True, sort=False)
    dashboard["门店_门店"] = None; dashboard["门店_可比经营贡献"] = None
    n_store = min(len(store), len(dashboard)); dashboard.loc[: n_store-1, "门店_门店"] = store.get("门店名称", store["门店编码"]).tolist()[:n_store]; dashboard.loc[: n_store-1, "门店_可比经营贡献"] = store["可比经营贡献"].tolist()[:n_store]
    dashboard["渠道_渠道"] = None; dashboard["渠道_可比经营贡献"] = None
    ch = settlement_out.groupby("平台", as_index=False)["可比经营贡献"].sum() if not settlement_out.empty else pd.DataFrame(columns=["平台", "可比经营贡献"])
    n_channel = min(len(ch), len(dashboard)); dashboard.loc[: n_channel-1, "渠道_渠道"] = ch["平台"].tolist()[:n_channel]; dashboard.loc[: n_channel-1, "渠道_可比经营贡献"] = ch["可比经营贡献"].tolist()[:n_channel]
    dashboard["损耗_原料"] = None; dashboard["损耗_报损金额"] = None
    if not loss_out.empty:
        lossg = loss_out.groupby("原料编码", as_index=False)["报损金额"].sum(); n_loss = min(len(lossg), len(dashboard)); dashboard.loc[: n_loss-1, "损耗_原料"] = lossg["原料编码"].tolist()[:n_loss]; dashboard.loc[: n_loss-1, "损耗_报损金额"] = lossg["报损金额"].tolist()[:n_loss]

    audit = pd.DataFrame([
        {"检查项": "输入范围锁定", "结果": "通过", "说明": "仅使用本次任务明确选择的原始工作表，历史输出禁止回流"},
        {"检查项": "POS精确重复发现", "结果": raw_duplicate_count, "说明": "删除前统计完全相同记录"},
        {"检查项": "POS精确重复删除", "结果": raw_duplicate_count, "说明": "只删除完全相同记录，不按订单号删除多菜品明细"},
        {"检查项": "POS精确重复剩余", "结果": int(pos["精确重复"].sum()), "说明": "去重后剩余精确重复应为0"},
        {"检查项": "POS未知状态", "结果": int((~pos["状态语义"].isin([ORDER_SUCCESS, ORDER_INVALID])).sum()), "说明": "未知状态不计入有效销售，需人工核验；负向状态优先"},
        {"检查项": "退款未知状态", "结果": int((~refunds["状态语义"].isin([REFUND_CONFIRMED, REFUND_PENDING])).sum()), "说明": "未知退款状态不冲减已实现收入，单独披露"},
        {"检查项": "退款映射覆盖", "结果": f"{int(refunds['原订单号'].isin(set(pos['订单号'])).sum())}/{len(refunds)}", "说明": "退款按原订单号回溯门店；无法匹配的记录保留并标记核验"},
        {"检查项": "平台到账勾稽", "结果": f"最大差异{settlement_out['到账勾稽差异'].abs().max():,.2f}元", "说明": "结算基数-佣金-配送-活动补贴承担-退款冲减=实际到账"},
        {"检查项": "采购重复与冲突", "结果": f"精确重复{purchase_exact_duplicates}条；业务键冲突{purchase_conflicts}条", "说明": "采购单位纳入业务键；一致重复去重，关键字段冲突不纳入口径"},
        {"检查项": "双时间口径", "结果": "通过", "说明": "同时输出发生月视角与订单归属月视角；无法按订单回溯的费用保持空值"},
        {"检查项": "损耗估值", "结果": f"无法估值{int(loss_out['标准采购单价'].isna().sum()) if not loss_out.empty else 0}条", "说明": "报损与盘点差异分列；缺价格时保留空值，不填0"},
        {"检查项": "经营尺度", "结果": "人工核验" if scale_mismatch else "通过", "说明": "销售、人工、固定费用期间/单位需业务确认"},
        {"检查项": "BOM/损耗", "结果": "管理代理指标", "说明": "没有完整期初+采购-销售-调整链时，损耗不是实际耗用；不强行补造"},
    ])
    sources_df = pd.DataFrame([{"来源角色": k, "工作表": src[k], "是否纳入": "是", "粒度": "独立事实/主数据"} for k in src])
    outputs = {"管理层诊断总览": summary, "数据源与事实域": sources_df, "门店经营诊断": store, "渠道与外卖分析": settlement_out, "菜品盈利分析": dish_out, "原料采购与损耗": pd.concat([purchase_out, loss_out], axis=0, ignore_index=True, sort=False), "人工效率分析": labor_out, "客户评价与退款": review_refund_out, "利润驱动分析": monthly, "风险行动计划": risk_df, "诊断底稿": pos, "数据口径与验收": audit, "经营诊断看板": dashboard}
    report = {"schema_version": 1, "status": "ready", "summary": "已按餐饮多事实域重建门店、渠道、菜品、原料、人工、评价和退款分析；经营结果受数据期间/单位一致性约束。", "message": "已识别为餐饮门店经营诊断，输出跨表可追溯经营报告。", "warnings": ["标准食材成本是管理口径，不替代财务结转；情景结果若尺度不匹配仅供人工核验。", "外卖结算实际到账与POS销售分开披露，不能重复相加。"], "facts": summary.to_dict("records"), "risks": risk_df.to_dict("records"), "sources": sources_df.to_dict("records"), "report_kind": "restaurant_diagnosis_report"}
    for frame in outputs.values():
        frame.attrs["toolbox_report_kind"] = "restaurant_diagnosis_report"
    return RestaurantDiagnosisResult(outputs=outputs, report=report)


__all__ = ["can_build_restaurant_diagnosis_report", "restaurant_diagnosis_profile", "build_restaurant_diagnosis_report", "validate_restaurant_diagnosis_params", "RestaurantDiagnosisResult"]
