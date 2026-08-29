"""Deterministic, evidence-backed multi-fact enterprise diagnosis."""

from __future__ import annotations

import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from difflib import SequenceMatcher
from typing import Any

import pandas as pd

_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:/.]+")


def _normalise(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


_ALIASES: Mapping[str, tuple[str, ...]] = {
    "date": ("日期", "业务日期", "订单日期", "交易日期", "开票日期", "date"),
    "month": ("月份", "期间", "年月", "会计期间", "month"),
    "document": ("单据号", "订单号", "订单编号", "流水号", "业务编号", "合同号", "orderid"),
    "type": ("业务类型", "订单类型"),
    "department": ("部门", "事业部", "组织"),
    "customer": ("客户/供应商", "客户", "客户名称", "客户单位", "客商", "customer"),
    "product": ("产品", "产品名称", "商品", "商品名称", "物料", "物料名称", "品类", "sku名称", "sku"),
    "revenue": (
        "收入",
        "营业收入",
        "主营业务收入",
        "销售收入",
        "销售额",
        "销售金额",
        "成交金额",
        "订单金额",
        "开票金额",
        "含税金额",
        "revenue",
    ),
    "cost": ("成本", "业务成本", "销售成本", "主营业务成本", "出库成本", "cost"),
    "payment": ("付款状态", "回款状态", "结算状态", "订单状态", "状态", "paymentstatus"),
    "owner": ("负责人", "销售人员", "销售员", "人员", "员工姓名", "业务员", "销售代表", "客户经理", "owner"),
    "industry": ("行业", "客户行业"),
    "years": ("合作年限", "合作时间"),
    "satisfaction": ("满意度", "客户满意度"),
    "credit": ("信用等级", "客户信用"),
    "risk": ("回款风险", "风险等级", "风险"),
    "cycle": ("回款周期", "账期", "付款周期"),
    "performance_sales": ("销售额", "业绩", "业绩金额"),
    "gross_profit": ("毛利", "销售毛利"),
    "collection": ("回款金额", "已回款金额", "实收金额"),
    "target": ("目标完成率", "业绩完成率"),
    "score": ("客户评分", "绩效评分", "评分"),
    "complaints": ("投诉次数", "投诉"),
    "expense_category": ("费用类别", "费用类型", "费用项目", "科目", "费用科目"),
    "amount": ("金额", "费用金额", "发生额", "amount"),
    "inventory_quantity": ("库存数量", "现有库存", "当前库存", "结存数量", "库存"),
    "inventory_amount": ("库存金额", "库存价值", "存货金额", "存货价值"),
    "monthly_sales": ("月均销量", "月销量", "月均出库", "近30天销量", "近30天出库"),
    "lead_time": ("供应周期", "采购提前期", "交期"),
    "inventory_status": ("状态", "库存状态"),
    "material_cost": ("材料成本", "直接材料"),
    "labor_cost": ("人工成本", "直接人工"),
    "overhead": ("制造费用", "制造间接费"),
    "output": ("产量", "产出数量"),
}


def _find(frame: pd.DataFrame, aliases: Sequence[str]) -> str | None:
    columns = {_normalise(column): str(column) for column in frame.columns}
    for alias in aliases:
        if _normalise(alias) in columns:
            return columns[_normalise(alias)]
    matches = {
        column
        for key, column in columns.items()
        for alias in aliases
        if len(_normalise(alias)) >= 2 and _normalise(alias) in key
    }
    return next(iter(matches)) if len(matches) == 1 else None


def _has(frame: pd.DataFrame, key: str) -> bool:
    return _find(frame, _ALIASES[key]) is not None


def _role_score(frame: pd.DataFrame, role: str) -> int | None:
    if role == "transactions":
        if not all(_has(frame, key) for key in ("date", "customer", "revenue")):
            return None
        return 12 + 2 * sum(_has(frame, key) for key in ("document", "payment", "owner", "product", "cost"))
    if role == "customers":
        extras = sum(_has(frame, key) for key in ("industry", "credit", "risk", "cycle", "years"))
        return 10 + 2 * extras if _has(frame, "customer") and _has(frame, "satisfaction") and extras else None
    if role == "performance":
        extras = sum(_has(frame, key) for key in ("gross_profit", "collection", "target", "score", "complaints"))
        return 10 + 3 * extras if _has(frame, "owner") and _has(frame, "performance_sales") and extras else None
    if role == "expenses":
        return 16 if all(_has(frame, key) for key in ("month", "expense_category", "amount")) else None
    if role == "inventory":
        required = ("product", "inventory_quantity", "inventory_amount", "monthly_sales")
        return (
            16 + sum(_has(frame, key) for key in ("lead_time", "inventory_status"))
            if all(_has(frame, key) for key in required)
            else None
        )
    if role == "production_cost":
        components = sum(_has(frame, key) for key in ("material_cost", "labor_cost", "overhead", "cost"))
        return 10 + 3 * components if _has(frame, "month") and _has(frame, "product") and components >= 2 else None
    return None


def infer_enterprise_table_roles(frames: Sequence[pd.DataFrame]) -> dict[str, int]:
    """Identify independent fact domains; never choose one universal main table."""
    if not isinstance(frames, Sequence) or isinstance(frames, (str, bytes)):
        raise TypeError("企业经营诊断输入必须是表格数组")
    labels = {
        "transactions": "销售/订单事实",
        "customers": "客户主数据",
        "performance": "人员绩效",
        "expenses": "期间费用",
        "inventory": "库存事实",
    }
    roles: dict[str, int] = {}
    used: set[int] = set()
    for role in (*labels, "production_cost"):
        candidates = []
        for index, frame in enumerate(frames):
            if index in used or not isinstance(frame, pd.DataFrame) or frame.empty:
                continue
            score = _role_score(frame, role)
            if score is not None:
                candidates.append((score, len(frame), index))
        candidates.sort(reverse=True)
        if not candidates:
            if role in labels:
                raise ValueError(f"无法识别“{labels[role]}”表")
            continue
        roles[role] = candidates[0][2]
        used.add(candidates[0][2])
    return roles


def can_build_enterprise_diagnosis_report(frames: Sequence[pd.DataFrame]) -> bool:
    from .restaurant_report import can_build_restaurant_diagnosis_report
    if can_build_restaurant_diagnosis_report(frames):
        return True
    try:
        infer_enterprise_table_roles(frames)
        return True
    except (TypeError, ValueError):
        # Complex enterprise projects are dispatched to domain fact plugins;
        # the generic manufacturing/enterprise role set is not a gatekeeper.
        from .ecommerce_report import can_build_ecommerce_diagnosis_report

        return can_build_ecommerce_diagnosis_report(frames)


def validate_enterprise_diagnosis_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("企业经营诊断参数必须是对象")
    names = params.get("source_names")
    if (
        not isinstance(names, (list, tuple))
        or not names
        or not all(isinstance(item, str) and item.strip() for item in names)
    ):
        raise TypeError("source_names 必须是非空字符串数组")
    request = params.get("user_request", "")
    if not isinstance(request, str) or len(request) > 8_000:
        raise TypeError("user_request 必须是不超过8000字符的文本")
    threshold = params.get("low_margin_threshold", 0.15)
    if isinstance(threshold, bool) or not isinstance(threshold, (int, float)):
        raise TypeError("low_margin_threshold 必须是数字")
    if not 0 <= float(threshold) <= 1:
        raise ValueError("low_margin_threshold 必须在0到1之间")


@dataclass(frozen=True)
class EnterpriseDiagnosisResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def _blank(frame: pd.DataFrame, value: Any = pd.NA) -> pd.Series:
    return pd.Series(value, index=frame.index)


def _column(frame: pd.DataFrame, key: str, *, required: bool = False) -> pd.Series:
    name = _find(frame, _ALIASES[key])
    if name is None:
        if required:
            raise ValueError(f"缺少可识别字段：{key}")
        return _blank(frame)
    return frame[name]


def _text(series: pd.Series) -> pd.Series:
    return series.astype("string").fillna("").str.replace(r"\s+", " ", regex=True).str.strip()


def _numeric(series: pd.Series, label: str) -> pd.Series:
    values = pd.to_numeric(series.astype("string").str.replace(r"[¥￥元,，%％\s]", "", regex=True), errors="coerce")
    available = series.notna() & series.astype("string").str.strip().ne("")
    if available.any() and values[available].notna().mean() < 0.8:
        raise ValueError(f"字段“{label}”无法可靠转换为数值")
    return values.astype("float64")


def _optional_numeric(frame: pd.DataFrame, key: str, label: str) -> pd.Series:
    source = _column(frame, key)
    return _numeric(source, label) if source.notna().any() else pd.Series(float("nan"), index=frame.index)


def _divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    return numerator.div(denominator.where(denominator.ne(0)))


def _ratio(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else float("nan")


def _month(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce").dt.to_period("M").astype("string")


def _rate(value: Any) -> float:
    if value is None or pd.isna(value):
        return float("nan")
    text = str(value).strip()
    parsed = pd.to_numeric(text.replace("%", "").replace("％", ""), errors="coerce")
    return (
        float(parsed / 100)
        if pd.notna(parsed) and ("%" in text or "％" in text)
        else float(parsed)
        if pd.notna(parsed)
        else float("nan")
    )


def _flag(series: pd.Series, tokens: Sequence[str]) -> pd.Series:
    return _text(series).str.contains("|".join(re.escape(token) for token in tokens), regex=True, na=False)


_RISK_ORDER = {"高": 3, "中": 2, "低": 1, "未知/待核验": 0}


def _risk_level(value: Any) -> str:
    """Normalize business risk labels without treating missing data as safe."""

    if value is None or pd.isna(value) or not str(value).strip():
        return "未知/待核验"
    text = _normalise(value)
    if any(token in text for token in ("高", "严重", "危险", "high", "red")):
        return "高"
    if any(token in text for token in ("中", "一般", "关注", "medium", "moderate", "yellow")):
        return "中"
    if any(token in text for token in ("低", "正常", "良好", "low", "green")):
        return "低"
    return "未知/待核验"


def _combined_risk(source_level: Any, transaction_level: Any, *, master_matched: bool) -> str:
    source = _risk_level(source_level)
    transaction = _risk_level(transaction_level)
    recognised = [level for level in (source, transaction) if level != "未知/待核验"]
    if not master_matched and not any(level == "高" for level in recognised):
        return "未知/待核验"
    return max(recognised, key=_RISK_ORDER.get) if recognised else "未知/待核验"


def _coverage(left: pd.Series, right: pd.Series) -> tuple[float, float]:
    left_text = _text(left)
    right_values = set(_text(right)) - {""}
    valid = left_text.ne("")
    row = float(left_text[valid].isin(right_values).mean()) if valid.any() else float("nan")
    unique = set(left_text[valid])
    return row, len(unique & right_values) / len(unique) if unique else float("nan")


def _aliases(left: pd.Series, right: pd.Series) -> list[tuple[str, str, float]]:
    left_values = sorted(set(_text(left)) - {""})
    right_values = sorted(set(_text(right)) - {""})
    exact = set(right_values)
    result = []
    for value in left_values:
        if value in exact:
            continue
        scored = []
        for target in right_values:
            a, b = _normalise(value), _normalise(target)
            score = SequenceMatcher(None, a, b).ratio()
            if a in b or b in a:
                score = max(score, min(len(a), len(b)) / max(len(a), len(b)))
            scored.append((score, target))
        if scored and max(scored)[0] >= 0.6:
            score, target = max(scored)
            result.append((value, target, score))
    return result


def _source_name(value: str) -> str:
    return str(value).rsplit("::", 1)[-1]


def _fact(
    metric: str, value: Any, unit: str, source: str, definition: str, meaning: str, confidence: str = "高"
) -> dict[str, Any]:
    if isinstance(value, float) and not math.isfinite(value):
        value = "未提供"
    return {
        "指标": metric,
        "结果": value,
        "单位": unit,
        "数据来源": source,
        "数据口径": definition,
        "管理含义": meaning,
        "结论类型": "事实",
        "置信度": confidence,
    }


def build_enterprise_diagnosis_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    user_request: str = "",
    low_margin_threshold: float = 0.15,
) -> EnterpriseDiagnosisResult:
    validate_enterprise_diagnosis_params(
        {"source_names": source_names, "user_request": user_request, "low_margin_threshold": low_margin_threshold}
    )
    if len(frames) != len(source_names):
        raise ValueError("source_names 数量必须与输入表数量一致")
    from .restaurant_report import can_build_restaurant_diagnosis_report, build_restaurant_diagnosis_report
    if can_build_restaurant_diagnosis_report(frames, source_names):
        return build_restaurant_diagnosis_report(
            frames, source_names=source_names, user_request=user_request,
            low_margin_threshold=low_margin_threshold,
        )
    from .ecommerce_report import build_ecommerce_diagnosis_report, can_build_ecommerce_diagnosis_report

    if can_build_ecommerce_diagnosis_report(frames, source_names):
        return build_ecommerce_diagnosis_report(
            frames,
            source_names=source_names,
            user_request=user_request,
        )
    roles = infer_enterprise_table_roles(frames)
    sources = {role: _source_name(source_names[index]) for role, index in roles.items()}
    raw_tx = frames[roles["transactions"]].copy()
    raw_customers = frames[roles["customers"]].copy()
    raw_perf = frames[roles["performance"]].copy()
    raw_expenses = frames[roles["expenses"]].copy()
    raw_inventory = frames[roles["inventory"]].copy()
    raw_production = frames[roles["production_cost"]].copy() if "production_cost" in roles else None

    tx = pd.DataFrame(
        {
            "日期": pd.to_datetime(_column(raw_tx, "date", required=True), errors="coerce"),
            "单据号": _text(_column(raw_tx, "document")),
            "业务类型": _text(_column(raw_tx, "type")),
            "部门": _text(_column(raw_tx, "department")),
            "客户": _text(_column(raw_tx, "customer", required=True)),
            "产品": _text(_column(raw_tx, "product")),
            "负责人": _text(_column(raw_tx, "owner")),
            "原始订单金额": _numeric(_column(raw_tx, "revenue", required=True), "销售额"),
            "业务成本": _optional_numeric(raw_tx, "cost", "业务成本"),
            "付款/订单状态": _text(_column(raw_tx, "payment")),
        }
    )
    if tx["单据号"].eq("").all():
        tx["单据号"] = [f"ROW-{index + 2}" for index in range(len(tx))]
    tx = tx.drop_duplicates("单据号", keep="first").reset_index(drop=True)
    tx["月份"] = tx["日期"].dt.to_period("M").astype("string")
    tx["退款标记"] = _flag(tx["付款/订单状态"], ("退款", "退货", "已退"))
    tx["逾期标记"] = _flag(tx["付款/订单状态"], ("逾期", "超期"))
    tx["未回款标记"] = _flag(tx["付款/订单状态"], ("未回款", "未付款", "未结清", "部分回款"))
    tx["管理口径收入_排除退款"] = tx["原始订单金额"].where(~tx["退款标记"], 0.0)
    tx["情景收入_退款冲减"] = tx["原始订单金额"].where(~tx["退款标记"], -tx["原始订单金额"].abs())
    tx["毛利_流水口径"] = tx["原始订单金额"] - tx["业务成本"]

    customers = pd.DataFrame(
        {
            "客户": _text(_column(raw_customers, "customer", required=True)),
            "行业": _text(_column(raw_customers, "industry")),
            "合作时间": _text(_column(raw_customers, "years")),
            "满意度": _optional_numeric(raw_customers, "satisfaction", "满意度"),
            "信用等级": _text(_column(raw_customers, "credit")),
            "回款周期": _text(_column(raw_customers, "cycle")),
            "源表风险": _text(_column(raw_customers, "risk")),
        }
    ).drop_duplicates("客户")
    tx = tx.merge(customers, on="客户", how="left", validate="many_to_one")

    performance = pd.DataFrame(
        {
            "负责人": _text(_column(raw_perf, "owner", required=True)),
            "绩效表销售额": _numeric(_column(raw_perf, "performance_sales", required=True), "绩效销售额"),
            "绩效表毛利": _optional_numeric(raw_perf, "gross_profit", "毛利"),
            "回款金额": _optional_numeric(raw_perf, "collection", "回款金额"),
            "目标完成率": _column(raw_perf, "target").map(_rate),
            "客户评分": _optional_numeric(raw_perf, "score", "客户评分"),
            "投诉次数": _optional_numeric(raw_perf, "complaints", "投诉次数"),
        }
    ).drop_duplicates("负责人")
    performance["回款率"] = _divide(performance["回款金额"], performance["绩效表销售额"])
    performance["毛利率"] = _divide(performance["绩效表毛利"], performance["绩效表销售额"])

    expenses = pd.DataFrame(
        {
            "月份": _month(_column(raw_expenses, "month", required=True)),
            "费用类别": _text(_column(raw_expenses, "expense_category", required=True)),
            "费用金额": _numeric(_column(raw_expenses, "amount", required=True), "费用金额"),
        }
    )
    production = pd.DataFrame()
    if raw_production is not None:
        production = pd.DataFrame(
            {
                "月份": _month(_column(raw_production, "month", required=True)),
                "产品": _text(_column(raw_production, "product", required=True)),
                "材料成本": _optional_numeric(raw_production, "material_cost", "材料成本").fillna(0),
                "人工成本": _optional_numeric(raw_production, "labor_cost", "人工成本").fillna(0),
                "制造费用": _optional_numeric(raw_production, "overhead", "制造费用").fillna(0),
                "产量": _optional_numeric(raw_production, "output", "产量"),
            }
        )
        production["生产成本"] = production[["材料成本", "人工成本", "制造费用"]].sum(axis=1)
        production["单位成本"] = _divide(production["生产成本"], production["产量"])

    inventory = pd.DataFrame(
        {
            "产品": _text(_column(raw_inventory, "product", required=True)),
            "库存数量": _numeric(_column(raw_inventory, "inventory_quantity", required=True), "库存数量"),
            "库存金额": _numeric(_column(raw_inventory, "inventory_amount", required=True), "库存金额"),
            "月均销量": _numeric(_column(raw_inventory, "monthly_sales", required=True), "月均销量"),
            "供应周期_天": _optional_numeric(raw_inventory, "lead_time", "供应周期"),
            "源表状态": _text(_column(raw_inventory, "inventory_status")),
        }
    )
    inventory["库存月数"] = _divide(inventory["库存数量"], inventory["月均销量"])
    inventory["供应周期_月"] = inventory["供应周期_天"] / 30
    inventory["覆盖/供应周期倍数"] = _divide(inventory["库存月数"], inventory["供应周期_月"])
    total_inventory = float(inventory["库存金额"].sum())
    inventory["库存金额占比"] = inventory["库存金额"] / total_inventory if total_inventory else 0
    inventory["系统诊断"] = "暂未触发"
    inventory.loc[inventory["库存月数"].le(0.5), "系统诊断"] = "缺货线索（待核验）"
    inventory.loc[inventory["库存月数"].gt(6) | inventory["覆盖/供应周期倍数"].gt(3), "系统诊断"] = (
        "库存偏高线索（待核验）"
    )
    inventory["诊断状态"] = inventory["系统诊断"].map(
        {
            "库存偏高线索（待核验）": "偏高线索",
            "缺货线索（待核验）": "缺货线索",
            "暂未触发": "暂未触发",
        }
    )
    inventory["状态一致性"] = inventory.apply(
        lambda row: "待核验" if row["源表状态"] in {"正常", "良好"} and "偏高" in row["系统诊断"] else "未见明显冲突",
        axis=1,
    )
    inventory["建议行动"] = inventory["系统诊断"].map(
        {
            "库存偏高线索（待核验）": "先核实安全库存、季节性和在途订单，再决定减缓补货或消化",
            "缺货线索（待核验）": "核实在途、交期和安全库存后再补货",
            "暂未触发": "按月监控周转偏差",
        }
    )
    inventory["人工核验边界"] = "缺少安全库存、目标库存天数和季节性，不自动给出精确采购量"

    raw_sales = float(tx["原始订单金额"].sum())
    ex_refund = float(tx["管理口径收入_排除退款"].sum())
    negative_refund = float(tx["情景收入_退款冲减"].sum())
    refund_amount = float(tx.loc[tx["退款标记"], "原始订单金额"].abs().sum())
    risk_mask = tx["逾期标记"] | tx["未回款标记"]
    risk_exposure = float(tx.loc[risk_mask, "原始订单金额"].sum())
    perf_sales = float(performance["绩效表销售额"].sum())
    total_expense = float(expenses["费用金额"].sum())
    collection = float(performance["回款金额"].sum()) if performance["回款金额"].notna().any() else float("nan")
    collection_rate = _ratio(collection, perf_sales) if not math.isnan(collection) else float("nan")
    production_total = float(production["生产成本"].sum()) if not production.empty else float("nan")
    if performance["绩效表毛利"].notna().any():
        gross_profit = float(performance["绩效表毛利"].sum())
        profit_source = sources["performance"]
        profit_basis = "人员绩效表毛利汇总"
        profit_revenue = perf_sales
        profit_confidence = "中"
    elif tx["业务成本"].notna().any():
        gross_profit = float(tx["毛利_流水口径"].sum())
        profit_source = sources["transactions"]
        profit_basis = "销售流水收入-业务成本"
        profit_revenue = raw_sales
        profit_confidence = "高"
    else:
        gross_profit = float("nan")
        profit_source = sources["transactions"]
        profit_basis = "缺少可勾稽的毛利/销售成本"
        profit_revenue = raw_sales
        profit_confidence = "低"
    profit_metric_label = "绩效口径毛利" if profit_basis == "人员绩效表毛利汇总" else "流水口径毛利"
    margin_metric_label = f"{profit_metric_label}率"
    gross_margin = _ratio(gross_profit, profit_revenue) if not math.isnan(gross_profit) else float("nan")
    operating_profit = gross_profit - total_expense if not math.isnan(gross_profit) else float("nan")
    operating_margin = _ratio(operating_profit, profit_revenue) if not math.isnan(operating_profit) else float("nan")

    monthly = tx.groupby("月份", as_index=False).agg(
        原始订单金额=("原始订单金额", "sum"),
        管理口径收入=("管理口径收入_排除退款", "sum"),
        单据数=("单据号", "nunique"),
        风险订单金额=("原始订单金额", lambda values: float(values[risk_mask.loc[values.index]].sum())),
    )
    if tx["业务成本"].notna().any():
        monthly_cost = (
            tx.groupby("月份", as_index=False)["业务成本"].sum().rename(columns={"业务成本": "业务或生产成本"})
        )
    elif not production.empty:
        monthly_cost = (
            production.groupby("月份", as_index=False)["生产成本"].sum().rename(columns={"生产成本": "业务或生产成本"})
        )
    else:
        monthly_cost = pd.DataFrame(columns=["月份", "业务或生产成本"])
    monthly = (
        monthly.merge(monthly_cost, on="月份", how="outer")
        .merge(expenses.groupby("月份", as_index=False)["费用金额"].sum(), on="月份", how="outer")
        .sort_values("月份")
    )
    for column in monthly.columns:
        if column != "月份":
            monthly[column] = pd.to_numeric(monthly[column], errors="coerce")
    for column in ("原始订单金额", "管理口径收入", "单据数", "风险订单金额", "费用金额"):
        monthly[column] = monthly[column].fillna(0)
    monthly["估算经营贡献"] = monthly["管理口径收入"] - monthly["业务或生产成本"] - monthly["费用金额"]
    monthly["收入环比"] = monthly["管理口径收入"].pct_change()
    monthly["成本费用率"] = _divide(monthly["业务或生产成本"] + monthly["费用金额"], monthly["管理口径收入"])
    monthly["计算边界"] = "月度贡献仅用于趋势；生产成本未必等于同期销售成本"

    customer_analysis = tx.groupby("客户", as_index=False).agg(
        原始订单金额=("原始订单金额", "sum"),
        管理口径收入=("管理口径收入_排除退款", "sum"),
        单据数=("单据号", "nunique"),
        退款涉及金额=("原始订单金额", lambda values: float(values[tx.loc[values.index, "退款标记"]].abs().sum())),
        风险订单金额=("原始订单金额", lambda values: float(values[risk_mask.loc[values.index]].sum())),
    )
    customer_analysis["收入占比"] = customer_analysis["管理口径收入"] / ex_refund if ex_refund else 0
    customer_analysis = customer_analysis.merge(customers, on="客户", how="left")
    customer_master_keys = set(customers["客户"].dropna().astype(str)) - {""}
    customer_analysis["主数据匹配"] = customer_analysis["客户"].astype(str).isin(customer_master_keys)
    customer_analysis["源业务风险"] = customer_analysis["源表风险"].map(_risk_level)
    customer_analysis["交易风险"] = "低"
    customer_analysis.loc[customer_analysis["退款涉及金额"].gt(0), "交易风险"] = "中"
    customer_analysis.loc[
        customer_analysis["风险订单金额"].gt(0) | customer_analysis["满意度"].lt(4),
        "交易风险",
    ] = "高"
    customer_analysis["综合风险"] = customer_analysis.apply(
        lambda row: _combined_risk(
            row["源业务风险"],
            row["交易风险"],
            master_matched=bool(row["主数据匹配"]),
        ),
        axis=1,
    )
    # Backward-compatible name used by existing exports and API consumers.
    customer_analysis["风险等级"] = customer_analysis["综合风险"]
    customer_analysis["数据证据"] = customer_analysis.apply(
        lambda row: (
            f"收入{row['管理口径收入']:,.0f}元，占比{row['收入占比']:.1%}，风险订单{row['风险订单金额']:,.0f}元"
        ),
        axis=1,
    )

    def _customer_risk_reason(row: pd.Series) -> str:
        reasons: list[str] = []
        risk_amount = pd.to_numeric(pd.Series([row.get("风险订单金额")]), errors="coerce").iloc[0]
        if pd.notna(risk_amount) and float(risk_amount) != 0:
            reasons.append("存在逾期/未回款订单")
        if pd.notna(row.get("源表风险")) and str(row.get("源表风险")).strip() == "高":
            reasons.append("源表标记高风险")
        elif row.get("源业务风险") == "中":
            reasons.append("源表标记中风险")
        refund = pd.to_numeric(pd.Series([row.get("退款涉及金额")]), errors="coerce").iloc[0]
        if pd.notna(refund) and float(refund) > 0:
            reasons.append(f"退款订单{float(refund):,.0f}元需复核")
        if pd.notna(row.get("满意度")) and float(row.get("满意度")) < 4:
            reasons.append("满意度低于4分")
        if not bool(row.get("主数据匹配")):
            reasons.append("客户主数据未匹配")
        return "；".join(reasons) or "未触发"

    customer_analysis["主要风险"] = customer_analysis.apply(_customer_risk_reason, axis=1)
    customer_analysis["__risk"] = customer_analysis["风险等级"].map({"高": 0, "中": 1, "未知/待核验": 2, "低": 3})
    customer_analysis = (
        customer_analysis.sort_values(["__risk", "管理口径收入"], ascending=[True, False])
        .drop(columns="__risk")
        .reset_index(drop=True)
    )
    customer_analysis["收入排名"] = (
        customer_analysis["管理口径收入"].rank(method="min", ascending=False).astype("Int64")
    )

    salesperson = tx.groupby("负责人", as_index=False).agg(
        流水净收入=("管理口径收入_排除退款", "sum"),
        流水原始订单额=("原始订单金额", "sum"),
        订单级成本=("业务成本", lambda values: values.sum(min_count=1)),
        风险订单金额=("原始订单金额", lambda values: float(values[risk_mask.loc[values.index]].sum())),
    )
    salesperson = salesperson.merge(performance, on="负责人", how="outer")
    salesperson["参考毛利"] = salesperson["绩效表毛利"].where(
        salesperson["绩效表毛利"].notna(), salesperson["流水原始订单额"] - salesperson["订单级成本"]
    )
    salesperson["参考毛利率"] = _divide(salesperson["参考毛利"], salesperson["绩效表销售额"])
    salesperson["毛利口径"] = salesperson["绩效表毛利"].notna().map({True: "绩效口径", False: "订单收入-订单级成本"})
    salesperson["管理诊断"] = salesperson.apply(
        lambda row: (
            "规模高但回款风险高"
            if pd.notna(row["回款率"])
            and row["回款率"] < 0.6
            and row["绩效表销售额"] == salesperson["绩效表销售额"].max()
            else (
                "重点关注"
                if (pd.notna(row["客户评分"]) and row["客户评分"] < 4)
                or (pd.notna(row["投诉次数"]) and row["投诉次数"] >= 2)
                else "暂未触发"
            )
        ),
        axis=1,
    )
    salesperson["建议动作"] = salesperson["管理诊断"].map(
        {
            "规模高但回款风险高": "业绩同时纳入毛利、回款和客户风险",
            "重点关注": "建立客户沟通、投诉和催收跟进清单",
            "暂未触发": "复制高质量成交和回款方法",
        }
    )
    salesperson = salesperson.sort_values("绩效表销售额", ascending=False).reset_index(drop=True)
    salesperson["收入排名"] = salesperson["绩效表销售额"].rank(method="min", ascending=False).astype("Int64")

    expense_category = (
        expenses.groupby("费用类别", as_index=False)["费用金额"].sum().sort_values("费用金额", ascending=False)
    )
    expense_category["占期间费用比"] = expense_category["费用金额"] / total_expense if total_expense else 0
    cost_rows = []
    if not production.empty:
        for product, group in production.groupby("产品"):
            row = {"分析层级": "生产成本-产品", "对象": product}
            row.update(
                {
                    column: float(group[column].sum())
                    for column in ("材料成本", "人工成本", "制造费用", "生产成本", "产量")
                }
            )
            row["单位成本"] = _ratio(row["生产成本"], row["产量"])
            row["费用金额"] = pd.NA
            row["占比"] = _ratio(row["生产成本"], production_total)
            row["经营解释"] = "不同产品单位可能不一致，不对产量做跨产品平均"
            cost_rows.append(row)
    elif tx["业务成本"].notna().any():
        cost_rows.append(
            {
                "分析层级": "业务成本",
                "对象": "全部流水",
                "生产成本": float(tx["业务成本"].sum()),
                "费用金额": pd.NA,
                "占比": 1.0,
                "经营解释": "与流水收入直接对应",
            }
        )
    for _, row in expense_category.iterrows():
        cost_rows.append(
            {
                "分析层级": "期间费用-类别",
                "对象": row["费用类别"],
                "生产成本": pd.NA,
                "费用金额": row["费用金额"],
                "占比": row["占期间费用比"],
                "经营解释": "费用增长不等于效率下降，缺少归因数据时仅提示核验",
            }
        )
    cost_analysis = pd.DataFrame(cost_rows)

    top1 = float(customer_analysis["收入占比"].max())
    top3 = (
        float(customer_analysis.nlargest(3, "管理口径收入")["管理口径收入"].sum() / ex_refund)
        if ex_refund
        else float("nan")
    )
    c_row, c_unique = _coverage(
        _column(raw_tx, "customer", required=True), _column(raw_customers, "customer", required=True)
    )
    cr_row, cr_unique = _coverage(
        _column(raw_customers, "customer", required=True), _column(raw_tx, "customer", required=True)
    )
    p_row, p_unique = _coverage(_column(raw_tx, "owner"), _column(raw_perf, "owner"))
    pr_row, pr_unique = _coverage(_column(raw_perf, "owner"), _column(raw_tx, "owner"))
    product_aliases = (
        _aliases(_column(raw_tx, "product"), _column(raw_inventory, "product", required=True))
        if _column(raw_tx, "product").notna().any()
        else []
    )
    product_unique = (
        _coverage(_column(raw_tx, "product"), _column(raw_inventory, "product", required=True))[1]
        if _column(raw_tx, "product").notna().any()
        else float("nan")
    )
    open_items = (
        int(refund_amount > 0)
        + int(c_unique < 1)
        + int(bool(product_aliases))
        + int(inventory["状态一致性"].eq("待核验").any())
    )
    confidence = "高" if open_items == 0 else "中" if open_items <= 7 else "低"

    summary_rows = [
        _fact(
            "原始订单金额",
            raw_sales,
            "元",
            sources["transactions"],
            "原值汇总，不擅自改变退款符号",
            "业务规模，不等于会计净收入",
        ),
        _fact(
            "净收入",
            raw_sales,
            "元",
            sources["transactions"],
            "兼容字段：实为原始订单金额",
            "存在退款正数时需财务确认",
            "中" if refund_amount else "高",
        ),
        _fact(
            "管理情景收入（排除退款）",
            ex_refund,
            "元",
            sources["transactions"],
            "退款行按0处理",
            "只用于管理情景",
            "中",
        ),
        _fact(
            "退款冲减情景收入",
            negative_refund,
            "元",
            sources["transactions"],
            "退款行按负数冲减",
            "敏感性情景，不替代会计凭证",
            "中",
        ),
        _fact(
            profit_metric_label,
            gross_profit,
            "元",
            profit_source,
            profit_basis,
            "优先使用已报毛利，避免生产成本重复扣减",
            profit_confidence,
        ),
        _fact(
            margin_metric_label,
            gross_margin,
            "%",
            profit_source,
            f"{profit_basis}÷对应销售额",
            "需确认绩效表毛利口径",
            profit_confidence,
        ),
        _fact("期间费用", total_expense, "元", sources["expenses"], "费用支出汇总", "与业务/生产成本分开披露"),
        _fact(
            "估算经营利润",
            operating_profit,
            "元",
            f"{profit_source}+{sources['expenses']}",
            "已报毛利-期间费用，未含未提供项",
            "不替代法定利润",
            "中",
        ),
        _fact("估算经营利润率", operating_margin, "%", "系统计算", "估算经营利润÷毛利对应销售额", "经营诊断口径", "中"),
        _fact("回款金额", collection, "元", sources["performance"], "人员绩效表汇总", "销售向现金转化", "中"),
        _fact(
            "回款率",
            collection_rate,
            "%",
            sources["performance"],
            "回款金额÷绩效表销售额",
            "不同期或含税口径需调整",
            "中",
        ),
        _fact(
            "未结清订单收入风险敞口",
            risk_exposure,
            "元",
            sources["transactions"],
            "逾期/未回款订单原始金额",
            "不等于应收余额",
            "中",
        ),
        _fact(
            "前三客户收入集中度", top3, "%", sources["transactions"], "Top3管理口径收入÷总收入", "结合信用和账期判断"
        ),
        _fact("库存总金额", total_inventory, "元", sources["inventory"], "库存表金额汇总", "资金占用规模"),
        _fact(
            "生产成本总额",
            production_total,
            "元",
            sources.get("production_cost", "未提供"),
            "材料+人工+制造费用",
            "不默认等同同期销售成本",
            "中",
        ),
        _fact(
            "高风险客户数",
            int(customer_analysis["风险等级"].eq("高").sum()),
            "个",
            sources["customers"],
            "风险、低满意度和风险订单联合标记",
            "安排授信和催收复核",
            "中",
        ),
        _fact(
            "待核验事项",
            open_items,
            "项",
            "数据审计",
            "退款、主数据、状态冲突",
            "未确认前不做不可逆处置",
            confidence,
        ),
        _fact(
            "数据可信度",
            confidence,
            "",
            "数据审计",
            "勾稽、关联覆盖和开放口径综合评估",
            "定义结论可使用的边界",
            confidence,
        ),
    ]
    diagnosis = (
        f"销售规模未转化为利润：{profit_metric_label}{gross_profit:,.0f}元低于期间费用{total_expense:,.0f}元，估算经营结果{operating_profit:,.0f}元。"
        if not math.isnan(operating_profit) and operating_profit < 0
        else "未发现经营利润为负，但仍需复核回款、客户集中和库存占压。"
    )
    summary_rows.append(
        {
            "指标": "企业目前最大的问题",
            "结果": diagnosis,
            "单位": "",
            "数据来源": "跨事实域诊断",
            "数据口径": "规模、利润、费用、回款、客户与库存联合判断",
            "管理含义": "先修复现金和利润转化，再追求规模",
            "结论类型": "诊断",
            "置信度": confidence,
        }
    )
    summary = pd.DataFrame(summary_rows)

    risks = []

    def risk(
        priority: str,
        theme: str,
        evidence: str,
        impact: str,
        cause: str,
        department: str,
        action: str,
        owner: str,
        due: str,
        metric: str,
        boundary: str,
        confidence_level: str = "中",
    ) -> None:
        risks.append(
            {
                "优先级": priority,
                "风险事项": theme,
                "问题描述": theme,
                "数据证据": evidence,
                "风险影响": impact,
                "原因判断": cause,
                "责任部门（建议）": department,
                "建议行动": action,
                "负责人（待确认）": owner,
                "完成期限（建议）": due,
                "验收指标": metric,
                "置信度": confidence_level,
                "人工核验边界": boundary,
            }
        )

    if not math.isnan(operating_profit) and operating_profit < 0:
        risk(
            "P0",
            "利润转化不足",
            f"毛利{gross_profit:,.0f}元，期间费用{total_expense:,.0f}元，估算经营结果{operating_profit:,.0f}元",
            "增长可能放大资金压力",
            "毛利不足以覆盖期间费用",
            "财务+业务",
            "按产品/客户复盘毛利，对费用建立项目归因",
            "财务负责人+业务负责人",
            "7天",
            "毛利与费用归因覆盖100%",
            "估算结果不含税费、折旧等未提供项",
        )
    if risk_exposure or (not math.isnan(collection_rate) and collection_rate < 0.8):
        risk(
            "P0"
            if risk_exposure > raw_sales * 0.2 or (not math.isnan(collection_rate) and collection_rate < 0.6)
            else "P1",
            "回款与现金转化",
            f"回款率{collection_rate:.1%}；逾期/未回款订单{risk_exposure:,.0f}元"
            if not math.isnan(collection_rate)
            else f"风险订单{risk_exposure:,.0f}元",
            "现金紧张和坏账风险上升",
            "长账期、高风险客户集中或催收机制不足",
            "销售+财务",
            "建立客户-单据级回款台账，明确责任人和承诺日",
            "销售负责人+财务应收",
            "3天",
            "风险订单责任人与承诺日覆盖100%",
            "订单金额不等于应收余额",
        )
    if refund_amount:
        risk(
            "P0" if refund_amount / raw_sales >= 0.05 else "P1",
            "退款收入口径",
            f"退款正数金额{refund_amount:,.0f}元；排除退款={ex_refund:,.0f}元，负数冲减={negative_refund:,.0f}元",
            "收入、利润和增长率可能失真",
            "状态与金额符号规则未明确",
            "财务+订单运营",
            "核对退款凭证，确认排除、冲减或另表处理",
            "财务经理",
            "2天",
            "每笔退款都有唯一口径和凭证链路",
            "系统只展示情景，不自动改原数据",
            "高",
        )
    if top1 >= 0.4:
        top_customer = customer_analysis.sort_values("收入占比", ascending=False).iloc[0]
        risk(
            "P0" if top_customer["风险等级"] == "高" else "P1",
            "客户集中",
            f"{top_customer['客户']}贡献{top_customer['收入占比']:.1%}，风险等级{top_customer['风险等级']}，风险订单{top_customer['风险订单金额']:,.0f}元",
            "客户延迟付款或流失会放大波动",
            "集中度与信用风险叠加",
            "销售+风控",
            "重新核定授信与账期，拓展中低风险客户",
            "销售总监",
            "30天",
            "按月跟踪Top1/Top3和高风险收入占比",
            "集中度本身不等于风险，需结合合同和回款",
            "高",
        )
    sales_quality = salesperson[
        salesperson["管理诊断"].eq("重点关注")
        & (salesperson["客户评分"].lt(4) | salesperson["投诉次数"].ge(2) | salesperson["回款率"].lt(0.6))
    ]
    if not sales_quality.empty:
        person = sales_quality.sort_values(["投诉次数", "风险订单金额"], ascending=[False, False]).iloc[0]
        collection_text = f"{person['回款率']:.1%}" if pd.notna(person["回款率"]) else "未提供"
        score_text = f"{person['客户评分']:.1f}" if pd.notna(person["客户评分"]) else "未提供"
        complaint_text = f"{person['投诉次数']:.0f}" if pd.notna(person["投诉次数"]) else "未提供"
        risk(
            "P1",
            "销售质量与客户体验",
            f"{person['负责人']}：回款率{collection_text}，客户评分{score_text}，投诉{complaint_text}次，风险订单{person['风险订单金额']:,.0f}元",
            "客户流失、回款延迟和口碑风险可能叠加",
            "成交承诺、客户沟通、投诉闭环或催收节奏存在薄弱环节",
            "销售+客户成功",
            "复盘重点客户投诉、承诺兑现和催收节奏，建立30天整改跟踪",
            "销售经理",
            "30天",
            "投诉闭环率100%，重点客户均有回款与服务行动记录",
            "人员结论仅用于管理复核，不自动触发奖惩或人事决定",
            "中",
        )
    overstock = inventory[inventory["系统诊断"].str.contains("偏高", na=False)]
    if not overstock.empty:
        risk(
            "P1",
            "库存资金占用线索",
            f"{'、'.join(overstock['产品'])}涉及{overstock['库存金额'].sum():,.0f}元，最高{overstock['库存月数'].max():.1f}个月",
            "占用现金并带来减值风险",
            "库存政策、需求预测或采购节奏可能不匹配",
            "供应链+销售",
            "先补齐安全库存和目标天数，再制定减缓补货和消化方案",
            "供应链负责人",
            "10天",
            "高占压SKU完成政策核验和行动清单",
            "缺少安全库存和季节性，不直接下达停采/促销",
        )
    if c_unique < 1 or product_aliases:
        risk(
            "P1",
            "主数据与表关系",
            f"订单→客户：行{c_row:.1%}/唯一值{c_unique:.1%}；产品别名：{'；'.join(f'{a}↔{b}({s:.0%})' for a, b, s in product_aliases) or '无'}",
            "风险、库存和利润分析可能漏计",
            "别名或主数据缺失",
            "数据管理+业务",
            "人工确认后建立可追溯别名映射表",
            "数据管理员",
            "14天",
            "交易事实到主数据的行覆盖100%",
            "值域相似不经人工确认不自动合并",
            "高",
        )
    actions = pd.DataFrame(risks).sort_values("优先级").reset_index(drop=True)

    def _transaction_risk_reason(row: pd.Series) -> str:
        reasons: list[str] = []
        for column, label in (
            ("退款标记", "退款口径待确认"),
            ("逾期标记", "逾期"),
            ("未回款标记", "未回款"),
        ):
            value = row.get(column)
            if pd.notna(value) and bool(value):
                reasons.append(label)
        credit = row.get("信用等级")
        if pd.isna(credit) or not str(credit).strip():
            reasons.append("客户主数据未匹配")
        source_risk = row.get("源表风险")
        if pd.notna(source_risk) and str(source_risk).strip() == "高":
            reasons.append("客户高风险")
        elif _risk_level(source_risk) == "中":
            reasons.append("客户中风险")
        satisfaction = row.get("满意度")
        if pd.notna(satisfaction) and float(satisfaction) < 4:
            reasons.append("低满意度")
        return "；".join(reasons) or "未触发"

    working = tx.copy()
    working["业务风险说明"] = working.apply(_transaction_risk_reason, axis=1)
    audits = []

    def audit(category: str, item: str, status: str, evidence: str, impact: str, action: str) -> None:
        audits.append(
            {
                "审计类别": category,
                "审计项": item,
                "状态": status,
                "数据证据/口径": evidence,
                "影响": impact,
                "建议处理": action,
            }
        )

    audit(
        "勾稽",
        "销售额跨表勾稽",
        "通过" if abs(raw_sales - perf_sales) < 0.01 else "待核验",
        f"订单={raw_sales:,.2f}；绩效={perf_sales:,.2f}；差额={raw_sales - perf_sales:,.2f}",
        "影响回款率和利润率",
        "核对时间、含税和退款口径",
    )
    audit(
        "关系",
        "订单→客户主数据",
        "通过" if c_unique == 1 else "待核验",
        f"行覆盖={c_row:.1%}；唯一客户覆盖={c_unique:.1%}",
        "未匹配客户缺少风险属性",
        "补充客户主数据或别名",
    )
    audit(
        "关系",
        "客户主数据→订单",
        "通过" if cr_unique == 1 else "关注",
        f"行覆盖={cr_row:.1%}；唯一值覆盖={cr_unique:.1%}",
        "发现无交易客户，不代表主数据错误",
        "按客户生命周期复核",
    )
    audit(
        "关系",
        "销售人员↔绩效人员",
        "通过" if p_unique == 1 and pr_unique == 1 else "待核验",
        f"订单→绩效：行{p_row:.1%}/唯一{p_unique:.1%}；绩效→订单：行{pr_row:.1%}/唯一{pr_unique:.1%}",
        "影响业绩与回款归属",
        "维护组织主数据",
    )
    if not math.isnan(product_unique):
        audit(
            "关系",
            "销售产品↔库存产品",
            "通过" if product_unique == 1 else "待核验",
            f"唯一值直接覆盖={product_unique:.1%}",
            "影响库存与销售联动",
            "确认别名后再建映射",
        )
    for a, b, score in product_aliases:
        audit(
            "主数据映射",
            f"{a}↔{b}",
            "待人工确认",
            f"值域相似度={score:.1%}",
            "未映射会拆分销售和库存",
            "确认后写入别名表",
        )
    audit(
        "口径",
        "退款处理",
        "通过" if refund_amount == 0 else "待财务确认",
        f"退款正数={refund_amount:,.2f}；排除收入={ex_refund:,.2f}；冲减收入={negative_refund:,.2f}",
        "影响收入、利润和增长率",
        "核对凭证并确认唯一口径",
    )
    audit(
        "口径",
        "经营利润",
        "待财务确认",
        f"估算结果={operating_profit:,.2f}；基础={profit_basis}",
        "不能等同法定利润",
        "与财务报表勾稽后对外使用",
    )
    audit(
        "边界",
        "生产成本与销售成本",
        "待核验" if not production.empty else "不适用",
        "生产成本按生产月/产品记录，未提供销售出库成本",
        "不得直接当作销售成本",
        "补充产品成本结转或出库成本",
    )
    audit(
        "边界",
        "库存阈值",
        "待业务确认",
        "暂以可售月数>6或超供应周期3倍作趋势线索",
        "可能误报",
        "补齐安全库存、目标天数和季节性",
    )
    audit_frame = pd.DataFrame(audits)
    # The overview count is reconciled to the audit sheet instead of using an
    # independently maintained approximation.
    open_items = int(audit_frame["状态"].astype(str).str.startswith("待").sum())
    confidence = "高" if open_items == 0 else "中" if open_items <= 7 else "低"
    summary.loc[summary["指标"].eq("待核验事项"), ["结果", "数据口径", "置信度"]] = [
        open_items,
        "与数据口径与验收页中所有‘待…’状态逐项勾稽",
        confidence,
    ]
    summary.loc[summary["指标"].eq("数据可信度"), ["结果", "置信度"]] = [confidence, confidence]
    summary.loc[summary["指标"].eq("企业目前最大的问题"), "置信度"] = confidence

    dashboard = monthly[["月份", "原始订单金额", "管理口径收入", "业务或生产成本", "费用金额", "估算经营贡献"]].copy()
    chart_frames = (
        (
            "客户",
            customer_analysis.sort_values("管理口径收入", ascending=False).head(8),
            ["客户", "管理口径收入", "风险订单金额"],
        ),
        (
            "销售",
            salesperson.sort_values("绩效表销售额", ascending=False).head(8),
            ["负责人", "绩效表销售额", "回款金额"],
        ),
        ("库存", inventory.sort_values("库存月数", ascending=False).head(8), ["产品", "库存月数", "库存金额"]),
    )
    max_len = max(len(dashboard), *(len(frame) for _, frame, _ in chart_frames), 1)
    dashboard = dashboard.reindex(range(max_len))
    for prefix, frame, columns in chart_frames:
        aligned = frame[columns].reset_index(drop=True).reindex(range(max_len))
        for column in columns:
            dashboard[f"{prefix}_{column}"] = aligned[column]

    dashboard.loc[0, "KPI_销售规模"] = raw_sales
    dashboard.loc[0, "KPI_回款率"] = collection_rate
    dashboard.loc[0, "KPI_毛利标题"] = margin_metric_label
    dashboard.loc[0, "KPI_毛利率"] = gross_margin
    dashboard.loc[0, "KPI_估算经营结果"] = operating_profit
    dashboard.loc[0, "KPI_风险订单"] = risk_exposure
    dashboard.loc[0, "KPI_库存金额"] = total_inventory
    dashboard.loc[0, "核心诊断"] = diagnosis
    for index, (_, action_row) in enumerate(actions.head(3).iterrows(), start=1):
        dashboard.loc[0, f"风险卡{index}_标题"] = f"{action_row['优先级']}｜{action_row['风险事项']}"
        dashboard.loc[0, f"风险卡{index}_证据"] = action_row["数据证据"]
        dashboard.loc[0, f"风险卡{index}_行动"] = action_row["建议行动"]

    customer_output = customer_analysis[
        [
            "客户",
            "管理口径收入",
            "收入占比",
            "风险订单金额",
            "退款涉及金额",
            "源业务风险",
            "交易风险",
            "综合风险",
            "主要风险",
            "信用等级",
            "满意度",
            "回款周期",
            "行业",
            "合作时间",
            "单据数",
            "原始订单金额",
            "数据证据",
            "主数据匹配",
            "收入排名",
        ]
    ].copy()
    salesperson_output = salesperson[
        [
            "负责人",
            "绩效表销售额",
            "回款金额",
            "回款率",
            "参考毛利",
            "参考毛利率",
            "毛利口径",
            "风险订单金额",
            "目标完成率",
            "客户评分",
            "投诉次数",
            "管理诊断",
            "建议动作",
            "收入排名",
            "流水净收入",
            "流水原始订单额",
            "订单级成本",
        ]
    ].copy()

    outputs = {
        "管理层诊断总览": summary,
        "利润驱动分析": monthly,
        "客户与回款风险": customer_output,
        "销售团队诊断": salesperson_output,
        "成本费用分析": cost_analysis,
        "库存风险分析": inventory,
        "风险行动计划": actions,
        "诊断底稿": working,
        "数据口径与验收": audit_frame,
        "经营诊断看板": dashboard,
    }
    for frame in outputs.values():
        frame.attrs["toolbox_report_kind"] = "enterprise_diagnosis_report"

    def finite_or_none(value: Any) -> Any:
        return value if not isinstance(value, float) or math.isfinite(value) else None

    report = {
        "net_revenue": raw_sales,
        "management_revenue_excluding_refunds": ex_refund,
        "gross_profit": finite_or_none(gross_profit),
        "gross_margin": finite_or_none(gross_margin),
        "operating_expense": total_expense,
        "estimated_operating_profit": finite_or_none(operating_profit),
        "estimated_operating_margin": finite_or_none(operating_margin),
        "collection_amount": finite_or_none(collection),
        "collection_rate": finite_or_none(collection_rate),
        "collection_risk_exposure": risk_exposure,
        "top1_customer_concentration": finite_or_none(top1),
        "top3_customer_concentration": finite_or_none(top3),
        "inventory_value": total_inventory,
        "production_cost": finite_or_none(production_total),
        "data_confidence": confidence,
        "open_definition_count": open_items,
        "sheet_count": len(outputs),
        "risk_count": len(actions),
        "source_tables": [_source_name(name) for name in source_names],
        "fact_domains": sorted(roles),
    }
    return EnterpriseDiagnosisResult(outputs=outputs, report=report)


__all__ = [
    "EnterpriseDiagnosisResult",
    "build_enterprise_diagnosis_report",
    "can_build_enterprise_diagnosis_report",
    "infer_enterprise_table_roles",
    "validate_enterprise_diagnosis_params",
]
