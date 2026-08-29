"""Deterministic procurement, sales and inventory management reporting.

DeepSeek may recognise that a request is an inventory-management order, but
all table-role detection, cleansing, stock arithmetic, thresholds and report
outputs are produced locally from an allow-listed workflow.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import timedelta
import math
import re
from typing import Any

import pandas as pd


_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:]+")


def _normalise_name(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


def _find_column(frame: pd.DataFrame, aliases: Sequence[str]) -> str | None:
    normalised = {_normalise_name(column): str(column) for column in frame.columns}
    for alias in aliases:
        key = _normalise_name(alias)
        if key in normalised:
            return normalised[key]
    candidates = {
        column
        for key, column in normalised.items()
        for alias in aliases
        if len(_normalise_name(alias)) >= 2 and _normalise_name(alias) in key
    }
    return next(iter(candidates)) if len(candidates) == 1 else None


def _has_columns(frame: pd.DataFrame, groups: Sequence[Sequence[str]]) -> bool:
    return all(_find_column(frame, aliases) is not None for aliases in groups)


def infer_inventory_table_roles(frames: Sequence[pd.DataFrame]) -> dict[str, int]:
    """Identify the five operational table roles without relying on sheet names."""

    if not isinstance(frames, Sequence) or isinstance(frames, (str, bytes)):
        raise TypeError("库存经营分析输入必须是表格数组")
    roles: dict[str, int] = {}
    signatures = {
        "products": (("商品编码", "产品编码", "SKU"), ("安全库存",), ("目标库存天数", "目标周转天数")),
        "opening": (("商品编码", "产品编码", "SKU"), ("期初库存", "期初数量"), ("已锁定", "锁定库存")),
        "purchases": (("入库单号", "采购单号"), ("商品编码", "产品编码", "SKU"), ("入库数量", "采购数量")),
        "sales": (("出库单号", "销售单号"), ("商品编码", "产品编码", "SKU"), ("销售额", "销售金额")),
        "adjustments": (("商品编码", "产品编码", "SKU"), ("调整数量",), ("类型", "调整类型")),
    }
    for role, groups in signatures.items():
        matches = [index for index, frame in enumerate(frames) if isinstance(frame, pd.DataFrame) and _has_columns(frame, groups)]
        if len(matches) != 1:
            label = {
                "products": "商品资料",
                "opening": "期初库存",
                "purchases": "采购入库",
                "sales": "销售出库",
                "adjustments": "库存调整",
            }[role]
            raise ValueError(f"无法唯一识别“{label}”表")
        roles[role] = matches[0]
    return roles


def can_build_inventory_report(frames: Sequence[pd.DataFrame]) -> bool:
    try:
        infer_inventory_table_roles(frames)
    except (TypeError, ValueError):
        return False
    return True


def validate_inventory_report_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("库存经营报告参数必须是对象")
    source_names = params.get("source_names")
    if not isinstance(source_names, (list, tuple)) or not source_names or not all(
        isinstance(item, str) and item.strip() for item in source_names
    ):
        raise TypeError("source_names 必须是非空字符串数组")
    recent_days = params.get("recent_days", 30)
    if isinstance(recent_days, bool) or not isinstance(recent_days, int):
        raise TypeError("recent_days 必须是整数")
    if not 7 <= recent_days <= 365:
        raise ValueError("recent_days 必须在 7 到 365 之间")
    multiplier = params.get("overstock_multiplier", 1.5)
    if isinstance(multiplier, bool) or not isinstance(multiplier, (int, float)):
        raise TypeError("overstock_multiplier 必须是数字")
    if not 1 <= float(multiplier) <= 10:
        raise ValueError("overstock_multiplier 必须在 1 到 10 之间")


def _clean_text(value: Any) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _clean_code(value: Any) -> str:
    return re.sub(r"\s+", "", _clean_text(value)).upper()


def _source_label(value: str) -> str:
    text = _clean_text(value)
    return text.rsplit("::", 1)[-1] if "::" in text else text


def _parse_number(value: Any) -> float:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return float("nan")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = _clean_text(value)
    if not text:
        return float("nan")
    negative = text.startswith("(") and text.endswith(")")
    cleaned = re.sub(r"[¥￥元件个台套箱分,%，,\s()]", "", text)
    try:
        number = float(cleaned)
    except ValueError:
        return float("nan")
    return -number if negative else number


def _parse_date(value: Any) -> pd.Timestamp | pd.NaT:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if 20_000 <= number <= 80_000:
            return (pd.Timestamp("1899-12-30") + timedelta(days=number)).normalize()
    text = _clean_text(value)
    if not text:
        return pd.NaT
    text = text.replace("年", "-").replace("月", "-").replace("日", "").replace(".", "-").replace("/", "-")
    text = re.sub(r"-+", "-", text).strip("-")
    match = re.fullmatch(r"(\d{1,2})-(\d{1,2})-(\d{4})", text)
    if match:
        month, day, year = (int(item) for item in match.groups())
        return pd.Timestamp(year=year, month=month, day=day)
    parsed = pd.to_datetime(text, errors="coerce", yearfirst=True)
    return pd.NaT if pd.isna(parsed) else pd.Timestamp(parsed).normalize()


def _frame_columns(frame: pd.DataFrame, mapping: Mapping[str, Sequence[str]]) -> dict[str, str | None]:
    return {key: _find_column(frame, aliases) for key, aliases in mapping.items()}


def _deduplicate_documents(frame: pd.DataFrame, key: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    work = frame.copy(deep=True).reset_index(drop=True)
    work["__source_order"] = range(len(work))
    duplicates = work[key].astype("string").ne("") & work.duplicated(key, keep="first")
    return work.loc[~duplicates].copy(), work.loc[duplicates].copy()


def _infer_as_of_date(frames: Sequence[pd.DataFrame], fallback_dates: Sequence[pd.Series]) -> pd.Timestamp:
    for frame in frames:
        for value in frame.astype("string").fillna("").to_numpy().ravel().tolist():
            text = str(value)
            if "截止" not in text:
                continue
            match = re.search(r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})", text)
            if match:
                return pd.Timestamp(year=int(match.group(1)), month=int(match.group(2)), day=int(match.group(3)))
    available = pd.concat([series.dropna() for series in fallback_dates if not series.dropna().empty], ignore_index=True)
    if available.empty:
        raise ValueError("无法识别分析截止日期")
    return pd.Timestamp(available.max()).normalize()


@dataclass(frozen=True)
class InventoryReportResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def build_inventory_management_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    recent_days: int = 30,
    overstock_multiplier: float = 1.5,
) -> InventoryReportResult:
    """Clean the movement tables and produce an auditable inventory report."""

    params = {
        "source_names": source_names,
        "recent_days": recent_days,
        "overstock_multiplier": overstock_multiplier,
    }
    validate_inventory_report_params(params)
    if len(frames) != len(source_names):
        raise ValueError("source_names 数量必须与输入表数量一致")
    roles = infer_inventory_table_roles(frames)
    names = list(source_names)

    product_frame = frames[roles["products"]].copy(deep=True)
    opening_frame = frames[roles["opening"]].copy(deep=True)
    purchase_frame = frames[roles["purchases"]].copy(deep=True)
    sales_frame = frames[roles["sales"]].copy(deep=True)
    adjustment_frame = frames[roles["adjustments"]].copy(deep=True)

    product_cols = _frame_columns(product_frame, {
        "code": ("商品编码", "产品编码", "SKU"), "name": ("商品名称", "产品名称"),
        "category": ("品类", "商品类别", "产品类别"), "supplier": ("供应商",),
        "purchase_price": ("采购单价", "采购价"), "retail_price": ("零售价", "销售单价"),
        "safety": ("安全库存",), "lead_days": ("采购提前期(天)", "采购提前期", "提前期"),
        "target_days": ("目标库存天数", "目标周转天数"), "status": ("商品状态", "状态"),
    })
    opening_cols = _frame_columns(opening_frame, {
        "code": ("商品编码", "产品编码", "SKU"), "name": ("商品名称", "产品名称"),
        "warehouse": ("仓库",), "opening": ("期初库存", "期初数量"),
        "locked": ("已锁定", "锁定库存"), "bad": ("不良品", "不良库存"), "remark": ("备注",),
    })
    purchase_cols = _frame_columns(purchase_frame, {
        "date": ("入库日期", "采购日期", "日期"), "doc": ("入库单号", "采购单号"),
        "code": ("商品编码", "产品编码", "SKU"), "quantity": ("入库数量", "采购数量", "数量"),
        "price": ("采购价", "采购单价"), "supplier": ("供应商",), "status": ("状态",), "remark": ("备注",),
    })
    sales_cols = _frame_columns(sales_frame, {
        "date": ("出库日期", "销售日期", "日期"), "doc": ("出库单号", "销售单号", "订单号"),
        "code": ("商品编码", "产品编码", "SKU"), "name": ("商品名称", "产品名称"),
        "quantity": ("数量", "出库数量", "销售数量"), "sales": ("销售额", "销售金额"),
        "channel": ("渠道", "销售渠道"), "status": ("状态",),
    })
    adjust_cols = _frame_columns(adjustment_frame, {
        "date": ("日期", "调整日期"), "code": ("商品编码", "产品编码", "SKU"),
        "quantity": ("调整数量",), "type": ("类型", "调整类型"), "status": ("状态",), "remark": ("说明", "备注"),
    })

    product_rows: list[dict[str, Any]] = []
    for _, row in product_frame.iterrows():
        code = _clean_code(row.get(product_cols["code"]))
        if not code:
            continue
        product_rows.append({
            "商品编码": code,
            "商品名称": _clean_text(row.get(product_cols["name"])),
            "品类": _clean_text(row.get(product_cols["category"])),
            "供应商": _clean_text(row.get(product_cols["supplier"])),
            "采购单价": _parse_number(row.get(product_cols["purchase_price"])),
            "零售价": _parse_number(row.get(product_cols["retail_price"])),
            "安全库存": _parse_number(row.get(product_cols["safety"])),
            "采购提前期(天)": _parse_number(row.get(product_cols["lead_days"])),
            "目标库存天数": _parse_number(row.get(product_cols["target_days"])),
            "商品状态": _clean_text(row.get(product_cols["status"])),
        })
    products = pd.DataFrame(product_rows).drop_duplicates("商品编码", keep="first").reset_index(drop=True)
    if products.empty:
        raise ValueError("商品资料表没有可用商品编码")
    product_codes = set(products["商品编码"])

    opening_rows = []
    for _, row in opening_frame.iterrows():
        opening_rows.append({
            "商品编码": _clean_code(row.get(opening_cols["code"])),
            "期初库存": _parse_number(row.get(opening_cols["opening"])),
            "已锁定": _parse_number(row.get(opening_cols["locked"])),
            "不良品": _parse_number(row.get(opening_cols["bad"])),
        })
    opening = pd.DataFrame(opening_rows)
    for column in ("期初库存", "已锁定", "不良品"):
        opening[column] = opening[column].fillna(0.0)
    opening = opening.groupby("商品编码", as_index=False, observed=True)[["期初库存", "已锁定", "不良品"]].sum()

    audit_rows: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []

    purchase_rows = []
    for _, row in purchase_frame.iterrows():
        quantity = _parse_number(row.get(purchase_cols["quantity"]))
        price = _parse_number(row.get(purchase_cols["price"]))
        purchase_rows.append({
            "日期": _parse_date(row.get(purchase_cols["date"])), "单据编号": _clean_code(row.get(purchase_cols["doc"])),
            "商品编码": _clean_code(row.get(purchase_cols["code"])), "数量": quantity, "单价": price,
            "金额": quantity * price if pd.notna(quantity) and pd.notna(price) else float("nan"),
            "供应商": _clean_text(row.get(purchase_cols["supplier"])), "状态": _clean_text(row.get(purchase_cols["status"])),
            "说明": _clean_text(row.get(purchase_cols["remark"])), "源工作表": _source_label(names[roles["purchases"]]),
        })
    purchases, purchase_duplicates = _deduplicate_documents(pd.DataFrame(purchase_rows), "单据编号")
    for _, row in purchase_duplicates.iterrows():
        audit_rows.append({"业务类型": "采购", **row.to_dict(), "处理结果": "删除重复", "原因": "同一入库单号重复导出"})

    purchase_status = purchases["状态"].astype("string")
    purchase_known = purchases["商品编码"].isin(product_codes)
    purchase_valid_qty = purchases["数量"].notna() & purchases["数量"].gt(0)
    received_mask = purchase_status.str.contains("已入库", na=False) & purchase_known & purchase_valid_qty
    transit_mask = purchase_status.str.contains("待入库", na=False) & purchase_known & purchase_valid_qty
    for index, row in purchases.loc[~received_mask & ~transit_mask].iterrows():
        reasons = []
        if not purchase_known.loc[index]: reasons.append("商品编码不在商品资料")
        if not purchase_valid_qty.loc[index]: reasons.append("入库数量缺失或非正数")
        if "取消" in str(row["状态"]): reasons.append("采购已取消")
        if not reasons: reasons.append("采购状态不计入库存")
        item = {"业务类型": "采购", **row.to_dict(), "处理结果": "排除库存计算", "原因": "；".join(reasons)}
        audit_rows.append(item)
        if not purchase_valid_qty.loc[index] or not purchase_known.loc[index]: review_rows.append(item)

    sales_rows = []
    for _, row in sales_frame.iterrows():
        sales_rows.append({
            "日期": _parse_date(row.get(sales_cols["date"])), "单据编号": _clean_code(row.get(sales_cols["doc"])),
            "商品编码": _clean_code(row.get(sales_cols["code"])), "数量": _parse_number(row.get(sales_cols["quantity"])),
            "金额": _parse_number(row.get(sales_cols["sales"])), "渠道": _clean_text(row.get(sales_cols["channel"])),
            "状态": _clean_text(row.get(sales_cols["status"])), "说明": "", "源工作表": _source_label(names[roles["sales"]]),
        })
    sales, sales_duplicates = _deduplicate_documents(pd.DataFrame(sales_rows), "单据编号")
    for _, row in sales_duplicates.iterrows():
        audit_rows.append({"业务类型": "销售", **row.to_dict(), "处理结果": "删除重复", "原因": "同一出库单号重复导出"})
    sales_status = sales["状态"].astype("string")
    sales_known = sales["商品编码"].isin(product_codes)
    sales_valid_qty = sales["数量"].notna() & sales["数量"].gt(0)
    completed_mask = sales_status.str.contains("已完成", na=False) & sales_known & sales_valid_qty
    for index, row in sales.loc[~completed_mask].iterrows():
        reasons = []
        if not sales_known.loc[index]: reasons.append("商品编码不在商品资料")
        if "取消" in str(row["状态"]): reasons.append("销售已取消")
        elif "退货" in str(row["状态"]): reasons.append("退货是否重新入库需人工确认")
        elif not sales_valid_qty.loc[index]: reasons.append("出库数量缺失或非正数")
        else: reasons.append("销售状态不计入正常出库")
        item = {"业务类型": "销售", **row.to_dict(), "处理结果": "排除正常销售", "原因": "；".join(reasons)}
        audit_rows.append(item)
        if "退货" in str(row["状态"]) or not sales_known.loc[index]: review_rows.append(item)

    adjustment_rows = []
    for position, (_, row) in enumerate(adjustment_frame.iterrows(), start=1):
        adjustment_rows.append({
            "日期": _parse_date(row.get(adjust_cols["date"])), "单据编号": f"ADJ-{position:04d}",
            "商品编码": _clean_code(row.get(adjust_cols["code"])), "数量": _parse_number(row.get(adjust_cols["quantity"])),
            "金额": float("nan"), "状态": _clean_text(row.get(adjust_cols["status"])),
            "说明": _clean_text(row.get(adjust_cols["remark"])), "源工作表": _source_label(names[roles["adjustments"]]),
        })
    adjustments = pd.DataFrame(adjustment_rows)
    adjustment_known = adjustments["商品编码"].isin(product_codes)
    adjustment_valid_qty = adjustments["数量"].notna() & adjustments["数量"].ne(0)
    confirmed_mask = adjustments["状态"].astype("string").str.contains("已确认", na=False) & adjustment_known & adjustment_valid_qty
    for index, row in adjustments.loc[~confirmed_mask].iterrows():
        reasons = []
        if not adjustment_known.loc[index]: reasons.append("商品编码不在商品资料")
        if not adjustment_valid_qty.loc[index]: reasons.append("调整数量缺失或为零")
        if "待确认" in str(row["状态"]): reasons.append("库存调整尚未审批")
        if not reasons: reasons.append("调整状态不生效")
        item = {"业务类型": "库存调整", **row.to_dict(), "处理结果": "暂不计入库存", "原因": "；".join(reasons)}
        audit_rows.append(item)
        review_rows.append(item)

    received = purchases.loc[received_mask].copy()
    transit = purchases.loc[transit_mask].copy()
    completed_sales = sales.loc[completed_mask].copy()
    confirmed_adjustments = adjustments.loc[confirmed_mask].copy()
    as_of_date = _infer_as_of_date(frames, [purchases["日期"], sales["日期"], adjustments["日期"]])
    recent_start = as_of_date - pd.Timedelta(days=recent_days - 1)

    def grouped_sum(frame: pd.DataFrame, value: str, output: str) -> pd.DataFrame:
        if frame.empty:
            return pd.DataFrame({"商品编码": pd.Series(dtype="string"), output: pd.Series(dtype="float64")})
        return frame.groupby("商品编码", as_index=False, observed=True)[value].sum().rename(columns={value: output})

    analysis = products.merge(opening, on="商品编码", how="left")
    for frame, value, output in (
        (received, "数量", "已入库数量"), (completed_sales, "数量", "已出库数量"),
        (confirmed_adjustments, "数量", "已确认调整"), (transit, "数量", "在途数量"),
    ):
        analysis = analysis.merge(grouped_sum(frame, value, output), on="商品编码", how="left")
    recent_sales = completed_sales.loc[completed_sales["日期"].between(recent_start, as_of_date, inclusive="both")]
    analysis = analysis.merge(grouped_sum(recent_sales, "数量", "近30天销量"), on="商品编码", how="left")
    numeric_fill = ["期初库存", "已锁定", "不良品", "已入库数量", "已出库数量", "已确认调整", "在途数量", "近30天销量"]
    for column in numeric_fill:
        analysis[column] = analysis[column].fillna(0.0)
    analysis["当前账面库存"] = analysis["期初库存"] + analysis["已入库数量"] - analysis["已出库数量"] + analysis["已确认调整"]
    analysis["可销售库存"] = analysis["当前账面库存"] - analysis["已锁定"] - analysis["不良品"]
    analysis["近30天日均销量"] = analysis["近30天销量"] / float(recent_days)
    analysis["可售库存天数"] = analysis["可销售库存"].div(analysis["近30天日均销量"].where(analysis["近30天日均销量"].gt(0)))
    analysis["补货触发库存"] = (
        analysis["安全库存"] + analysis["近30天日均销量"] * analysis["采购提前期(天)"]
    ).map(math.ceil)
    analysis["目标库存量"] = (
        analysis["安全库存"] + analysis["近30天日均销量"] * analysis["目标库存天数"]
    ).map(math.ceil)
    normal_product = ~analysis["商品状态"].astype("string").str.contains("停售", na=False)
    shortage = normal_product & (analysis["可销售库存"] <= analysis["补货触发库存"])
    analysis["建议补货量"] = 0.0
    analysis.loc[shortage, "建议补货量"] = (
        analysis.loc[shortage, "目标库存量"] - analysis.loc[shortage, "可销售库存"] - analysis.loc[shortage, "在途数量"]
    ).clip(lower=0).map(math.ceil)
    overstock = normal_product & (
        ((analysis["近30天日均销量"] > 0) & (analysis["可售库存天数"] > analysis["目标库存天数"] * float(overstock_multiplier)))
        | ((analysis["近30天日均销量"] <= 0) & (analysis["可销售库存"] > analysis["安全库存"]))
    )
    discontinued = ~normal_product & analysis["可销售库存"].gt(0)
    analysis["积压数量"] = 0.0
    analysis.loc[overstock, "积压数量"] = (
        analysis.loc[overstock, "可销售库存"] - analysis.loc[overstock, "目标库存量"]
    ).clip(lower=0).map(math.floor)
    analysis.loc[discontinued, "积压数量"] = analysis.loc[discontinued, "可销售库存"].clip(lower=0).map(math.floor)
    analysis["可售库存金额"] = analysis["可销售库存"] * analysis["采购单价"]
    analysis["积压金额"] = analysis["积压数量"] * analysis["采购单价"]

    statuses = []
    suggestions = []
    for _, row in analysis.iterrows():
        if row["可销售库存"] < 0:
            status, suggestion = "库存异常", "核对出入库和调整单据，避免负库存"
        elif "停售" in str(row["商品状态"]):
            status = "停售积压" if row["可销售库存"] > 0 else "停售清零"
            suggestion = "停止补货并制定清仓/调拨方案" if row["可销售库存"] > 0 else "保持停采"
        elif row["建议补货量"] > 0:
            status, suggestion = "需要补货", f"建议采购 {int(row['建议补货量'])} 件，并结合在途与供应商交期下单"
        elif bool(overstock.loc[row.name]):
            status, suggestion = "库存积压", "暂停或减少采购，优先促销、调拨和消化库存"
        else:
            status, suggestion = "正常", "维持当前补货节奏并持续监控"
        statuses.append(status)
        suggestions.append(suggestion)
    analysis["库存状态"] = statuses
    analysis["管理建议"] = suggestions

    for _, row in analysis.loc[analysis["当前账面库存"].lt(0) | analysis["可销售库存"].lt(0)].iterrows():
        review_rows.append({
            "业务类型": "库存结果", "日期": as_of_date, "单据编号": "", "商品编码": row["商品编码"],
            "数量": row["可销售库存"], "金额": row["可售库存金额"], "状态": row["库存状态"],
            "说明": row["管理建议"], "源工作表": "计算结果", "处理结果": "需要人工核验", "原因": "账面或可销售库存为负数",
        })

    analysis_columns = [
        "商品编码", "商品名称", "品类", "供应商", "商品状态", "期初库存", "已锁定", "不良品",
        "已入库数量", "已出库数量", "已确认调整", "当前账面库存", "可销售库存", "在途数量",
        "安全库存", "采购提前期(天)", "目标库存天数", "近30天销量", "近30天日均销量",
        "可售库存天数", "补货触发库存", "目标库存量", "建议补货量", "积压数量", "采购单价",
        "可售库存金额", "积压金额", "库存状态", "管理建议",
    ]
    analysis = analysis[analysis_columns].sort_values(["库存状态", "可售库存金额"], ascending=[True, False], kind="stable").reset_index(drop=True)

    replenish_columns = ["商品编码", "商品名称", "品类", "供应商", "可销售库存", "在途数量", "安全库存", "补货触发库存", "目标库存量", "采购提前期(天)", "建议补货量", "采购单价", "预计采购金额", "管理建议"]
    replenishment = analysis.loc[analysis["建议补货量"].gt(0)].copy()
    replenishment["预计采购金额"] = replenishment["建议补货量"] * replenishment["采购单价"]
    replenishment = replenishment[replenish_columns].sort_values("预计采购金额", ascending=False, kind="stable").reset_index(drop=True)

    excess_columns = ["商品编码", "商品名称", "品类", "商品状态", "可销售库存", "近30天销量", "可售库存天数", "目标库存天数", "积压数量", "采购单价", "积压金额", "管理建议"]
    excess = analysis.loc[analysis["积压数量"].gt(0), excess_columns].sort_values("积压金额", ascending=False, kind="stable").reset_index(drop=True)

    purchase_summary = products[["商品编码", "商品名称", "品类", "供应商"]].copy()
    purchase_agg = received.groupby("商品编码", as_index=False, observed=True).agg(已入库数量=("数量", "sum"), 已入库金额=("金额", "sum"), 已入库单数=("单据编号", "nunique"))
    transit_agg = transit.groupby("商品编码", as_index=False, observed=True).agg(在途数量=("数量", "sum"), 在途金额=("金额", "sum"))
    purchase_summary = purchase_summary.merge(purchase_agg, on="商品编码", how="left").merge(transit_agg, on="商品编码", how="left").fillna({"已入库数量": 0, "已入库金额": 0, "已入库单数": 0, "在途数量": 0, "在途金额": 0})
    purchase_summary["平均采购单价"] = purchase_summary["已入库金额"].div(purchase_summary["已入库数量"].where(purchase_summary["已入库数量"].gt(0)))
    purchase_summary = purchase_summary.sort_values("已入库金额", ascending=False, kind="stable").reset_index(drop=True)

    sales_summary = products[["商品编码", "商品名称", "品类"]].copy()
    sales_agg = completed_sales.groupby("商品编码", as_index=False, observed=True).agg(销售数量=("数量", "sum"), 销售额=("金额", "sum"), 销售单数=("单据编号", "nunique"))
    sales_summary = sales_summary.merge(sales_agg, on="商品编码", how="left").merge(grouped_sum(recent_sales, "数量", "近30天销量"), on="商品编码", how="left").merge(products[["商品编码", "采购单价"]], on="商品编码", how="left")
    for column in ("销售数量", "销售额", "销售单数", "近30天销量"):
        sales_summary[column] = sales_summary[column].fillna(0.0)
    sales_summary["估算销售成本"] = sales_summary["销售数量"] * sales_summary["采购单价"]
    sales_summary["估算毛利"] = sales_summary["销售额"] - sales_summary["估算销售成本"]
    sales_summary["估算毛利率"] = sales_summary["估算毛利"].div(sales_summary["销售额"].where(sales_summary["销售额"].ne(0)))
    sales_summary = sales_summary.sort_values("销售额", ascending=False, kind="stable").reset_index(drop=True)

    audit_detail = pd.DataFrame(audit_rows)
    audit_columns = ["记录类型", "业务类型", "日期", "单据编号", "商品编码", "源工作表", "状态", "数量", "金额", "处理结果", "原因", "说明"]
    audit_summary = pd.DataFrame([
        {"记录类型": "汇总", "业务类型": "采购", "单据编号": "已入库纳入", "数量": len(received), "处理结果": float(received["数量"].sum()), "原因": "状态为已入库且数量、商品编码有效"},
        {"记录类型": "汇总", "业务类型": "销售", "单据编号": "已完成纳入", "数量": len(completed_sales), "处理结果": float(completed_sales["数量"].sum()), "原因": "状态为已完成且数量、商品编码有效"},
        {"记录类型": "汇总", "业务类型": "库存调整", "单据编号": "已确认纳入", "数量": len(confirmed_adjustments), "处理结果": float(confirmed_adjustments["数量"].sum()), "原因": "仅已确认库存调整生效"},
        {"记录类型": "汇总", "业务类型": "全部", "单据编号": "排除/待核验", "数量": len(audit_rows), "处理结果": len(review_rows), "原因": "重复、取消、退货、待确认、缺失数量或未知商品"},
    ])
    if audit_detail.empty:
        audit_detail = pd.DataFrame(columns=audit_columns)
    else:
        audit_detail.insert(0, "记录类型", "明细")
    for column in audit_columns:
        if column not in audit_summary: audit_summary[column] = pd.NA
        if column not in audit_detail: audit_detail[column] = pd.NA
    audit = pd.concat([audit_summary[audit_columns], audit_detail[audit_columns]], ignore_index=True)
    reviews = pd.DataFrame(review_rows)
    for column in audit_columns[1:]:
        if column not in reviews: reviews[column] = pd.NA
    reviews = reviews[audit_columns[1:]].reset_index(drop=True)

    period_purchase_qty = float(received["数量"].sum())
    period_purchase_amount = float(received["金额"].sum())
    period_sales_qty = float(completed_sales["数量"].sum())
    period_sales_amount = float(completed_sales["金额"].sum())
    completed_with_cost = completed_sales.merge(products[["商品编码", "采购单价"]], on="商品编码", how="left")
    estimated_cogs = float((completed_with_cost["数量"] * completed_with_cost["采购单价"]).sum())
    estimated_profit = period_sales_amount - estimated_cogs
    current_qty = float(analysis["当前账面库存"].sum())
    available_qty = float(analysis["可销售库存"].sum())
    available_value = float(analysis["可售库存金额"].sum())
    replenish_count = int(analysis["建议补货量"].gt(0).sum())
    excess_count = int(analysis["积压数量"].gt(0).sum())
    excess_value = float(analysis["积压金额"].sum())
    discontinued_value = float(analysis.loc[analysis["商品状态"].astype("string").str.contains("停售", na=False), "可售库存金额"].sum())
    overview = pd.DataFrame([
        {"指标": "分析截止日期", "结果": as_of_date, "单位": "", "数据口径": "优先读取仓库说明中的截止日期"},
        {"指标": "商品SKU数", "结果": len(products), "单位": "个", "数据口径": "商品资料有效商品编码去重"},
        {"指标": "当前账面库存", "结果": current_qty, "单位": "件", "数据口径": "期初+已入库-已完成出库+已确认调整"},
        {"指标": "可销售库存", "结果": available_qty, "单位": "件", "数据口径": "当前账面库存-已锁定-不良品"},
        {"指标": "可销售库存金额", "结果": available_value, "单位": "元", "数据口径": "可销售库存×商品采购单价"},
        {"指标": "采购入库金额", "结果": period_purchase_amount, "单位": "元", "数据口径": "仅统计已入库采购单"},
        {"指标": "销售出库金额", "结果": period_sales_amount, "单位": "元", "数据口径": "仅统计已完成销售出库单"},
        {"指标": "估算销售毛利", "结果": estimated_profit, "单位": "元", "数据口径": "销售额-销售数量×商品资料采购单价"},
        {"指标": "需要补货SKU数", "结果": replenish_count, "单位": "个", "数据口径": "可售库存≤安全库存+近30天日均销量×采购提前期"},
        {"指标": "建议补货总量", "结果": float(analysis["建议补货量"].sum()), "单位": "件", "数据口径": "目标库存量-可售库存-在途库存，不低于0"},
        {"指标": "积压SKU数", "结果": excess_count, "单位": "个", "数据口径": f"可售库存天数超过目标天数的{float(overstock_multiplier):g}倍，或无近期销量/停售仍有库存"},
        {"指标": "积压库存金额", "结果": excess_value, "单位": "元", "数据口径": "积压数量×采购单价"},
        {"指标": "停售库存金额", "结果": discontinued_value, "单位": "元", "数据口径": "停售商品可销售库存×采购单价"},
        {"指标": "人工核验事项", "结果": len(reviews), "单位": "项", "数据口径": "退货、数量缺失、待确认调整、未知商品或负库存"},
    ])

    purchase_monthly = received.assign(月份=received["日期"].dt.to_period("M").astype(str)).groupby("月份", as_index=False, observed=True)["数量"].sum().rename(columns={"数量": "采购入库数量"})
    sales_monthly = completed_sales.assign(月份=completed_sales["日期"].dt.to_period("M").astype(str)).groupby("月份", as_index=False, observed=True)["数量"].sum().rename(columns={"数量": "销售出库数量"})
    monthly = purchase_monthly.merge(sales_monthly, on="月份", how="outer").fillna(0).sort_values("月份", kind="stable")
    category_stock = analysis.groupby("品类", as_index=False, observed=True)["可售库存金额"].sum().sort_values("可售库存金额", ascending=False, kind="stable")
    status_counts = analysis.groupby("库存状态", as_index=False, observed=True)["商品编码"].count().rename(columns={"商品编码": "SKU数量"}).sort_values("SKU数量", ascending=False, kind="stable")
    chart_rows = max(len(monthly), len(category_stock), len(status_counts), 1)
    charts = pd.DataFrame(index=range(chart_rows))
    for column in ("月份", "采购入库数量", "销售出库数量"):
        charts[column] = monthly[column].reindex(range(chart_rows)) if column in monthly else pd.NA
    for column in ("品类", "可售库存金额"):
        charts[column] = category_stock[column].reindex(range(chart_rows)) if column in category_stock else pd.NA
    for column in ("库存状态", "SKU数量"):
        charts[column] = status_counts[column].reindex(range(chart_rows)) if column in status_counts else pd.NA

    outputs = {
        "管理层库存总览": overview,
        "商品库存分析": analysis,
        "补货建议": replenishment,
        "积压清单": excess,
        "采购分析": purchase_summary,
        "销售分析": sales_summary,
        "人工核验": reviews,
        "数据审计": audit,
        "库存图表看板": charts,
    }
    for output in outputs.values():
        output.attrs["toolbox_report_kind"] = "inventory_management_report"
    report = {
        "as_of_date": str(as_of_date.date()), "sku_count": len(products),
        "current_inventory_qty": current_qty, "available_inventory_qty": available_qty,
        "available_inventory_value": available_value, "period_purchase_qty": period_purchase_qty,
        "period_purchase_amount": period_purchase_amount, "period_sales_qty": period_sales_qty,
        "period_sales_amount": period_sales_amount, "estimated_gross_profit": estimated_profit,
        "replenishment_sku_count": replenish_count, "replenishment_qty": float(analysis["建议补货量"].sum()),
        "overstock_sku_count": excess_count, "overstock_value": excess_value,
        "manual_review_count": len(reviews), "audit_issue_count": len(audit_rows),
        "sheet_count": len(outputs), "chart_count": 3,
    }
    return InventoryReportResult(outputs=outputs, report=report)


__all__ = [
    "InventoryReportResult", "build_inventory_management_report", "can_build_inventory_report",
    "infer_inventory_table_roles", "validate_inventory_report_params",
]
