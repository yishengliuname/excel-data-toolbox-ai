"""Evidence-backed diagnosis for multi-platform e-commerce workbooks."""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Mapping, Sequence

import pandas as pd

from .semantic_model import assess_relationship, clean_key, find_column, infer_table_roles

_ROLE_SPECS: Mapping[str, Mapping[str, Sequence[str]]] = {
    "products": {
        "name_tokens": ("商品主数据", "商品资料", "SKU主数据"),
        "required": ("SKU", "标准单位成本"),
        "fields": ("SKU", "商品名称", "商品类型", "标准单位成本", "建议零售价", "安全库存"),
    },
    "orders": {
        "name_tokens": ("订单明细", "销售订单", "订单"),
        "required": ("订单号", "买家实付", "订单状态"),
        "fields": ("下单日期", "订单号", "渠道", "客户ID", "SKU", "数量", "商品原价", "优惠金额", "买家实付", "订单状态"),
    },
    "refunds": {
        "name_tokens": ("售后退款", "退款", "售后"),
        "required": ("原订单号", "退款金额", "售后状态"),
        "fields": ("申请日期", "售后单号", "原订单号", "SKU", "退款金额", "售后状态", "原因"),
    },
    "settlements": {
        "name_tokens": ("平台结算", "结算", "回款"),
        "required": ("结算月份", "渠道", "实际到账"),
        "fields": ("结算月份", "渠道", "订单结算基数", "平台佣金", "技术服务费", "物流/服务扣款", "退款冲减", "实际到账"),
    },
    "ads": {
        "name_tokens": ("广告投放", "投放", "广告"),
        "required": ("月份", "渠道", "广告花费"),
        "fields": ("月份", "渠道", "广告花费", "曝光量", "点击量", "归因成交额", "归因订单数"),
    },
    "purchases": {
        "name_tokens": ("采购入库", "采购", "入库"),
        "required": ("入库单号", "SKU", "采购单价"),
        "fields": ("入库日期", "入库单号", "SKU", "入库数量", "采购单价", "采购金额", "状态"),
    },
    "inventory": {
        "name_tokens": ("月末库存", "库存"),
        "required": ("月份", "SKU", "可售库存", "库存金额"),
        "fields": ("月份", "SKU", "账面库存", "已锁定", "不良品", "可售库存", "库存金额"),
    },
    "customers": {
        "name_tokens": ("客户会员", "客户", "会员"),
        "required": ("客户ID", "客户评分"),
        "fields": ("客户ID", "会员等级", "累计实付", "售后次数", "退款金额", "客户评分", "地区"),
    },
    "notes": {
        "name_tokens": ("经营说明", "口径说明", "项目说明"),
        "required": (),
        "fields": ("项目背景与口径说明", "老板要求", "订单口径", "退款口径", "成本口径", "广告口径"),
    },
}


def infer_ecommerce_table_roles(
    frames: Sequence[pd.DataFrame], source_names: Sequence[str] | None = None
) -> tuple[dict[str, int], list[Any]]:
    names = list(source_names or [f"表{index + 1}" for index in range(len(frames))])
    return infer_table_roles(
        frames,
        names,
        _ROLE_SPECS,
        required_roles=("products", "orders", "refunds", "settlements", "ads", "inventory"),
    )


def can_build_ecommerce_diagnosis_report(
    frames: Sequence[pd.DataFrame], source_names: Sequence[str] | None = None
) -> bool:
    try:
        roles, _ = infer_ecommerce_table_roles(frames, source_names)
    except (TypeError, ValueError):
        return False
    return len(roles) >= 6


@dataclass(frozen=True)
class ECommerceDiagnosisResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def _col(frame: pd.DataFrame, *aliases: str, required: bool = True) -> pd.Series:
    name = find_column(frame, aliases)
    if name is None:
        if required:
            raise ValueError(f"缺少字段：{'/'.join(aliases)}")
        return pd.Series(pd.NA, index=frame.index)
    return frame[name]


def _num(series: pd.Series) -> pd.Series:
    text = series.astype("string").str.replace(r"[¥￥元,，%％\s件个台套]", "", regex=True)
    return pd.to_numeric(text, errors="coerce").astype("float64")


def _date(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce")
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    serial_mask = numeric.between(20_000, 80_000, inclusive="both")
    parsed.loc[serial_mask] = pd.to_datetime(numeric.loc[serial_mask], unit="D", origin="1899-12-30")
    text = series.astype("string").str.replace("年", "-", regex=False).str.replace("月", "-", regex=False).str.replace("日", "", regex=False)
    parsed.loc[~serial_mask] = pd.to_datetime(text.loc[~serial_mask], errors="coerce", format="mixed")
    return parsed


def _month_text(series: pd.Series) -> pd.Series:
    parsed = _date(series)
    direct = series.astype("string").str.extract(r"((?:19|20)\d{2})\D*([01]?\d)", expand=True)
    result = parsed.dt.to_period("M").astype("string")
    direct_value = direct[0].fillna("") + "-" + direct[1].fillna("").str.zfill(2)
    return result.mask(parsed.isna() & direct[0].notna(), direct_value)


def _ratio(numerator: float, denominator: float) -> float:
    return float(numerator) / float(denominator) if denominator else float("nan")


def _safe_sum(series: pd.Series) -> float:
    return float(pd.to_numeric(series, errors="coerce").sum(min_count=1) or 0.0)


def _risk_level(value: float, *, high: float, medium: float, reverse: bool = False) -> str:
    if pd.isna(value):
        return "待核验"
    if reverse:
        return "P0" if value <= high else "P1" if value <= medium else "P2"
    return "P0" if value >= high else "P1" if value >= medium else "P2"


def build_ecommerce_diagnosis_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str] | None = None,
    user_request: str = "",
) -> ECommerceDiagnosisResult:
    roles, role_evidence = infer_ecommerce_table_roles(frames, source_names)
    products_raw = frames[roles["products"]].copy()
    orders_raw = frames[roles["orders"]].copy()
    refunds_raw = frames[roles["refunds"]].copy()
    settlements_raw = frames[roles["settlements"]].copy()
    ads_raw = frames[roles["ads"]].copy()
    inventory_raw = frames[roles["inventory"]].copy()
    purchases_raw = frames[roles["purchases"]].copy() if "purchases" in roles else pd.DataFrame()
    customers_raw = frames[roles["customers"]].copy() if "customers" in roles else pd.DataFrame()

    products = pd.DataFrame(
        {
            "SKU": clean_key(_col(products_raw, "SKU")),
            "商品名称": _col(products_raw, "商品名称", "产品名称").astype("string").str.strip(),
            "商品类型": _col(products_raw, "商品类型", required=False).astype("string").str.strip(),
            "标准单位成本": _num(_col(products_raw, "标准单位成本", "标准成本")),
            "建议零售价": _num(_col(products_raw, "建议零售价", required=False)),
            "安全库存": _num(_col(products_raw, "安全库存", required=False)),
        }
    ).drop_duplicates("SKU", keep="last")

    orders = pd.DataFrame(
        {
            "下单日期": _date(_col(orders_raw, "下单日期", "订单日期")),
            "订单号": clean_key(_col(orders_raw, "订单号")),
            "渠道": _col(orders_raw, "渠道", "平台").astype("string").str.strip(),
            "客户ID": clean_key(_col(orders_raw, "客户ID", "客户编号")),
            "SKU": clean_key(_col(orders_raw, "SKU")),
            "商品名称": _col(orders_raw, "商品名称", required=False).astype("string").str.strip(),
            "数量": _num(_col(orders_raw, "数量")),
            "商品原价": _num(_col(orders_raw, "商品原价", "销售原价")),
            "优惠金额": _num(_col(orders_raw, "优惠金额", required=False)).fillna(0),
            "买家实付": _num(_col(orders_raw, "买家实付", "实付金额")),
            "订单状态": _col(orders_raw, "订单状态", "状态").astype("string").str.strip(),
        }
    )
    orders["月份"] = orders["下单日期"].dt.to_period("M").astype("string")
    order_duplicate_rows = int(orders.duplicated(keep="first").sum())
    orders = orders.drop_duplicates(keep="first").copy()
    invalid_pattern = r"取消|待付款|关闭|作废|未支付"
    valid_orders = orders.loc[~orders["订单状态"].str.contains(invalid_pattern, na=False)].copy()
    invalid_orders = orders.loc[orders["订单状态"].str.contains(invalid_pattern, na=False)].copy()
    valid_orders = valid_orders.merge(products, on="SKU", how="left", suffixes=("", "_主数据"), validate="many_to_one")
    valid_orders["标准商品成本"] = valid_orders["数量"] * valid_orders["标准单位成本"]
    valid_orders["折扣率"] = valid_orders["优惠金额"] / valid_orders["商品原价"].where(valid_orders["商品原价"].ne(0))

    refunds = pd.DataFrame(
        {
            "申请日期": _date(_col(refunds_raw, "申请日期")),
            "售后单号": clean_key(_col(refunds_raw, "售后单号")),
            "原订单号": clean_key(_col(refunds_raw, "原订单号")),
            "SKU": clean_key(_col(refunds_raw, "SKU")),
            "退款类型": _col(refunds_raw, "退款类型", required=False).astype("string").str.strip(),
            "退款数量": _num(_col(refunds_raw, "退款数量", required=False)),
            "退款金额": _num(_col(refunds_raw, "退款金额")),
            "售后状态": _col(refunds_raw, "售后状态", "状态").astype("string").str.strip(),
            "原因": _col(refunds_raw, "原因", required=False).astype("string").str.strip(),
        }
    )
    order_keys = valid_orders[["订单号", "渠道", "客户ID", "月份"]].drop_duplicates("订单号")
    refunds = refunds.merge(order_keys, left_on="原订单号", right_on="订单号", how="left", validate="many_to_one")
    occurred_refunds = refunds.loc[refunds["售后状态"].str.contains("已退款|退款成功", na=False)].copy()
    pending_refunds = refunds.loc[~refunds.index.isin(occurred_refunds.index)].copy()

    settlements = pd.DataFrame(
        {
            "月份": _month_text(_col(settlements_raw, "结算月份", "月份")),
            "渠道": _col(settlements_raw, "渠道", "平台").astype("string").str.strip(),
            "订单结算基数": _num(_col(settlements_raw, "订单结算基数")),
            "平台佣金": _num(_col(settlements_raw, "平台佣金", required=False)).fillna(0),
            "技术服务费": _num(_col(settlements_raw, "技术服务费", required=False)).fillna(0),
            "物流服务扣款": _num(_col(settlements_raw, "物流/服务扣款", "物流服务扣款", required=False)).fillna(0),
            "退款冲减": _num(_col(settlements_raw, "退款冲减", required=False)).fillna(0),
            "实际到账": _num(_col(settlements_raw, "实际到账")),
        }
    )
    settlements["平台费用"] = settlements[["平台佣金", "技术服务费", "物流服务扣款"]].sum(axis=1)
    settlements["勾稽差额"] = (
        settlements["订单结算基数"] - settlements["平台费用"] - settlements["退款冲减"] - settlements["实际到账"]
    )

    ads = pd.DataFrame(
        {
            "月份": _month_text(_col(ads_raw, "月份")),
            "渠道": _col(ads_raw, "渠道", "平台").astype("string").str.strip(),
            "广告花费": _num(_col(ads_raw, "广告花费")),
            "曝光量": _num(_col(ads_raw, "曝光量", required=False)),
            "点击量": _num(_col(ads_raw, "点击量", required=False)),
            "归因成交额": _num(_col(ads_raw, "归因成交额")),
            "归因订单数": _num(_col(ads_raw, "归因订单数", required=False)),
        }
    )
    ads["ROAS"] = ads["归因成交额"] / ads["广告花费"].where(ads["广告花费"].ne(0))
    ads["点击率"] = ads["点击量"] / ads["曝光量"].where(ads["曝光量"].ne(0))

    refund_by_order = occurred_refunds.groupby("原订单号", dropna=False)["退款金额"].sum()
    valid_orders["订单已退款"] = valid_orders["订单号"].map(refund_by_order).fillna(0)
    valid_orders["订单退款分摊"] = valid_orders["订单已退款"] * valid_orders["买家实付"] / valid_orders.groupby("订单号")["买家实付"].transform("sum").where(lambda x: x.ne(0))
    valid_orders["退款后管理收入"] = valid_orders["买家实付"] - valid_orders["订单退款分摊"].fillna(0)
    valid_orders["商品毛利"] = valid_orders["退款后管理收入"] - valid_orders["标准商品成本"]

    order_total = _safe_sum(valid_orders["买家实付"])
    original_total = _safe_sum(valid_orders["商品原价"])
    discount_total = _safe_sum(valid_orders["优惠金额"])
    refund_total = _safe_sum(occurred_refunds["退款金额"])
    pending_refund_total = _safe_sum(pending_refunds["退款金额"])
    net_sales = order_total - refund_total
    standard_cost = _safe_sum(valid_orders["标准商品成本"])
    product_gross_profit = net_sales - standard_cost
    platform_fees = _safe_sum(settlements["平台费用"])
    actual_arrival = _safe_sum(settlements["实际到账"])
    ad_spend = _safe_sum(ads["广告花费"])
    attributed_sales = _safe_sum(ads["归因成交额"])
    management_contribution = actual_arrival - standard_cost - ad_spend

    channel_sales = valid_orders.groupby("渠道", as_index=False).agg(
        成交实付=("买家实付", "sum"), 商品原价=("商品原价", "sum"), 标准商品成本=("标准商品成本", "sum"), 订单数=("订单号", "nunique")
    )
    channel_refunds = occurred_refunds.groupby("渠道", as_index=False)["退款金额"].sum()
    channel_settlement = settlements.groupby("渠道", as_index=False).agg(
        平台费用=("平台费用", "sum"), 实际到账=("实际到账", "sum"), 结算基数=("订单结算基数", "sum")
    )
    channel_ads = ads.groupby("渠道", as_index=False).agg(广告费=("广告花费", "sum"), 归因成交额=("归因成交额", "sum"))
    channel = channel_sales.merge(channel_refunds, on="渠道", how="left").merge(channel_settlement, on="渠道", how="left").merge(channel_ads, on="渠道", how="left")
    channel[["退款金额", "平台费用", "实际到账", "结算基数", "广告费", "归因成交额"]] = channel[["退款金额", "平台费用", "实际到账", "结算基数", "广告费", "归因成交额"]].fillna(0)
    channel["退款率"] = channel["退款金额"] / channel["成交实付"].where(channel["成交实付"].ne(0))
    channel["平台费率"] = channel["平台费用"] / channel["结算基数"].where(channel["结算基数"].ne(0))
    channel["ROAS"] = channel["归因成交额"] / channel["广告费"].where(channel["广告费"].ne(0))
    channel["管理贡献"] = channel["实际到账"] - channel["标准商品成本"] - channel["广告费"]
    channel["贡献率"] = channel["管理贡献"] / channel["成交实付"].where(channel["成交实付"].ne(0))
    channel["风险优先级"] = channel["管理贡献"].map(lambda value: "P0" if value < 0 else "P1" if value / max(order_total, 1) < 0.03 else "P2")
    channel = channel.sort_values(["风险优先级", "管理贡献"], ascending=[True, True]).reset_index(drop=True)

    month_sales = valid_orders.groupby("月份", as_index=False).agg(成交实付=("买家实付", "sum"), 标准商品成本=("标准商品成本", "sum"), 订单数=("订单号", "nunique"))
    month_refund = occurred_refunds.groupby("月份", as_index=False)["退款金额"].sum()
    month_settle = settlements.groupby("月份", as_index=False).agg(实际到账=("实际到账", "sum"), 平台费用=("平台费用", "sum"))
    month_ads = ads.groupby("月份", as_index=False).agg(广告费=("广告花费", "sum"), 归因成交额=("归因成交额", "sum"))
    month = month_sales.merge(month_refund, on="月份", how="left").merge(month_settle, on="月份", how="left").merge(month_ads, on="月份", how="left").fillna(0)
    month["退款后管理收入"] = month["成交实付"] - month["退款金额"]
    month["商品毛利"] = month["退款后管理收入"] - month["标准商品成本"]
    month["ROAS"] = month["归因成交额"] / month["广告费"].where(month["广告费"].ne(0))
    month["趋势经营贡献"] = month["实际到账"] - month["标准商品成本"] - month["广告费"]
    month = month.sort_values("月份").reset_index(drop=True)

    product = valid_orders.groupby(["SKU", "商品名称"], as_index=False).agg(
        成交实付=("买家实付", "sum"), 商品原价=("商品原价", "sum"), 优惠金额=("优惠金额", "sum"), 标准商品成本=("标准商品成本", "sum"), 销量=("数量", "sum"), 订单数=("订单号", "nunique")
    )
    product_refunds = occurred_refunds.groupby("SKU", as_index=False)["退款金额"].sum()
    product = product.merge(product_refunds, on="SKU", how="left").fillna({"退款金额": 0})
    product["退款后管理收入"] = product["成交实付"] - product["退款金额"]
    product["退款后商品毛利"] = product["退款后管理收入"] - product["标准商品成本"]
    product["商品毛利率"] = product["退款后商品毛利"] / product["退款后管理收入"].where(product["退款后管理收入"].ne(0))
    product["优惠率"] = product["优惠金额"] / product["商品原价"].where(product["商品原价"].ne(0))
    product["退款率"] = product["退款金额"] / product["成交实付"].where(product["成交实付"].ne(0))
    product["风险优先级"] = product.apply(lambda row: "P0" if row["退款后商品毛利"] < 0 else "P1" if row["商品毛利率"] < 0.2 or row["退款率"] >= 0.3 else "P2", axis=1)
    product = product.sort_values(["风险优先级", "退款后商品毛利"], ascending=[True, True]).reset_index(drop=True)

    inventory = pd.DataFrame(
        {
            "月份": _month_text(_col(inventory_raw, "月份")),
            "SKU": clean_key(_col(inventory_raw, "SKU")),
            "商品名称": _col(inventory_raw, "商品名称", required=False).astype("string").str.strip(),
            "账面库存": _num(_col(inventory_raw, "账面库存", required=False)),
            "已锁定": _num(_col(inventory_raw, "已锁定", required=False)).fillna(0),
            "不良品": _num(_col(inventory_raw, "不良品", required=False)).fillna(0),
            "可售库存": _num(_col(inventory_raw, "可售库存")),
            "库存金额": _num(_col(inventory_raw, "库存金额")),
        }
    ).merge(products[["SKU", "商品类型", "安全库存", "标准单位成本"]], on="SKU", how="left", validate="many_to_one")
    inventory_trend = inventory.groupby("月份", as_index=False).agg(库存金额=("库存金额", "sum"), 可售库存=("可售库存", "sum")).sort_values("月份")
    latest_month = inventory["月份"].dropna().max()
    latest_inventory = inventory.loc[inventory["月份"].eq(latest_month) & ~inventory["商品类型"].eq("服务")].copy()
    latest_sales_month = valid_orders["月份"].dropna().max()
    latest_units = valid_orders.loc[valid_orders["月份"].eq(latest_sales_month)].groupby("SKU")["数量"].sum()
    latest_inventory["最近月销量"] = latest_inventory["SKU"].map(latest_units).fillna(0)
    latest_inventory["库存覆盖月数"] = latest_inventory["可售库存"] / latest_inventory["最近月销量"].where(latest_inventory["最近月销量"].ne(0))
    latest_inventory["安全库存差额"] = latest_inventory["可售库存"] - latest_inventory["安全库存"]
    latest_inventory["风险优先级"] = latest_inventory.apply(lambda row: "P0" if row["库存覆盖月数"] >= 18 else "P1" if row["库存覆盖月数"] >= 9 else "P2", axis=1)
    latest_inventory["管理诊断"] = latest_inventory.apply(lambda row: f"按最近单月销量约{row['库存覆盖月数']:.1f}个月；不含季节性、促销与预测", axis=1)
    latest_inventory = latest_inventory.sort_values("库存覆盖月数", ascending=False).reset_index(drop=True)
    latest_inventory_value = _safe_sum(latest_inventory["库存金额"])
    inventory_growth = _ratio(inventory_trend.iloc[-1]["库存金额"] - inventory_trend.iloc[0]["库存金额"], inventory_trend.iloc[0]["库存金额"]) if len(inventory_trend) > 1 else float("nan")

    purchase_output = pd.DataFrame()
    purchase_duplicate_rows = 0
    pending_purchase_amount = 0.0
    if not purchases_raw.empty:
        purchases = pd.DataFrame(
            {
                "入库日期": _date(_col(purchases_raw, "入库日期")),
                "入库单号": clean_key(_col(purchases_raw, "入库单号")),
                "SKU": clean_key(_col(purchases_raw, "SKU")),
                "入库数量": _num(_col(purchases_raw, "入库数量")),
                "采购单价": _num(_col(purchases_raw, "采购单价")),
                "采购金额": _num(_col(purchases_raw, "采购金额")),
                "状态": _col(purchases_raw, "状态").astype("string").str.strip(),
            }
        )
        dedupe_key = ["入库单号", "SKU", "入库数量", "采购单价", "采购金额"]
        purchase_duplicate_rows = int(purchases.duplicated(dedupe_key, keep="first").sum())
        purchases = purchases.drop_duplicates(dedupe_key, keep="first")
        pending_purchase_amount = _safe_sum(purchases.loc[~purchases["状态"].str.contains("已入库", na=False), "采购金额"])
        received = purchases.loc[purchases["状态"].str.contains("已入库", na=False)].copy()
        received = received.merge(products[["SKU", "商品名称", "标准单位成本"]], on="SKU", how="left", validate="many_to_one")
        received = received.sort_values("入库日期")
        purchase_output = received.groupby(["SKU", "商品名称", "标准单位成本"], as_index=False).agg(
            入库数量=("入库数量", "sum"), 实际采购金额=("采购金额", "sum"), 最近采购价=("采购单价", "last"), 首次采购价=("采购单价", "first")
        )
        purchase_output["较标准成本偏差"] = purchase_output["最近采购价"] / purchase_output["标准单位成本"].where(purchase_output["标准单位成本"].ne(0)) - 1
        purchase_output["较首次采购价变化"] = purchase_output["最近采购价"] / purchase_output["首次采购价"].where(purchase_output["首次采购价"].ne(0)) - 1
        purchase_output["管理诊断"] = purchase_output["较标准成本偏差"].map(lambda value: "采购价高于标准成本，需复核成本标准" if value > 0.02 else "未触发显著上行线索")
        purchase_output = purchase_output.sort_values("较标准成本偏差", ascending=False)

    customer = valid_orders.groupby("客户ID", as_index=False).agg(成交实付=("买家实付", "sum"), 订单数=("订单号", "nunique"), 主要渠道=("渠道", lambda s: s.mode().iloc[0] if not s.mode().empty else ""))
    customer_refund = occurred_refunds.groupby("客户ID", as_index=False)["退款金额"].sum()
    customer = customer.merge(customer_refund, on="客户ID", how="left").fillna({"退款金额": 0})
    if not customers_raw.empty:
        customer_master = pd.DataFrame(
            {
                "客户ID": clean_key(_col(customers_raw, "客户ID")),
                "会员等级": _col(customers_raw, "会员等级", required=False).astype("string").str.strip(),
                "客户评分": _num(_col(customers_raw, "客户评分")),
                "地区": _col(customers_raw, "地区", required=False).astype("string").str.strip(),
            }
        ).drop_duplicates("客户ID", keep="last")
        customer = customer.merge(customer_master, on="客户ID", how="left", validate="one_to_one")
    else:
        customer["会员等级"] = pd.NA
        customer["客户评分"] = math.nan
        customer["地区"] = pd.NA
    customer["退款率"] = customer["退款金额"] / customer["成交实付"].where(customer["成交实付"].ne(0))
    customer["风险优先级"] = customer.apply(lambda row: "P0" if row["退款率"] >= 0.4 or (pd.notna(row["客户评分"]) and row["客户评分"] < 3) else "P1" if row["退款率"] >= 0.2 or (pd.notna(row["客户评分"]) and row["客户评分"] < 4) else "P2", axis=1)
    customer["风险证据"] = customer.apply(lambda row: f"退款率{row['退款率']:.1%}；客户评分{row['客户评分'] if pd.notna(row['客户评分']) else '待核验'}", axis=1)
    customer = customer.sort_values(["风险优先级", "退款率", "成交实付"], ascending=[True, False, False]).reset_index(drop=True)

    problems: list[dict[str, Any]] = []
    worst_channel = channel.sort_values("管理贡献").iloc[0]
    if worst_channel["管理贡献"] < 0:
        problems.append({"优先级": "P0", "风险事项": f"{worst_channel['渠道']}渠道出现负管理贡献", "数据证据": f"成交{worst_channel['成交实付']:,.0f}元、到账{worst_channel['实际到账']:,.0f}元、标准成本{worst_channel['标准商品成本']:,.0f}元、广告{worst_channel['广告费']:,.0f}元，管理贡献{worst_channel['管理贡献']:,.0f}元", "风险影响": "规模增长可能由高投放买来，进一步放量会扩大现金消耗", "建议行动": "暂停低效计划扩量，按渠道×计划复盘毛利后ROAS并设置止损线", "责任角色": "电商负责人/投放负责人", "完成期限": "7天", "验收指标": "负贡献计划全部停投或完成纠偏", "人工审批点": "预算调整与停投需负责人审批"})
    if len(month) >= 2 and month.iloc[-1]["ROAS"] < month.iloc[0]["ROAS"]:
        problems.append({"优先级": "P0", "风险事项": "广告投入增长但效率持续下降", "数据证据": f"广告费从{month.iloc[0]['广告费']:,.0f}元增至{month.iloc[-1]['广告费']:,.0f}元，ROAS从{month.iloc[0]['ROAS']:.2f}降至{month.iloc[-1]['ROAS']:.2f}", "风险影响": "新增销售未形成可持续利润和现金", "建议行动": "按渠道、计划、素材建立边际ROAS监控，低于盈亏线自动预警", "责任角色": "投放负责人/财务BP", "完成期限": "14天", "验收指标": "ROAS止跌且预算向高贡献渠道迁移", "人工审批点": "盈亏线需财务确认成本口径"})
    if inventory_growth > 0.3:
        problems.append({"优先级": "P0", "风险事项": "库存资金占用快速上升", "数据证据": f"库存金额从{inventory_trend.iloc[0]['库存金额']:,.0f}元升至{inventory_trend.iloc[-1]['库存金额']:,.0f}元，增长{inventory_growth:.1%}", "风险影响": "现金被存货占用，滞销与跌价风险上升", "建议行动": "冻结超覆盖SKU补货，制定清库存与采购降速清单", "责任角色": "供应链负责人/财务负责人", "完成期限": "14天", "验收指标": "P0库存SKU补货归零并建立周转目标", "人工审批点": "促销折价和采购取消需审批"})
    negative_products = product.loc[product["退款后商品毛利"].lt(0)]
    if not negative_products.empty:
        item = negative_products.iloc[0]
        problems.append({"优先级": "P1", "风险事项": f"{item['商品名称']}退款后商品毛利为负", "数据证据": f"退款后收入{item['退款后管理收入']:,.0f}元、标准成本{item['标准商品成本']:,.0f}元、商品毛利{item['退款后商品毛利']:,.0f}元", "风险影响": "继续促销可能形成越卖越亏", "建议行动": "复核折扣、套装成本和退款原因，未完成纠偏前限制促销", "责任角色": "商品负责人/财务BP", "完成期限": "14天", "验收指标": "退款后商品毛利率转正", "人工审批点": "标准成本与可回收入库价值需人工确认"})
    if (customer["风险优先级"] == "P0").any():
        item = customer.loc[customer["风险优先级"].eq("P0")].iloc[0]
        problems.append({"优先级": "P1", "风险事项": "高退款低评分客户需要重点复核", "数据证据": f"客户{item['客户ID']}：{item['风险证据']}", "风险影响": "可能反映商品、直播承诺或履约质量问题", "建议行动": "逐单复盘售后原因并回溯渠道、主播与SKU", "责任角色": "客服负责人/渠道负责人", "完成期限": "7天", "验收指标": "高风险客户逐单形成原因与整改记录", "人工审批点": "不得因风险标签自动限制客户权益"})
    if pending_refund_total:
        problems.append({"优先级": "P1", "风险事项": "处理中退款形成潜在现金风险", "数据证据": f"处理中退款{pending_refund_total:,.0f}元，未冲减已实现收入", "风险影响": "未来到账和利润可能继续下降", "建议行动": "跟踪判责、预计退款时间及退货可回收入库状态", "责任角色": "客服负责人/财务", "完成期限": "3天", "验收指标": "所有处理中退款状态闭环", "人工审批点": "未完成退款不得直接冲减已实现收入"})
    actions = pd.DataFrame(problems).sort_values("优先级").reset_index(drop=True)

    summary = pd.DataFrame(
        [
            ("有效成交订单数", valid_orders["订单号"].nunique(), "单", "已完成等有效状态；同一订单多SKU按一个订单计"),
            ("有效成交明细数", len(valid_orders), "行", "删除1条完整重复导出后保留多SKU明细"),
            ("买家实付", order_total, "元", "有效订单商品明细买家实付求和"),
            ("整体优惠率", _ratio(discount_total, original_total), "%", "优惠金额÷商品原价"),
            ("已发生退款", refund_total, "元", "仅售后状态为已退款/退款成功"),
            ("退款后管理收入", net_sales, "元", "买家实付-已发生退款；非财务确认收入"),
            ("标准商品成本", standard_cost, "元", "有效订单数量×商品主数据标准单位成本"),
            ("管理口径商品毛利", product_gross_profit, "元", "退款后管理收入-标准商品成本"),
            ("管理口径商品毛利率", _ratio(product_gross_profit, net_sales), "%", "管理口径商品毛利÷退款后管理收入"),
            ("平台费用", platform_fees, "元", "佣金+技术服务费+物流/服务扣款"),
            ("实际到账", actual_arrival, "元", "平台结算表实际到账求和；存在周期差"),
            ("广告花费", ad_spend, "元", "广告投放表求和"),
            ("整体ROAS", _ratio(attributed_sales, ad_spend), "倍", "平台7日归因成交额÷广告花费；不与订单收入相加"),
            ("趋势性管理贡献", management_contribution, "元", "实际到账-标准商品成本-广告费；不等同财务净利润"),
            ("期末库存金额", latest_inventory_value, "元", f"{latest_month}月末实物库存金额"),
            ("库存金额增幅", inventory_growth, "%", "首月到末月库存金额变化"),
            ("待处理退款", pending_refund_total, "元", "风险披露，不冲减已实现收入"),
        ],
        columns=["指标", "结果", "单位", "数据口径"],
    )

    profit_bridge = month[["月份", "订单数", "成交实付", "退款金额", "退款后管理收入", "实际到账", "标准商品成本", "平台费用", "广告费", "ROAS", "商品毛利", "趋势经营贡献"]].copy()
    settlement_output = settlements[["月份", "渠道", "订单结算基数", "平台费用", "退款冲减", "实际到账", "勾稽差额"]].copy()
    refund_output = refunds[["申请日期", "售后单号", "原订单号", "渠道", "客户ID", "SKU", "退款类型", "退款金额", "售后状态", "原因"]].copy()
    refund_output["是否冲减管理收入"] = refund_output["售后状态"].str.contains("已退款|退款成功", na=False).map({True: "是", False: "否，风险披露"})

    relation_rows: list[dict[str, Any]] = []
    for relation in (
        assess_relationship(valid_orders, "SKU", products, "SKU", left_name="订单明细", right_name="商品主数据"),
        assess_relationship(refunds, "原订单号", valid_orders, "订单号", left_name="售后退款", right_name="订单明细", require_right_unique=False),
    ):
        relation_rows.append({"审计类型": "关系证据", "审计项": f"{relation.left_table}.{relation.left_key}→{relation.right_table}.{relation.right_key}", "状态": "通过" if relation.accepted else "待核验", "数据证据/口径": relation.reason, "处理边界": "仅使用业务键；禁止金额字段同名自动关联"})
    if not customers_raw.empty:
        relation = assess_relationship(valid_orders, "客户ID", customer_master, "客户ID", left_name="订单明细", right_name="客户会员")
        relation_rows.append({"审计类型": "关系证据", "审计项": "订单明细.客户ID→客户会员.客户ID", "状态": "通过" if relation.accepted else "待核验", "数据证据/口径": relation.reason, "处理边界": "客户画像只通过客户ID关联"})
    audit_rows = [
        {"审计类型": "表角色", "审计项": item.role, "状态": "通过", "数据证据/口径": f"{item.table_name}；得分{item.score}；字段：{','.join(item.matched_fields)}", "处理边界": "每个事实域独立聚合，不选择单一主表"}
        for item in role_evidence
    ] + relation_rows + [
        {"审计类型": "清洗", "审计项": "订单重复导出", "状态": "通过", "数据证据/口径": f"删除{order_duplicate_rows}条完全重复明细；保留同一订单多个SKU", "处理边界": "不按订单号单字段去重"},
        {"审计类型": "清洗", "审计项": "无效订单", "状态": "通过", "数据证据/口径": f"排除{len(invalid_orders)}条取消/待付款/关闭明细", "处理边界": "状态规则来自经营说明"},
        {"审计类型": "清洗", "审计项": "采购重复与待质检", "状态": "通过", "数据证据/口径": f"采购重复{purchase_duplicate_rows}条；待质检{pending_purchase_amount:,.0f}元未作为已入库", "处理边界": "待质检不能直接计可售库存"},
        {"审计类型": "勾稽", "审计项": "平台结算", "状态": "通过" if settlements["勾稽差额"].abs().max() < 0.01 else "待核验", "数据证据/口径": f"最大勾稽差额{settlements['勾稽差额'].abs().max():,.2f}元", "处理边界": "结算基数-平台费用-退款冲减=实际到账"},
        {"审计类型": "人工边界", "审计项": "利润名称", "状态": "待财务确认", "数据证据/口径": "管理贡献=实际到账-标准商品成本-广告费", "处理边界": "未包含工资、仓储、税费、总部费用，不得称净利润"},
        {"审计类型": "人工边界", "审计项": "退款成本冲回", "状态": "待业务确认", "数据证据/口径": f"已退款{refund_total:,.0f}元", "处理边界": "默认不假设退货可重新入库并冲回成本"},
        {"审计类型": "人工边界", "审计项": "广告归因", "状态": "待业务确认", "数据证据/口径": "平台7日归因", "处理边界": "归因成交额不与订单收入相加，也不强行分摊到SKU"},
        {"审计类型": "人工边界", "审计项": "库存覆盖", "状态": "待业务确认", "数据证据/口径": "期末可售库存÷最近单月销量", "处理边界": "不含季节性、促销、在途采购及预测"},
    ]
    audit = pd.DataFrame(audit_rows)

    diagnosis = (
        f"增长是真的，但增长质量在恶化：成交实付{order_total:,.0f}元、实际到账{actual_arrival:,.0f}元，"
        f"扣标准商品成本和广告后趋势性管理贡献{management_contribution:,.0f}元。"
        f"{worst_channel['渠道']}渠道贡献最低（{worst_channel['管理贡献']:,.0f}元），"
        f"期末库存{latest_inventory_value:,.0f}元，较首月上升{inventory_growth:.1%}。"
    )
    top_actions = actions.head(3).to_dict("records")
    dashboard = pd.DataFrame(
        {
            "KPI_销售规模": [order_total],
            "KPI_回款率": [_ratio(actual_arrival, order_total)],
            "KPI_毛利标题": ["退款后商品毛利率"],
            "KPI_毛利率": [_ratio(product_gross_profit, net_sales)],
            "KPI_估算经营结果": [management_contribution],
            "KPI_风险订单": [refund_total],
            "KPI_库存金额": [latest_inventory_value],
            "核心诊断": [diagnosis],
            **{f"风险卡{i + 1}_标题": [f"{item['优先级']}｜{item['风险事项']}"] for i, item in enumerate(top_actions)},
            **{f"风险卡{i + 1}_证据": [item["数据证据"]] for i, item in enumerate(top_actions)},
            **{f"风险卡{i + 1}_行动": [item["建议行动"]] for i, item in enumerate(top_actions)},
        }
    )
    chart_rows = max(len(month), len(channel), min(8, len(product)), min(8, len(latest_inventory)))
    dashboard = dashboard.reindex(range(chart_rows)).copy()
    for column in dashboard.columns:
        if len(dashboard):
            dashboard.loc[1:, column] = pd.NA
    chart_payloads = {
        "月份": month["月份"], "成交实付": month["成交实付"], "实际到账": month["实际到账"], "标准成本": month["标准商品成本"], "广告费": month["广告费"], "趋势贡献": month["趋势经营贡献"],
        "渠道_渠道": channel["渠道"], "渠道_成交实付": channel["成交实付"], "渠道_实际到账": channel["实际到账"], "渠道_广告费": channel["广告费"], "渠道_管理贡献": channel["管理贡献"], "渠道_ROAS": channel["ROAS"],
        "商品_商品": product.head(8)["商品名称"], "商品_退款后商品毛利": product.head(8)["退款后商品毛利"],
        "库存_产品": latest_inventory.head(8)["商品名称"], "库存_库存月数": latest_inventory.head(8)["库存覆盖月数"],
    }
    for column, values in chart_payloads.items():
        dashboard[column] = pd.Series(values).reset_index(drop=True).reindex(range(chart_rows))

    outputs: dict[str, pd.DataFrame] = {
        "管理层诊断总览": summary,
        "利润驱动分析": profit_bridge,
        "渠道与广告诊断": channel,
        "商品利润质量": product,
        "退款售后风险": refund_output,
        "平台费用与回款": settlement_output,
        "广告效率分析": ads.sort_values(["月份", "渠道"]),
        "采购成本分析": purchase_output,
        "库存风险分析": latest_inventory[["SKU", "商品名称", "可售库存", "库存金额", "安全库存", "最近月销量", "库存覆盖月数", "安全库存差额", "风险优先级", "管理诊断"]],
        "客户与回款风险": customer,
        "风险行动计划": actions,
        "数据口径与验收": audit,
        "经营诊断看板": dashboard,
    }
    for frame in outputs.values():
        frame.attrs["toolbox_report_kind"] = "ecommerce_diagnosis_report"
    report = {
        "valid_order_count": int(valid_orders["订单号"].nunique()),
        "valid_line_count": int(len(valid_orders)),
        "buyer_paid": order_total,
        "refund_amount": refund_total,
        "net_management_sales": net_sales,
        "standard_cost": standard_cost,
        "product_gross_profit": product_gross_profit,
        "product_gross_margin": _ratio(product_gross_profit, net_sales),
        "platform_fees": platform_fees,
        "actual_arrival": actual_arrival,
        "ad_spend": ad_spend,
        "roas": _ratio(attributed_sales, ad_spend),
        "management_contribution": management_contribution,
        "latest_inventory_value": latest_inventory_value,
        "inventory_growth": inventory_growth,
        "pending_refund_amount": pending_refund_total,
        "open_definition_count": int(audit["状态"].astype(str).str.startswith("待").sum()),
        "diagnosis": diagnosis,
        "request": user_request,
    }
    return ECommerceDiagnosisResult(outputs=outputs, report=report)


__all__ = [
    "ECommerceDiagnosisResult",
    "build_ecommerce_diagnosis_report",
    "can_build_ecommerce_diagnosis_report",
    "infer_ecommerce_table_roles",
]
