"""Deterministic sales-management analysis used by the AI command workflow.

The language model may recognise the user's intent, but calculations and the
Excel-ready output tables are produced locally from an allow-listed set of
column mappings.  No model-generated formula or code is executed.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from datetime import timedelta
import re
from typing import Any

import pandas as pd


_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:]+")


def _normalise_name(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


_ALIASES: Mapping[str, tuple[str, ...]] = {
    "date_column": ("日期", "订单日期", "销售日期", "交易日期", "下单日期", "业务日期", "时间"),
    "product_column": ("产品类别", "产品", "商品类别", "商品", "品类", "产品名称"),
    "region_column": ("地区", "区域", "大区", "销售区域", "省份", "城市"),
    "salesperson_column": ("销售人员", "销售员", "业务员", "负责人", "员工", "姓名"),
    "sales_column": ("销售金额", "销售额", "订单金额", "成交金额", "成交额", "含税销售额", "营业收入", "收入", "金额"),
    "cost_column": ("成本", "销售成本", "订单成本", "总成本", "商品成本", "采购成本", "采购/服务成本"),
    "satisfaction_column": ("客户满意度", "满意度", "客户评分", "满意评分", "满意度评分", "评分"),
    "quantity_column": ("订单数量", "销售数量", "成交数量", "数量", "件数"),
}

_OPTIONAL_ALIASES: Mapping[str, tuple[str, ...]] = {
    "order_column": ("订单编号", "订单号", "流水号", "单号"),
    "status_column": ("订单状态", "状态", "是否有效"),
    "remark_column": ("备注", "临时列"),
    "source_column": ("数据来源", "来源"),
}

_REQUIRED_KEYS = (
    "date_column",
    "product_column",
    "region_column",
    "salesperson_column",
    "sales_column",
    "cost_column",
    "satisfaction_column",
)


@dataclass(frozen=True)
class SalesReportResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def infer_sales_report_columns(frame: pd.DataFrame) -> dict[str, str]:
    """Resolve common Chinese sales columns without guessing numeric fields."""

    if not isinstance(frame, pd.DataFrame):
        raise TypeError("销售经营分析输入必须是 DataFrame")
    normalised = {_normalise_name(column): str(column) for column in frame.columns}
    result: dict[str, str] = {}
    for key, aliases in _ALIASES.items():
        chosen = next(
            (normalised[_normalise_name(alias)] for alias in aliases if _normalise_name(alias) in normalised),
            None,
        )
        if chosen is None:
            # A contained alias is accepted only when it resolves uniquely.
            candidates = {
                column
                for normalised_name, column in normalised.items()
                for alias in aliases
                if len(_normalise_name(alias)) >= 2 and _normalise_name(alias) in normalised_name
            }
            if len(candidates) == 1:
                chosen = next(iter(candidates))
        if chosen is not None:
            result[key] = chosen
    missing = [key for key in _REQUIRED_KEYS if key not in result]
    if missing:
        labels = {
            "date_column": "日期",
            "product_column": "产品",
            "region_column": "地区",
            "salesperson_column": "销售人员",
            "sales_column": "销售额",
            "cost_column": "成本",
            "satisfaction_column": "客户满意度",
        }
        raise ValueError("缺少销售经营分析必要字段：" + "、".join(labels[key] for key in missing))
    return result


def validate_sales_report_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("销售经营分析参数必须是对象")
    missing = [key for key in _REQUIRED_KEYS if not isinstance(params.get(key), str) or not str(params[key]).strip()]
    if missing:
        raise ValueError("销售经营分析字段映射不完整：" + "、".join(missing))
    threshold = params.get("satisfaction_threshold", 4)
    if isinstance(threshold, bool) or not isinstance(threshold, (int, float)):
        raise TypeError("satisfaction_threshold 必须是数字")
    if not 0 <= float(threshold) <= 100:
        raise ValueError("satisfaction_threshold 必须在 0 到 100 之间")
    quantity = params.get("quantity_column")
    if quantity is not None and (not isinstance(quantity, str) or not quantity.strip()):
        raise TypeError("quantity_column 必须是非空字符串或 null")


def sales_report_column_names(params: Mapping[str, Any]) -> list[str]:
    validate_sales_report_params(params)
    names = [str(params[key]) for key in _REQUIRED_KEYS]
    if params.get("quantity_column"):
        names.append(str(params["quantity_column"]))
    return list(dict.fromkeys(names))


def _numeric(series: pd.Series, label: str) -> pd.Series:
    result = pd.to_numeric(series, errors="coerce")
    valid_source = series.notna() & series.astype("string").str.strip().ne("")
    if valid_source.any() and result[valid_source].notna().mean() < 0.8:
        raise ValueError(f"字段“{label}”无法可靠转换为数值")
    return result.astype("float64")


def _rank_table(
    work: pd.DataFrame,
    *,
    dimension: str,
    dimension_label: str,
) -> pd.DataFrame:
    grouped = (
        work.groupby(dimension, dropna=False, observed=True)
        .agg(销售额=("__sales", "sum"), 成本=("__cost", "sum"), 利润=("__profit", "sum"), 订单数=("__sales", "size"))
        .reset_index()
        .rename(columns={dimension: dimension_label})
    )
    grouped[dimension_label] = grouped[dimension_label].fillna("缺失值").astype("string")
    grouped["利润率"] = grouped["利润"].div(grouped["销售额"].where(grouped["销售额"].ne(0)))
    total_sales = float(grouped["销售额"].sum())
    grouped["销售占比"] = grouped["销售额"].div(total_sales) if total_sales else 0.0
    grouped["销售额排名"] = grouped["销售额"].rank(method="min", ascending=False).astype("Int64")
    grouped["利润排名"] = grouped["利润"].rank(method="min", ascending=False).astype("Int64")
    return grouped.sort_values(["销售额排名", dimension_label], kind="stable").reset_index(drop=True)


def _attention_rows(work: pd.DataFrame, source: pd.DataFrame, threshold: float) -> pd.DataFrame:
    q1 = work["__sales"].quantile(0.25)
    q3 = work["__sales"].quantile(0.75)
    iqr = q3 - q1
    high_value_limit = q3 + 1.5 * iqr if pd.notna(iqr) else float("inf")
    rows: list[dict[str, Any]] = []
    for position in range(len(work)):
        reasons: list[str] = []
        level = "一般"
        row = work.iloc[position]
        if pd.isna(row["__satisfaction"]):
            reasons.append("客户满意度缺失")
            level = "高"
        elif float(row["__satisfaction"]) < threshold:
            reasons.append(f"客户满意度低于 {threshold:g} 分")
            level = "高"
        if pd.isna(row["__sales"]) or pd.isna(row["__cost"]):
            reasons.append("销售额或成本缺失")
            level = "高"
        elif float(row["__sales"]) <= 0:
            reasons.append("销售额小于或等于 0")
            level = "高"
        elif float(row["__profit"]) < 0:
            reasons.append("订单利润为负")
            level = "高"
        elif pd.notna(high_value_limit) and float(row["__sales"]) > float(high_value_limit):
            reasons.append("销售额超过 IQR 异常上限")
            if level != "高":
                level = "中"
        if not reasons:
            continue
        item = {str(column): source.iloc[position][column] for column in source.columns}
        item.update(
            {
                "利润": None if pd.isna(row["__profit"]) else float(row["__profit"]),
                "利润率": None if pd.isna(row["__margin"]) else float(row["__margin"]),
                "关注级别": level,
                "关注原因": "；".join(reasons),
            }
        )
        rows.append(item)
    columns = [str(column) for column in source.columns] + ["利润", "利润率", "关注级别", "关注原因"]
    return pd.DataFrame(rows, columns=columns)


def build_sales_management_report(
    frame: pd.DataFrame,
    *,
    date_column: str,
    product_column: str,
    region_column: str,
    salesperson_column: str,
    sales_column: str,
    cost_column: str,
    satisfaction_column: str,
    quantity_column: str | None = None,
    satisfaction_threshold: float = 4,
) -> SalesReportResult:
    """Create the five requested management-report tables from one sales table."""

    params = {
        "date_column": date_column,
        "product_column": product_column,
        "region_column": region_column,
        "salesperson_column": salesperson_column,
        "sales_column": sales_column,
        "cost_column": cost_column,
        "satisfaction_column": satisfaction_column,
        "quantity_column": quantity_column,
        "satisfaction_threshold": satisfaction_threshold,
    }
    validate_sales_report_params(params)
    missing_columns = [column for column in sales_report_column_names(params) if column not in frame.columns]
    if missing_columns:
        raise KeyError("销售经营分析缺少字段：" + "、".join(missing_columns))
    if frame.empty:
        raise ValueError("销售经营分析输入表不能为空")

    source = frame.copy(deep=True).reset_index(drop=True)
    work = source.copy(deep=True)
    work["__date"] = pd.to_datetime(work[date_column], errors="coerce")
    work["__sales"] = _numeric(work[sales_column], sales_column)
    work["__cost"] = _numeric(work[cost_column], cost_column)
    work["__satisfaction"] = _numeric(work[satisfaction_column], satisfaction_column)
    work["__profit"] = work["__sales"] - work["__cost"]
    work["__margin"] = work["__profit"].div(work["__sales"].where(work["__sales"].ne(0)))

    valid_amounts = work["__sales"].notna() & work["__cost"].notna()
    total_sales = float(work.loc[valid_amounts, "__sales"].sum())
    total_cost = float(work.loc[valid_amounts, "__cost"].sum())
    total_profit = float(work.loc[valid_amounts, "__profit"].sum())
    overall_margin = total_profit / total_sales if total_sales else 0.0

    products = _rank_table(work, dimension=product_column, dimension_label="产品类别")
    salespeople = _rank_table(work, dimension=salesperson_column, dimension_label="销售人员")
    anomalies = _attention_rows(work, source, float(satisfaction_threshold))

    top_product_sales = str(products.iloc[0]["产品类别"]) if not products.empty else "—"
    top_product_profit_row = products.sort_values(["利润排名", "产品类别"], kind="stable").iloc[0] if not products.empty else None
    top_salesperson = str(salespeople.iloc[0]["销售人员"]) if not salespeople.empty else "—"
    overview = pd.DataFrame(
        [
            {"指标": "总销售额", "结果": total_sales, "单位": "元", "数据口径": f"{sales_column} 求和"},
            {"指标": "总成本", "结果": total_cost, "单位": "元", "数据口径": f"{cost_column} 求和"},
            {"指标": "总利润", "结果": total_profit, "单位": "元", "数据口径": "总销售额 - 总成本"},
            {"指标": "平均利润率", "结果": overall_margin, "单位": "", "数据口径": "总利润 ÷ 总销售额（整体加权口径）"},
            {"指标": "销售额最高产品", "结果": top_product_sales, "单位": "", "数据口径": "按产品汇总销售额后降序"},
            {"指标": "利润最高产品", "结果": "—" if top_product_profit_row is None else str(top_product_profit_row["产品类别"]), "单位": "", "数据口径": "按产品汇总利润后降序"},
            {"指标": "业绩最佳销售人员", "结果": top_salesperson, "单位": "", "数据口径": "按销售人员汇总销售额后降序"},
            {"指标": "订单记录数", "结果": int(len(source)), "单位": "条", "数据口径": "源表数据行数"},
            {"指标": "重点关注记录数", "结果": int(len(anomalies)), "单位": "条", "数据口径": f"满意度低于 {float(satisfaction_threshold):g} 分、负利润、金额缺失/异常等规则"},
        ]
    )

    monthly = (
        work.dropna(subset=["__date", "__sales"])
        .assign(月份=lambda data: data["__date"].dt.to_period("M").astype(str))
        .groupby("月份", as_index=False, observed=True)["__sales"]
        .sum()
        .rename(columns={"__sales": "月度销售额"})
        .sort_values("月份", kind="stable")
    )
    product_chart = products[["产品类别", "销售额"]].rename(columns={"销售额": "产品销售额"})
    region_chart = (
        work.groupby(region_column, dropna=False, observed=True)["__sales"]
        .sum()
        .reset_index()
        .rename(columns={region_column: "地区", "__sales": "地区销售额"})
        # Horizontal bar charts render their first category at the bottom in
        # both Excel and WPS. Ascending source order therefore places the
        # largest region at the top without reversing the value axis.
        .sort_values("地区销售额", ascending=True, kind="stable")
        .reset_index(drop=True)
    )
    region_chart["地区"] = region_chart["地区"].fillna("缺失值").astype("string")
    chart_rows = max(len(monthly), len(product_chart), len(region_chart), 1)
    chart_data = pd.DataFrame(index=range(chart_rows))
    for column in ("月份", "月度销售额"):
        chart_data[column] = monthly[column].reindex(range(chart_rows)) if column in monthly else pd.NA
    for column in ("产品类别", "产品销售额"):
        chart_data[column] = product_chart[column].reindex(range(chart_rows)) if column in product_chart else pd.NA
    for column in ("地区", "地区销售额"):
        chart_data[column] = region_chart[column].reindex(range(chart_rows)) if column in region_chart else pd.NA

    outputs = {
        "管理层数据总览": overview,
        "产品分析": products,
        "销售人员分析": salespeople,
        "异常数据提醒": anomalies,
        "图表展示": chart_data,
    }
    for output in outputs.values():
        output.attrs["toolbox_report_kind"] = "sales_management_report"
    report = {
        "source_rows": len(source),
        "total_sales": total_sales,
        "total_cost": total_cost,
        "total_profit": total_profit,
        "overall_profit_margin": overall_margin,
        "top_product_by_sales": top_product_sales,
        "top_product_by_profit": "—" if top_product_profit_row is None else str(top_product_profit_row["产品类别"]),
        "top_salesperson": top_salesperson,
        "attention_rows": len(anomalies),
        "sheet_count": 5,
        "chart_count": 3,
    }
    return SalesReportResult(outputs=outputs, report=report)


def validate_quarterly_sales_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("季度销售报告参数必须是对象")
    threshold = params.get("satisfaction_threshold", 4)
    if isinstance(threshold, bool) or not isinstance(threshold, (int, float)):
        raise TypeError("satisfaction_threshold 必须是数字")
    if not 0 <= float(threshold) <= 100:
        raise ValueError("satisfaction_threshold 必须在 0 到 100 之间")
    source_names = params.get("source_names")
    if source_names is not None and (
        not isinstance(source_names, (list, tuple))
        or not all(isinstance(item, str) and item.strip() for item in source_names)
    ):
        raise TypeError("source_names 必须是非空字符串数组")


def _find_alias_column(frame: pd.DataFrame, aliases: tuple[str, ...]) -> str | None:
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


def _compact_text(value: Any) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def _parse_business_number(value: Any) -> float:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return float("nan")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip()
    if not text:
        return float("nan")
    negative = text.startswith("(") and text.endswith(")")
    cleaned = re.sub(r"[¥￥元件分,%，,\s()]", "", text)
    try:
        number = float(cleaned)
    except ValueError:
        return float("nan")
    return -number if negative else number


def _parse_business_date(value: Any) -> pd.Timestamp | pd.NaT:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value.normalize()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if 20_000 <= number <= 80_000:
            return (pd.Timestamp("1899-12-30") + timedelta(days=number)).normalize()
    text = str(value).strip()
    if not text:
        return pd.NaT
    if re.fullmatch(r"\d{5}(?:\.0+)?", text):
        return (pd.Timestamp("1899-12-30") + timedelta(days=float(text))).normalize()
    text = text.replace("年", "-").replace("月", "-").replace("日", "").replace(".", "-").replace("/", "-")
    text = re.sub(r"-+", "-", text).strip("-")
    day_month_year = re.fullmatch(r"(\d{1,2})-(\d{1,2})-(\d{4})", text)
    if day_month_year:
        first, second, year = (int(part) for part in day_month_year.groups())
        if second > 12:
            return pd.Timestamp(year=year, month=first, day=second)
        return pd.Timestamp(year=year, month=second, day=first)
    parsed = pd.to_datetime(text, errors="coerce", yearfirst=True)
    return pd.NaT if pd.isna(parsed) else pd.Timestamp(parsed).normalize()


def _standardise_quarterly_sales_frame(frame: pd.DataFrame, source_name: str) -> pd.DataFrame | None:
    try:
        columns = infer_sales_report_columns(frame)
    except (TypeError, ValueError):
        return None
    order_column = _find_alias_column(frame, _OPTIONAL_ALIASES["order_column"])
    if order_column is None:
        return None
    status_column = _find_alias_column(frame, _OPTIONAL_ALIASES["status_column"])
    remark_column = _find_alias_column(frame, _OPTIONAL_ALIASES["remark_column"])
    source_column = _find_alias_column(frame, _OPTIONAL_ALIASES["source_column"])
    quantity_column = columns.get("quantity_column")

    records: list[dict[str, Any]] = []
    for position, (_, row) in enumerate(frame.iterrows(), start=1):
        records.append(
            {
                "订单编号": _compact_text(row.get(order_column)),
                "日期": _parse_business_date(row.get(columns["date_column"])),
                "产品类别": _compact_text(row.get(columns["product_column"])),
                "地区": _compact_text(row.get(columns["region_column"])),
                "销售人员": _compact_text(row.get(columns["salesperson_column"])),
                "数量": _parse_business_number(row.get(quantity_column)) if quantity_column else float("nan"),
                "销售额": _parse_business_number(row.get(columns["sales_column"])),
                "成本": _parse_business_number(row.get(columns["cost_column"])),
                "客户满意度": _parse_business_number(row.get(columns["satisfaction_column"])),
                "订单状态": _compact_text(row.get(status_column)) if status_column else "",
                "业务备注": str(row.get(remark_column) or "").strip() if remark_column else "",
                "数据来源": _compact_text(row.get(source_column)) if source_column else "",
                "源工作表": source_name,
                "源记录序号": position,
            }
        )
    return pd.DataFrame(records)


def _invalid_quarterly_reason(row: pd.Series) -> str:
    reasons: list[str] = []
    status = _compact_text(row.get("订单状态")).casefold()
    if "取消" in status or "作废" in status:
        reasons.append("已取消/作废")
    if "退款" in status:
        reasons.append("已退款")
    if status in {"无效", "否", "n", "no", "false"}:
        reasons.append("标记为无效")
    if not _compact_text(row.get("订单编号")):
        reasons.append("订单编号缺失")
    if pd.isna(row.get("日期")):
        reasons.append("日期无法解析")
    for label in ("产品类别", "地区", "销售人员"):
        if not _compact_text(row.get(label)):
            reasons.append(f"{label}缺失")
    sales = row.get("销售额")
    cost = row.get("成本")
    if pd.isna(sales):
        reasons.append("销售额缺失")
    elif float(sales) <= 0:
        reasons.append("销售额小于或等于0")
    if pd.isna(cost):
        reasons.append("成本缺失")
    return "；".join(dict.fromkeys(reasons))


def build_quarterly_sales_management_report(
    frames: list[pd.DataFrame] | tuple[pd.DataFrame, ...],
    *,
    source_names: list[str] | tuple[str, ...] | None = None,
    satisfaction_threshold: float = 4,
) -> SalesReportResult:
    """Clean heterogeneous monthly sheets and create an auditable quarterly report."""

    params = {
        "source_names": source_names,
        "satisfaction_threshold": satisfaction_threshold,
    }
    validate_quarterly_sales_params(params)
    if not isinstance(frames, (list, tuple)) or len(frames) < 2:
        raise ValueError("季度销售报告至少需要两张销售表")
    names = list(source_names or [f"销售表{index}" for index in range(1, len(frames) + 1)])
    if len(names) != len(frames):
        raise ValueError("source_names 数量必须与输入表数量一致")

    standardised: list[pd.DataFrame] = []
    sales_source_names: list[str] = []
    for frame, name in zip(frames, names):
        normalised = _standardise_quarterly_sales_frame(frame, name)
        if normalised is not None and not normalised.empty:
            standardised.append(normalised)
            sales_source_names.append(name)
    if len(standardised) < 2:
        raise ValueError("未能识别至少两张包含订单、日期、产品、地区、人员、销售额和成本的销售表")
    combined = pd.concat(standardised, ignore_index=True, sort=False)
    combined["__source_order"] = range(len(combined))
    combined["__order_key"] = combined["订单编号"].astype("string").str.casefold()

    duplicate_drop = pd.Series(False, index=combined.index)
    for order_key, indexes in combined.groupby("__order_key", dropna=False, sort=False).groups.items():
        positions = list(indexes)
        if not order_key or len(positions) <= 1:
            continue
        scored = []
        for index in positions:
            row = combined.loc[index]
            completeness = sum(
                not (pd.isna(row[column]) if not isinstance(row[column], str) else not row[column].strip())
                for column in ("日期", "产品类别", "地区", "销售人员", "销售额", "成本", "客户满意度")
            )
            scored.append((completeness, -int(row["__source_order"]), index))
        winner = max(scored)[2]
        for index in positions:
            if index != winner:
                duplicate_drop.loc[index] = True

    invalid_reasons = combined.apply(_invalid_quarterly_reason, axis=1)
    invalid_mask = invalid_reasons.astype("string").str.len().gt(0) & ~duplicate_drop
    valid = combined.loc[~duplicate_drop & ~invalid_mask].copy(deep=True)
    if valid.empty:
        raise ValueError("按部门口径清洗后没有可纳入统计的有效订单")

    analysis_columns = [
        "订单编号", "日期", "产品类别", "地区", "销售人员", "数量", "销售额", "成本",
        "客户满意度", "订单状态", "业务备注", "数据来源", "源工作表",
    ]
    analysis_source = valid[analysis_columns].sort_values(["日期", "订单编号"], kind="stable").reset_index(drop=True)
    base = build_sales_management_report(
        analysis_source,
        date_column="日期",
        product_column="产品类别",
        region_column="地区",
        salesperson_column="销售人员",
        sales_column="销售额",
        cost_column="成本",
        satisfaction_column="客户满意度",
        quantity_column="数量",
        satisfaction_threshold=satisfaction_threshold,
    )

    merged = analysis_source.copy(deep=True)
    merged.insert(2, "月份", merged["日期"].dt.to_period("M").astype(str))
    merged["利润"] = merged["销售额"] - merged["成本"]
    merged["利润率"] = merged["利润"].div(merged["销售额"].where(merged["销售额"].ne(0)))

    ranking_work = analysis_source.copy(deep=True)
    ranking_work["__sales"] = ranking_work["销售额"]
    ranking_work["__cost"] = ranking_work["成本"]
    ranking_work["__profit"] = ranking_work["销售额"] - ranking_work["成本"]
    regions = _rank_table(ranking_work, dimension="地区", dimension_label="地区")

    excluded_rows: list[dict[str, Any]] = []
    for index, row in combined.loc[duplicate_drop | invalid_mask].iterrows():
        item = {column: row.get(column) for column in analysis_columns}
        if duplicate_drop.loc[index]:
            item.update({"处理结果": "剔除重复", "排除原因": "同一订单编号重复导出，保留信息更完整且更早出现的一条"})
        else:
            item.update({"处理结果": "排除统计", "排除原因": invalid_reasons.loc[index]})
        excluded_rows.append(item)
    excluded = pd.DataFrame(excluded_rows, columns=analysis_columns + ["处理结果", "排除原因"])

    raw_rows = int(len(combined))
    duplicate_rows = int(duplicate_drop.sum())
    invalid_rows = int(invalid_mask.sum())
    valid_rows = int(len(analysis_source))
    attention_rows = int(len(base.outputs["异常数据提醒"]))
    report = dict(base.report)
    top_region = str(regions.iloc[0]["地区"]) if not regions.empty else "—"
    overview = pd.DataFrame(
        [
            {"指标": "季度总销售额", "结果": report["total_sales"], "单位": "元", "数据口径": "仅纳入有效、未取消、未退款且金额完整的去重订单"},
            {"指标": "季度总成本", "结果": report["total_cost"], "单位": "元", "数据口径": "有效订单成本求和"},
            {"指标": "季度总利润", "结果": report["total_profit"], "单位": "元", "数据口径": "季度总销售额 - 季度总成本"},
            {"指标": "整体利润率", "结果": report["overall_profit_margin"], "单位": "", "数据口径": "总利润 ÷ 总销售额（整体加权口径）"},
            {"指标": "有效订单记录数", "结果": valid_rows, "单位": "条", "数据口径": "清洗、去重和无效订单排除后"},
            {"指标": "原始销售记录数", "结果": raw_rows, "单位": "条", "数据口径": f"自动识别 {len(sales_source_names)} 张销售表"},
            {"指标": "重复记录剔除数", "结果": duplicate_rows, "单位": "条", "数据口径": "同一订单编号只保留一条"},
            {"指标": "无效订单排除数", "结果": invalid_rows, "单位": "条", "数据口径": "取消、退款、非正金额、金额/成本缺失或关键维度缺失"},
            {"指标": "重点关注订单数", "结果": attention_rows, "单位": "条", "数据口径": f"有效订单中满意度低于 {float(satisfaction_threshold):g} 分或评分缺失"},
            {"指标": "销售额最高产品", "结果": report["top_product_by_sales"], "单位": "", "数据口径": "按产品汇总有效订单销售额"},
            {"指标": "销售额最高地区", "结果": top_region, "单位": "", "数据口径": "按地区汇总有效订单销售额"},
            {"指标": "业绩最佳销售人员", "结果": report["top_salesperson"], "单位": "", "数据口径": "按销售人员汇总有效订单销售额"},
        ]
    )

    audit_summary = pd.DataFrame(
        [
            {"记录类型": "汇总", "订单编号": "原始记录", "源工作表": "、".join(sales_source_names), "处理结果": raw_rows, "排除原因": "三张销售表的原始数据行"},
            {"记录类型": "汇总", "订单编号": "重复剔除", "源工作表": "", "处理结果": duplicate_rows, "排除原因": "同订单编号保留信息更完整且更早出现的记录"},
            {"记录类型": "汇总", "订单编号": "无效排除", "源工作表": "", "处理结果": invalid_rows, "排除原因": "取消/退款/无效状态、非正金额、金额或成本缺失、关键维度缺失"},
            {"记录类型": "汇总", "订单编号": "有效纳入", "源工作表": "", "处理结果": valid_rows, "排除原因": "用于季度经营指标、排名和图表"},
        ]
    )
    excluded_audit = excluded.rename(columns={"订单状态": "原订单状态"})
    excluded_audit.insert(0, "记录类型", "明细")
    audit_columns = ["记录类型", "订单编号", "源工作表", "处理结果", "排除原因", "原订单状态", "销售额", "成本", "业务备注"]
    for column in audit_columns:
        if column not in audit_summary:
            audit_summary[column] = pd.NA
        if column not in excluded_audit:
            excluded_audit[column] = pd.NA
    audit = pd.concat([audit_summary[audit_columns], excluded_audit[audit_columns]], ignore_index=True)

    outputs = {
        "管理层数据总览": overview,
        "季度合并数据": merged,
        "产品分析": base.outputs["产品分析"],
        "地区分析": regions,
        "销售人员分析": base.outputs["销售人员分析"],
        "异常数据提醒": base.outputs["异常数据提醒"],
        "清洗审计": audit,
        "图表展示": base.outputs["图表展示"],
    }
    for output in outputs.values():
        output.attrs["toolbox_report_kind"] = "quarterly_sales_management_report"
    report.update(
        {
            "raw_rows": raw_rows,
            "duplicate_rows_removed": duplicate_rows,
            "invalid_rows_removed": invalid_rows,
            "valid_rows": valid_rows,
            "attention_rows": attention_rows,
            "top_region_by_sales": top_region,
            "source_sheet_count": len(sales_source_names),
            "sheet_count": len(outputs),
        }
    )
    return SalesReportResult(outputs=outputs, report=report)


__all__ = [
    "SalesReportResult",
    "build_quarterly_sales_management_report",
    "build_sales_management_report",
    "infer_sales_report_columns",
    "sales_report_column_names",
    "validate_quarterly_sales_params",
    "validate_sales_report_params",
]
