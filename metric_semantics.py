"""Reusable metric and sheet semantics for adaptive business analysis.

The module intentionally separates *physical type* (a numeric column) from
*business aggregation behaviour*.  Amounts add, balances use an ending value,
scores average and ratios are recomputed from their numerator/denominator.
This prevents a common analytical failure: averaging already-calculated rates
or averaging additive expenses simply because a table contains few rows.
"""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any, Iterable, Mapping

import pandas as pd


_NORMALISE = re.compile(r"[\s_\-（）()【】\[\]：:/.]+")
_DATE = re.compile(r"日期|时间|月份|年月|期间|年度|季度|date|time|month|year|period", re.I)
_IDENTIFIER = re.compile(r"(^id$|编号|编码|单号|流水|序号|工号|账号|sku|code|no$)", re.I)
_RATIO = re.compile(r"率|比例|占比|完成度|转化|margin|rate|ratio|percent", re.I)
_SCORE = re.compile(r"评分|满意度|得分|指数|score|rating", re.I)
_BALANCE = re.compile(r"期末|余额|结余|结存|现有库存|当前库存|库存余额|balance|ending", re.I)
_ADDITIVE = re.compile(
    r"销售|营业额|收入|金额|成本|费用|利润|退款|回款|工资|薪资|租金|水电|营销|佣金|"
    r"平台费|广告|采购|损耗|实付|到账|结算基数|revenue|sales|amount|cost|expense|profit|refund|salary",
    re.I,
)
_COUNT = re.compile(r"数量|销量|订单数|人数|件数|次数|小时|天数|qty|quantity|count|hours|days", re.I)
_SUMMARY_SHEET = re.compile(
    r"汇总|总览|看板|驾驶舱|报告|分析报告|数据字典|数据质量|关系建议|风险行动|"
    r"诊断底稿|数据口径|验收|图表展示|dashboard|report|summary",
    re.I,
)
_NOTES_SHEET = re.compile(r"说明|备注|口径|readme|notes?", re.I)


def normalise(value: Any) -> str:
    return _NORMALISE.sub("", str(value or "")).casefold()


@dataclass(frozen=True)
class MetricSemantic:
    kind: str
    aggregation: str
    unit: str
    explanation: str


def classify_metric(name: Any) -> MetricSemantic:
    """Classify a field by aggregation semantics, not cardinality."""

    text = str(name or "").strip()
    if _IDENTIFIER.search(text):
        return MetricSemantic("identifier", "none", "", "标识字段，不参与数值聚合")
    if _DATE.search(text) and not (_ADDITIVE.search(text) or _COUNT.search(text) or _RATIO.search(text) or _SCORE.search(text)):
        return MetricSemantic("date", "none", "", "时间维度")
    if _RATIO.search(text):
        return MetricSemantic("ratio", "weighted_ratio", "%", "比例字段优先按分子合计÷分母合计重算")
    if _SCORE.search(text):
        return MetricSemantic("score", "mean", "分", "评分/指数采用有效值平均，必要时披露权重")
    if _BALANCE.search(text):
        return MetricSemantic("balance", "last", "", "时点余额按期末值汇总，不跨期求和")
    if _ADDITIVE.search(text):
        return MetricSemantic("additive", "sum", "元" if re.search(r"额|金额|成本|费用|利润|退款|回款|工资|薪资|租金|水电|营销|佣金|实付|到账", text) else "", "可加金额采用有效值求和")
    if _COUNT.search(text):
        return MetricSemantic("count", "sum", "", "可加数量采用有效值求和")
    return MetricSemantic("numeric", "unknown", "", "未知数值不自动聚合，必须先确认业务语义")


def classify_sheet_role(name: Any, frame: pd.DataFrame) -> str:
    """Return ``fact``/``dimension``/``summary``/``notes`` for one sheet.

    Name semantics is deliberately authoritative for obvious generated or
    management-summary tabs.  Structural hints are only used as a fallback.
    """

    label = str(name or "").rsplit("__", 1)[-1].strip()
    if _SUMMARY_SHEET.search(label):
        return "summary"
    if _NOTES_SHEET.search(label):
        return "notes"
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return "notes"
    columns = [str(column) for column in frame.columns]
    metric_count = sum(classify_metric(column).kind in {"additive", "count", "balance", "ratio", "score"} for column in columns)
    date_count = sum(classify_metric(column).kind == "date" for column in columns)
    identifier_count = sum(classify_metric(column).kind == "identifier" for column in columns)
    # A compact KPI/verification grid frequently has one or two rows and many
    # derived metrics.  Treat it as reference evidence, never as a new fact.
    if len(frame) <= 3 and metric_count >= 3 and date_count == 0:
        return "summary"
    if date_count or metric_count >= 2:
        return "fact"
    if identifier_count and metric_count <= 1:
        return "dimension"
    return "dimension" if len(frame) <= 200 else "fact"


_RATIO_RULES: tuple[tuple[re.Pattern[str], tuple[str, ...], tuple[str, ...]], ...] = (
    (re.compile(r"管理利润率|经营利润率|净利润率|利润率", re.I), ("管理利润", "经营利润", "净利润", "利润"), ("净营业收入", "营业收入", "销售额", "收入")),
    (re.compile(r"毛利率", re.I), ("毛利", "毛利润"), ("净营业收入", "营业收入", "销售额", "收入")),
    (re.compile(r"退款率", re.I), ("退款金额", "退款"), ("营业额", "销售额", "收入", "实付")),
    (re.compile(r"平台.*费率|平台费率", re.I), ("平台费", "平台费用", "佣金"), ("净营业收入", "营业额", "销售额", "收入")),
    (re.compile(r"人工.*率|人工成本率", re.I), ("人工成本", "工资成本", "人工费用"), ("净营业收入", "营业额", "销售额", "收入")),
    (re.compile(r"食材.*率|食材成本率", re.I), ("食材成本", "原料成本", "材料成本"), ("净营业收入", "营业额", "销售额", "收入")),
    (re.compile(r"成本率", re.I), ("成本",), ("净营业收入", "营业额", "销售额", "收入")),
    (re.compile(r"回款率|到账率", re.I), ("回款金额", "实际到账", "到账金额"), ("销售额", "营业额", "结算基数", "收入")),
    (re.compile(r"完成率", re.I), ("完成数量", "完成额", "实际"), ("目标数量", "目标额", "目标")),
    (re.compile(r"转化率", re.I), ("转化数", "成交数", "订单数"), ("访问数", "流量", "线索数")),
)


def _best_column(columns: Iterable[Any], candidates: Iterable[str], *, exclude: str = "") -> str | None:
    choices = [str(column) for column in columns if str(column) != exclude]
    normalised: Mapping[str, str] = {normalise(column): column for column in choices}
    for candidate in candidates:
        key = normalise(candidate)
        if key in normalised:
            return normalised[key]
    for candidate in candidates:
        key = normalise(candidate)
        matches = [column for column in choices if key and key in normalise(column)]
        if len(matches) == 1:
            return matches[0]
    return None


def ratio_components(metric_name: str, columns: Iterable[Any]) -> tuple[str, str] | None:
    """Resolve the numerator and denominator for a recognised ratio field."""

    for pattern, numerator_candidates, denominator_candidates in _RATIO_RULES:
        if not pattern.search(metric_name):
            continue
        numerator = _best_column(columns, numerator_candidates, exclude=metric_name)
        denominator = _best_column(columns, denominator_candidates, exclude=metric_name)
        if numerator and denominator and numerator != denominator:
            return numerator, denominator
    return None


def _single_time_column(frame: pd.DataFrame, *, exclude: Iterable[str] = ()) -> str | None:
    excluded = {str(column) for column in exclude}
    candidates = []
    for column in frame.columns:
        name = str(column)
        if name in excluded or classify_metric(name).kind != "date":
            continue
        if pd.to_datetime(frame[column], errors="coerce").notna().any():
            candidates.append(name)
    return candidates[0] if len(candidates) == 1 else None


def aggregate_metric(frame: pd.DataFrame, column: str) -> tuple[float, str, MetricSemantic]:
    """Aggregate one metric and return value, disclosed method and semantics."""

    semantic = classify_metric(column)
    values = pd.to_numeric(frame[column], errors="coerce")
    if semantic.aggregation == "weighted_ratio":
        components = ratio_components(column, frame.columns)
        if components:
            numerator, denominator = components
            num = pd.to_numeric(frame[numerator], errors="coerce").sum(min_count=1)
            den = pd.to_numeric(frame[denominator], errors="coerce").sum(min_count=1)
            value = float(num / den) if pd.notna(num) and pd.notna(den) and den != 0 else float("nan")
            return value, f"{numerator}合计÷{denominator}合计（加权口径）", semantic
        return float("nan"), "未识别比例分子/分母，禁止自动平均；需要人工确认", semantic
    if semantic.aggregation == "sum":
        return float(values.sum(min_count=1)), "有效值求和", semantic
    if semantic.aggregation == "last":
        date_column = _single_time_column(frame)
        if date_column is None:
            return float("nan"), "缺少唯一可用时间字段，禁止按当前行顺序认定期末值", semantic
        ordered = pd.DataFrame({"_date": pd.to_datetime(frame[date_column], errors="coerce"), "_value": values}).dropna()
        ordered = ordered.sort_values("_date", kind="stable")
        return (float(ordered.iloc[-1]["_value"]) if not ordered.empty else float("nan")), f"按{date_column}排序后取期末有效值", semantic
    if semantic.aggregation == "unknown":
        return float("nan"), "未知数值语义，禁止自动求和或平均；需要人工确认", semantic
    return float(values.mean()), "有效值算术平均", semantic


def grouped_metric(frame: pd.DataFrame, group_column: str, metric_column: str) -> pd.DataFrame:
    """Group a metric with its business-safe aggregation semantics."""

    semantic = classify_metric(metric_column)
    if semantic.aggregation == "weighted_ratio":
        components = ratio_components(metric_column, frame.columns)
        if components:
            numerator, denominator = components
            grouped = frame.groupby(group_column, dropna=False, observed=True)[[numerator, denominator]].sum(min_count=1)
            grouped[metric_column] = grouped[numerator].div(grouped[denominator].where(grouped[denominator].ne(0)))
            return grouped[[metric_column]].reset_index()
        result = frame[[group_column]].drop_duplicates().reset_index(drop=True)
        result[metric_column] = float("nan")
        return result
    if semantic.aggregation == "unknown":
        result = frame[[group_column]].drop_duplicates().reset_index(drop=True)
        result[metric_column] = float("nan")
        return result
    if semantic.aggregation == "last":
        date_column = _single_time_column(frame, exclude=(group_column,))
        if date_column is None:
            counts = pd.to_numeric(frame[metric_column], errors="coerce").groupby(frame[group_column], dropna=False).count()
            if bool(counts.le(1).all()):
                return frame.groupby(group_column, dropna=False, observed=True)[metric_column].last().reset_index()
            result = frame[[group_column]].drop_duplicates().reset_index(drop=True)
            result[metric_column] = float("nan")
            return result
        work = frame[[group_column, date_column, metric_column]].copy()
        work[date_column] = pd.to_datetime(work[date_column], errors="coerce")
        work[metric_column] = pd.to_numeric(work[metric_column], errors="coerce")
        work = work.dropna(subset=[date_column, metric_column]).sort_values(date_column, kind="stable")
        return work.groupby(group_column, dropna=False, observed=True)[metric_column].last().reset_index()
    method = "sum" if semantic.aggregation == "sum" else "mean"
    return frame.groupby(group_column, dropna=False, observed=True)[metric_column].agg(method).reset_index()


__all__ = [
    "MetricSemantic",
    "aggregate_metric",
    "classify_metric",
    "classify_sheet_role",
    "grouped_metric",
    "normalise",
    "ratio_components",
]
