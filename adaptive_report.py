"""General-purpose adaptive analysis for unfamiliar structured workbooks.

The engine is intentionally deterministic.  It profiles every uploaded table,
infers column roles from names/types/value distributions, selects a primary
fact-like table, combines truly isomorphic period sheets, and generates a
fully auditable management workbook.  It is the local fallback after specific
sales, inventory, HR and finance workflows have had the opportunity to match.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
import math
import re
from typing import Any

import pandas as pd

try:
    from .analysis_compiler import compile_analysis
    from .metric_semantics import aggregate_metric, classify_metric, classify_sheet_role, grouped_metric
except ImportError:  # Supports: python adaptive_report.py
    from excel_data_toolbox.analysis_compiler import compile_analysis
    from excel_data_toolbox.metric_semantics import aggregate_metric, classify_metric, classify_sheet_role, grouped_metric


_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:/.]+")
_IDENTIFIER_HINT = re.compile(r"(^id$|编号|编码|单号|流水|序号|工号|账号|sku|code|no$)", re.I)
_DATE_HINT = re.compile(r"日期|时间|月份|年月|期间|年度|季度|date|time|month|year|period", re.I)
_RATE_HINT = re.compile(r"率|比例|占比|完成度|增长|转化|满意度|评分|score|rate|ratio|percent|margin", re.I)
_AMOUNT_HINT = re.compile(r"销售|收入|金额|成本|利润|费用|工资|薪资|预算|回款|库存金额|price|amount|revenue|sales|cost|profit|salary", re.I)
_QUANTITY_HINT = re.compile(r"数量|件数|人数|次数|天数|小时|库存|销量|采购量|出库量|入库量|qty|quantity|count|hours|days", re.I)


def _normalise(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


def _text(value: Any, *, limit: int = 160) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    cleaned = re.sub(r"\s+", " ", str(value)).strip()
    return cleaned[:limit]


def _numeric(series: pd.Series) -> tuple[pd.Series, float, bool]:
    if pd.api.types.is_bool_dtype(series.dtype):
        return pd.Series(float("nan"), index=series.index), 0.0, False
    if pd.api.types.is_numeric_dtype(series.dtype):
        converted = pd.to_numeric(series, errors="coerce")
        return converted, float(converted.notna().mean()), False
    raw = series.astype("string").str.strip()
    percent = bool(raw.str.contains(r"[%％]", na=False).mean() >= 0.5)
    cleaned = raw.str.replace(r"[¥￥$,，元件个台套箱人次天小时%％\s]", "", regex=True)
    converted = pd.to_numeric(cleaned, errors="coerce")
    if percent:
        converted = converted / 100.0
    denominator = int(raw.ne("").sum())
    ratio = float(converted.notna().sum() / denominator) if denominator else 0.0
    return converted, ratio, percent


def _dates(series: pd.Series) -> tuple[pd.Series, float]:
    raw = series.astype("string").str.strip()
    parsed = pd.to_datetime(raw, errors="coerce", yearfirst=True, format="mixed")
    denominator = int(raw.ne("").sum())
    ratio = float(parsed.notna().sum() / denominator) if denominator else 0.0
    return parsed, ratio


def infer_column_roles(frame: pd.DataFrame) -> dict[str, dict[str, Any]]:
    """Infer semantic roles without requiring any industry-specific columns."""

    roles: dict[str, dict[str, Any]] = {}
    rows = max(len(frame), 1)
    for column in frame.columns:
        name = str(column)
        series = frame[column]
        non_missing = series.dropna().map(_text).ne("")
        non_missing_count = int(non_missing.sum())
        unique_count = int(series.dropna().astype("string").nunique())
        uniqueness = unique_count / max(non_missing_count, 1)
        numeric_values, numeric_ratio, percent_text = _numeric(series)
        date_candidate = (
            pd.api.types.is_datetime64_any_dtype(series.dtype)
            or bool(_DATE_HINT.search(name))
            or bool(series.astype("string").str.contains(r"\d{2,4}[-/年]\d{1,2}(?:[-/月]\d{1,2})?", na=False).mean() >= 0.5)
        )
        if date_candidate:
            date_values, date_ratio = _dates(series)
        else:
            date_values, date_ratio = pd.Series(pd.NaT, index=series.index), 0.0
        semantic = classify_metric(name)
        if semantic.kind == "identifier":
            role = "标识符"
        elif semantic.kind == "date" and date_ratio >= 0.5:
            role = "日期"
        elif date_ratio >= 0.9 and numeric_ratio < 0.9:
            role = "日期"
        elif semantic.kind in {"ratio", "score"} and numeric_ratio >= 0.6:
            role = "比例/评分"
        elif semantic.kind in {"additive", "count", "balance"} and numeric_ratio >= 0.6:
            # Business field names outrank low cardinality.  A quarterly total
            # column with only one value is still an amount, not a category.
            role = "数值指标"
        elif numeric_ratio >= 0.8:
            role = "数值指标"
        elif unique_count <= min(50, max(5, math.ceil(rows * 0.25))) and unique_count >= 2:
            role = "分类维度"
        elif uniqueness >= 0.9 and non_missing_count >= 3:
            role = "文本标识"
        else:
            role = "文本描述"
        roles[name] = {
            "role": role,
            "numeric_ratio": numeric_ratio,
            "date_ratio": date_ratio,
            "unique_count": unique_count,
            "uniqueness": uniqueness,
            "missing_rate": float(1 - non_missing_count / rows),
            "percent_text": percent_text,
            "numeric_values": numeric_values,
            "date_values": date_values,
        }
    return roles


def _clean_frame(frame: pd.DataFrame, roles: Mapping[str, Mapping[str, Any]] | None = None) -> tuple[pd.DataFrame, int]:
    work = frame.copy(deep=True)
    work.columns = [str(column).strip() or f"未命名字段{index + 1}" for index, column in enumerate(work.columns)]
    work = work.dropna(axis=0, how="all").dropna(axis=1, how="all")
    # Re-infer after header cleanup so integer/blank/whitespace headers cannot
    # make the role dictionary drift away from the actual columns.
    roles = infer_column_roles(work)
    for column in work.columns:
        role = roles.get(column, {}).get("role")
        if role == "日期":
            work[column] = pd.to_datetime(work[column], errors="coerce", yearfirst=True, format="mixed")
        elif role in {"数值指标", "比例/评分"}:
            numeric_values, ratio, _ = _numeric(work[column])
            if ratio >= 0.6:
                work[column] = numeric_values
        elif role not in {"标识符", "文本标识"}:
            work[column] = work[column].map(_text)
    before = len(work)
    work = work.drop_duplicates(keep="first").reset_index(drop=True)
    return work, before - len(work)


def _schema(frame: pd.DataFrame) -> tuple[str, ...]:
    return tuple(sorted(_normalise(column) for column in frame.columns if _normalise(column)))


def _select_primary(frames: Sequence[pd.DataFrame], names: Sequence[str] | None = None) -> int:
    scores = []
    for index, frame in enumerate(frames):
        roles = infer_column_roles(frame)
        role_values = [item["role"] for item in roles.values()]
        score = (
            min(len(frame), 100_000) * 0.01 + len(frame.columns) * 2
            + role_values.count("数值指标") * 12 + role_values.count("比例/评分") * 8
            + role_values.count("日期") * 12 + role_values.count("分类维度") * 7
            + role_values.count("标识符") * 5
        )
        if len(frame) < 2 or len(frame.columns) < 2:
            score -= 50
        sheet_role = classify_sheet_role(names[index] if names and index < len(names) else f"表{index + 1}", frame)
        if sheet_role == "summary":
            score -= 500
        elif sheet_role == "notes":
            score -= 250
        elif sheet_role == "dimension":
            score -= 30
        scores.append(score)
    return int(max(range(len(scores)), key=scores.__getitem__))


def can_build_adaptive_report(frames: Sequence[pd.DataFrame]) -> bool:
    return bool(frames) and any(isinstance(frame, pd.DataFrame) and not frame.empty and frame.shape[1] >= 1 for frame in frames)


def validate_adaptive_report_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("通用自适应报告参数必须是对象")
    names = params.get("source_names")
    if not isinstance(names, (list, tuple)) or not names or not all(isinstance(item, str) and item.strip() for item in names):
        raise TypeError("source_names 必须是非空字符串数组")
    request = params.get("user_request", "")
    if not isinstance(request, str) or len(request) > 8_000:
        raise TypeError("user_request 必须是不超过8000字符的文本")
    top_n = params.get("top_n", 10)
    if isinstance(top_n, bool) or not isinstance(top_n, int) or not 3 <= top_n <= 30:
        raise ValueError("top_n 必须是3到30之间的整数")
    multiplier = params.get("outlier_multiplier", 1.5)
    if isinstance(multiplier, bool) or not isinstance(multiplier, (int, float)) or not 1 <= float(multiplier) <= 10:
        raise ValueError("outlier_multiplier 必须是1到10之间的数字")


@dataclass(frozen=True)
class AdaptiveReportResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def _metric_priority(name: str) -> tuple[int, str]:
    if re.search(r"利润|profit", name, re.I): return (1, name)
    if re.search(r"销售|收入|金额|revenue|sales|amount", name, re.I): return (2, name)
    if re.search(r"成本|费用|工资|薪资|cost|expense|salary", name, re.I): return (3, name)
    if _QUANTITY_HINT.search(name): return (4, name)
    if _RATE_HINT.search(name): return (5, name)
    return (9, name)


def build_adaptive_analysis_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    user_request: str = "",
    top_n: int = 10,
    outlier_multiplier: float = 1.5,
) -> AdaptiveReportResult:
    """Build an adaptive nine-sheet management report from unfamiliar tables."""

    validate_adaptive_report_params({
        "source_names": source_names, "user_request": user_request,
        "top_n": top_n, "outlier_multiplier": outlier_multiplier,
    })
    if len(frames) != len(source_names):
        raise ValueError("source_names 数量必须与输入表数量一致")
    valid = [(index, frame) for index, frame in enumerate(frames) if isinstance(frame, pd.DataFrame) and not frame.empty]
    if not valid:
        raise ValueError("没有可分析的非空数据表")

    valid_frames = [item[1] for item in valid]
    valid_names = [source_names[item[0]] for item in valid]
    compiled = compile_analysis(valid_frames, source_names=valid_names, user_request=user_request)
    sheet_roles = [classify_sheet_role(name, frame) for name, frame in zip(valid_names, valid_frames)]
    primary_local = compiled.primary_index
    primary_schema = _schema(valid_frames[primary_local])
    compatible = [
        index
        for index, frame in enumerate(valid_frames)
        if _schema(frame) == primary_schema and sheet_roles[index] == "fact"
    ]
    if primary_local not in compatible:
        compatible = [primary_local]

    cleaned_frames: list[pd.DataFrame] = []
    table_roles: list[dict[str, dict[str, Any]]] = []
    duplicate_counts: list[int] = []
    for frame in valid_frames:
        roles = infer_column_roles(frame)
        clean, duplicates = _clean_frame(frame, roles)
        cleaned_frames.append(clean)
        table_roles.append(infer_column_roles(clean))
        duplicate_counts.append(duplicates)

    primary_name = valid_names[primary_local]
    provenance_column = "来源数据表"
    while provenance_column in cleaned_frames[primary_local].columns:
        provenance_column = f"_{provenance_column}"
    primary_parts = []
    for index in compatible:
        part = cleaned_frames[index].copy(deep=True)
        part.insert(0, provenance_column, valid_names[index])
        primary_parts.append(part)
    primary = pd.concat(primary_parts, ignore_index=True, sort=False)
    primary = primary.drop_duplicates(
        subset=[column for column in primary.columns if column != provenance_column],
        keep="first",
    ).reset_index(drop=True)
    primary_roles = infer_column_roles(primary.drop(columns=[provenance_column], errors="ignore"))

    dictionary_rows = []
    quality_rows = []
    role_labels = {
        "fact": "事实数据表",
        "dimension": "主数据/维度表",
        "summary": "汇总/报表表（不参与计算）",
        "notes": "说明/备注表（不参与计算）",
    }
    compiled_fields = {(field.table_index, field.field): field for field in compiled.fields}
    for table_index, (name, frame, roles, duplicate_count) in enumerate(zip(valid_names, cleaned_frames, table_roles, duplicate_counts)):
        cells = max(frame.shape[0] * frame.shape[1], 1)
        missing_cells = int(frame.isna().sum().sum() + frame.astype("string").apply(lambda col: col.str.strip().eq("").sum()).sum())
        semantic_role = sheet_roles[table_index]
        calculated_role = (
            "主分析事实表"
            if table_index == primary_local
            else ("同构合并事实表" if table_index in compatible else role_labels[semantic_role])
        )
        quality_rows.append({
            "数据表": name, "表角色": calculated_role,
            "行数": len(frame), "列数": frame.shape[1], "缺失单元格": missing_cells,
            "缺失率": min(missing_cells / cells, 1.0), "完全重复行": duplicate_count,
            "质量状态": (
                "参考表/不评分"
                if semantic_role in {"summary", "notes"}
                else ("需关注" if missing_cells / cells > 0.1 or duplicate_count else "良好")
            ),
        })
        for column in frame.columns:
            info = roles[column]
            binding = compiled_fields.get((table_index, str(column)))
            samples = [value for value in frame[column].dropna().map(_text).unique().tolist() if value][:3]
            dictionary_rows.append({
                "数据表": name, "字段": column, "推断角色": info["role"], "数据类型": str(frame[column].dtype),
                "标准概念": binding.concept if binding and binding.concept else "未映射",
                "聚合规则": binding.aggregation if binding and binding.aggregation != "none" else "不聚合",
                "语义置信度": binding.concept_confidence if binding else 0.0,
                "非空数": int(frame[column].notna().sum()), "缺失率": info["missing_rate"], "唯一值数": info["unique_count"],
                "唯一性": info["uniqueness"], "示例值": "｜".join(samples),
                "推断依据": "字段业务语义优先，其次为数据类型、可解析比例和基数分布",
            })

    relations = []
    for left in range(len(cleaned_frames)):
        for right in range(left + 1, len(cleaned_frames)):
            if sheet_roles[left] in {"summary", "notes"} or sheet_roles[right] in {"summary", "notes"}:
                continue
            left_map = {_normalise(column): column for column in cleaned_frames[left].columns}
            right_map = {_normalise(column): column for column in cleaned_frames[right].columns}
            shared = [key for key in left_map if key and key in right_map]
            for key in shared[:8]:
                lcol, rcol = left_map[key], right_map[key]
                lvals = set(cleaned_frames[left][lcol].dropna().map(_text)) - {""}
                rvals = set(cleaned_frames[right][rcol].dropna().map(_text)) - {""}
                if not lvals or not rvals:
                    continue
                coverage = len(lvals & rvals) / max(min(len(lvals), len(rvals)), 1)
                left_unique = cleaned_frames[left][lcol].dropna().astype("string").is_unique
                right_unique = cleaned_frames[right][rcol].dropna().astype("string").is_unique
                if coverage < 0.2:
                    continue
                relation = "一对一" if left_unique and right_unique else ("多对一" if right_unique else ("一对多" if left_unique else "多对多/需核验"))
                relations.append({
                    "左表": valid_names[left], "右表": valid_names[right], "关联字段": f"{lcol} ↔ {rcol}",
                    "建议关系": relation, "值覆盖率": coverage,
                    "置信度": "高" if coverage >= 0.8 and (left_unique or right_unique) else ("中" if coverage >= 0.5 else "低"),
                    "人工核验": "核对字段业务含义及重复键后再执行正式连接",
                })
    relation_frame = pd.DataFrame(relations, columns=["左表", "右表", "关联字段", "建议关系", "值覆盖率", "置信度", "人工核验"])

    inferred_metrics = sorted(
        [column for column, info in primary_roles.items() if info["role"] in {"数值指标", "比例/评分"}],
        key=_metric_priority,
    )
    metric_columns = [column for column in compiled.metrics if column in primary.columns]
    metric_columns.extend(column for column in inferred_metrics if column not in metric_columns)
    metric_columns = metric_columns[:12]
    date_columns = [column for column in compiled.dates if column in primary.columns]
    date_columns.extend(
        column for column, info in primary_roles.items() if info["role"] == "日期" and column not in date_columns
    )
    category_columns = [column for column in compiled.dimensions if column in primary.columns]
    category_columns.extend(
        column for column, info in primary_roles.items() if info["role"] == "分类维度" and column not in category_columns
    )
    identifier_columns = [column for column in compiled.identifiers if column in primary.columns]
    identifier_columns.extend(
        column for column, info in primary_roles.items() if info["role"] == "标识符" and column not in identifier_columns
    )
    key_column = next((column for column in identifier_columns if primary_roles[column]["uniqueness"] >= 0.8), identifier_columns[0] if identifier_columns else None)

    overview_rows = [
        {"指标": "分析主题", "结果": _text(user_request, limit=500) or "通用自适应数据分析", "单位": "", "数据口径": "用户自然语言需求；未提供时采用通用经营分析"},
        {"指标": "识别经营领域", "结果": compiled.domain_label, "单位": "", "数据口径": f"配置化领域词典匹配；置信度 {compiled.domain_confidence:.0%}"},
        {"指标": "已启用分析", "结果": "、".join(compiled.capabilities), "单位": "", "数据口径": "仅启用现有字段证据能够支持的分析能力"},
        {"指标": "证据缺口", "结果": "；".join(compiled.missing_evidence) or "未发现用户明确要求但证据不足的主题", "单位": "", "数据口径": "证据不足时不猜测、不强行计算"},
        {"指标": "主数据粒度", "结果": " + ".join(compiled.table_profiles[primary_local].grain) or "待人工确认", "单位": "", "数据口径": "按日期、标识符和分类维度推断"},
        {"指标": "上传数据表", "结果": len(valid_frames), "单位": "张", "数据口径": "非空工作表数量"},
        {"指标": "主分析表", "结果": primary_name, "单位": "", "数据口径": "按行列规模、数值/日期/分类字段丰富度自动选择"},
        {"指标": "同构合并表", "结果": len(compatible), "单位": "张", "数据口径": "字段集合完全一致的期间/分表自动纵向合并"},
        {"指标": "主数据记录数", "结果": len(primary), "单位": "行", "数据口径": "同构表合并并删除完全重复行"},
        {"指标": "识别数值指标", "结果": len(metric_columns), "单位": "个", "数据口径": "字段名、类型与数值解析率综合识别"},
        {"指标": "识别分类维度", "结果": len(category_columns), "单位": "个", "数据口径": "低至中等基数的分类字段"},
        {"指标": "识别时间字段", "结果": len(date_columns), "单位": "个", "数据口径": "字段名和日期解析率"},
        {"指标": "建议表关系", "结果": len(relation_frame), "单位": "条", "数据口径": "同名字段、值覆盖率和键唯一性推断"},
    ]
    for column in metric_columns[:6]:
        values = pd.to_numeric(primary[column], errors="coerce")
        if not values.notna().any():
            continue
        result, method, semantic = aggregate_metric(primary, column)
        unit = semantic.unit or ("%/分" if primary_roles[column]["role"] == "比例/评分" else "")
        overview_rows.append({"指标": f"核心指标：{column}", "结果": result, "单位": unit, "数据口径": method})
    overview = pd.DataFrame(overview_rows)

    ranking_rows = []
    ranking_metric = metric_columns[0] if metric_columns else None
    for dimension in category_columns[:4]:
        if ranking_metric:
            grouped = grouped_metric(primary, dimension, ranking_metric).rename(columns={ranking_metric: "指标值"})
            method = aggregate_metric(primary, ranking_metric)[1]
        else:
            grouped = primary.groupby(dimension, dropna=False, observed=True).size().reset_index(name="指标值")
            method = "记录数"
        grouped = grouped.dropna(subset=["指标值"]).sort_values("指标值", ascending=False, kind="stable").head(top_n)
        total = float(grouped["指标值"].sum())
        for rank, (_, row) in enumerate(grouped.iterrows(), start=1):
            ranking_rows.append({
                "分析维度": dimension, "分类": _text(row[dimension]) or "（空值）", "指标字段": ranking_metric or "记录数",
                "汇总方式": method, "指标值": float(row["指标值"]), "排名": rank,
                "占比": float(row["指标值"] / total) if total else float("nan"),
            })
    ranking = pd.DataFrame(ranking_rows, columns=["分析维度", "分类", "指标字段", "汇总方式", "指标值", "排名", "占比"])

    trend = pd.DataFrame(columns=["月份", *metric_columns[:3]])
    if date_columns and metric_columns:
        date_column = date_columns[0]
        trend_source = primary.copy(deep=True)
        trend_source[date_column] = pd.to_datetime(trend_source[date_column], errors="coerce")
        trend_source = trend_source.dropna(subset=[date_column])
        if not trend_source.empty:
            trend_source["月份"] = trend_source[date_column].dt.to_period("M").astype(str)
            trend = trend_source[["月份"]].drop_duplicates().sort_values("月份", kind="stable")
            for column in metric_columns[:3]:
                grouped = grouped_metric(trend_source, "月份", column)
                trend = trend.merge(grouped, on="月份", how="left")

    anomaly_rows = []
    for column in metric_columns[:8]:
        values = pd.to_numeric(primary[column], errors="coerce")
        valid_values = values.dropna()
        if len(valid_values) < 4:
            continue
        q1, q3 = valid_values.quantile([0.25, 0.75])
        iqr = q3 - q1
        if not math.isfinite(float(iqr)) or iqr <= 0:
            continue
        lower, upper = q1 - float(outlier_multiplier) * iqr, q3 + float(outlier_multiplier) * iqr
        mask = values.lt(lower) | values.gt(upper)
        for index in primary.index[mask][:2000]:
            anomaly_rows.append({
                "来源表": primary.at[index, provenance_column],
                "源行号": int(index) + 2, "记录标识": _text(primary.at[index, key_column]) if key_column else f"第{int(index)+2}行",
                "异常类型": "IQR数值异常", "异常字段": column, "异常值": values.at[index],
                "判定依据": f"低于{lower:,.4g}或高于{upper:,.4g}", "建议动作": "结合业务凭证确认是否为真实极值或录入错误",
            })
    for column in identifier_columns[:3]:
        missing_mask = primary[column].isna() | primary[column].astype("string").str.strip().eq("")
        for index in primary.index[missing_mask][:500]:
            anomaly_rows.append({
                "来源表": primary.at[index, provenance_column], "源行号": int(index) + 2,
                "记录标识": f"第{int(index)+2}行", "异常类型": "关键标识缺失", "异常字段": column, "异常值": "",
                "判定依据": "推断为标识符但单元格为空", "建议动作": "补齐标识或人工确认是否应排除",
            })
    anomalies = pd.DataFrame(anomaly_rows, columns=["来源表", "源行号", "记录标识", "异常类型", "异常字段", "异常值", "判定依据", "建议动作"])
    if compiled.missing_evidence:
        missing_rows = pd.DataFrame([
            {
                "来源表": "分析编译器", "源行号": "", "记录标识": f"证据缺口{index}",
                "异常类型": "分析证据不足", "异常字段": "", "异常值": "",
                "判定依据": message, "建议动作": "补充对应字段或业务口径后重新运行；当前报告不输出确定性结论",
            }
            for index, message in enumerate(compiled.missing_evidence, start=1)
        ])
        anomalies = pd.concat([anomalies, missing_rows], ignore_index=True)
    overview.loc[len(overview)] = {
        "指标": "检测异常记录", "结果": len(anomalies), "单位": "条",
        "数据口径": "IQR 数值异常与推断标识字段缺失；仅作核验线索，不替代业务判断",
    }

    quality = pd.DataFrame(quality_rows)
    dictionary = pd.DataFrame(dictionary_rows)
    risk_summary = (
        anomalies.groupby("异常类型", dropna=False, observed=True).size().reset_index(name="风险数量")
        if not anomalies.empty else pd.DataFrame(columns=["异常类型", "风险数量"])
    )
    dashboard_rows = max(len(ranking.head(top_n)), len(trend), len(risk_summary), 1)
    dashboard = pd.DataFrame(index=range(dashboard_rows))
    top_rank = ranking.loc[ranking["分析维度"].eq(ranking.iloc[0]["分析维度"])] if not ranking.empty else ranking
    for column in ("分类", "指标值"):
        dashboard[f"排名{column}"] = top_rank[column].reindex(range(dashboard_rows)) if column in top_rank else pd.NA
    if not trend.empty:
        dashboard["月份"] = trend["月份"].reindex(range(dashboard_rows))
        for column in metric_columns[:3]:
            if column in trend: dashboard[f"趋势_{column}"] = trend[column].reindex(range(dashboard_rows))
    if any(chart.kind == "composition" for chart in compiled.charts) and not top_rank.empty:
        dashboard["结构分类"] = top_rank["分类"].reindex(range(dashboard_rows))
        dashboard["结构指标值"] = top_rank["指标值"].reindex(range(dashboard_rows))
    if not risk_summary.empty:
        dashboard["异常类型"] = risk_summary["异常类型"].reindex(range(dashboard_rows))
        dashboard["风险数量"] = risk_summary["风险数量"].reindex(range(dashboard_rows))

    outputs = {
        "管理层通用总览": overview,
        "主数据分析": primary,
        "数据字典": dictionary,
        "数据质量": quality,
        "表关系建议": relation_frame,
        "分类排名": ranking,
        "时间趋势": trend,
        "异常数据": anomalies,
        "自适应图表看板": dashboard,
    }
    for output in outputs.values():
        output.attrs["toolbox_report_kind"] = "adaptive_analysis_report"
    report = {
        "source_table_count": len(valid_frames), "primary_table": primary_name,
        "domain_id": compiled.domain_id, "domain_label": compiled.domain_label,
        "domain_confidence": compiled.domain_confidence,
        "intent_topics": list(compiled.intent_topics), "capabilities": list(compiled.capabilities),
        "missing_evidence": list(compiled.missing_evidence), "analysis_plan": compiled.as_dict(),
        "combined_table_count": len(compatible), "primary_row_count": len(primary),
        "metric_count": len(metric_columns), "dimension_count": len(category_columns), "date_count": len(date_columns),
        "relation_count": len(relation_frame), "anomaly_count": len(anomalies),
        "sheet_count": len(outputs), "chart_count": len(compiled.charts),
    }
    return AdaptiveReportResult(outputs=outputs, report=report)


__all__ = [
    "AdaptiveReportResult", "build_adaptive_analysis_report", "can_build_adaptive_report",
    "infer_column_roles", "validate_adaptive_report_params",
]
