"""Compile unfamiliar workbook data and natural language into an analysis plan.

The compiler is deliberately domain-configurable.  Python code owns stable
behaviour (profiling, confidence, evidence gates and safe aggregations), while
``domain_packs.json`` owns vocabulary.  Adding a new industry therefore does
not require another request-specific route or report template.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
import json
import os
from pathlib import Path
import re
from typing import Any, Mapping, Sequence

import pandas as pd

try:
    from .metric_semantics import classify_metric, classify_sheet_role, normalise, ratio_components
except ImportError:  # pragma: no cover - direct module execution support
    from excel_data_toolbox.metric_semantics import classify_metric, classify_sheet_role, normalise, ratio_components


_PUNCTUATION = re.compile(r"[\s,，。；;：:、/\\|（）()【】\[\]{}<>《》]+")


@dataclass(frozen=True)
class FieldBinding:
    table_index: int
    table_name: str
    field: str
    role: str
    concept: str
    concept_confidence: float
    aggregation: str
    unit: str
    numerator: str = ""
    denominator: str = ""


@dataclass(frozen=True)
class TableProfile:
    index: int
    name: str
    role: str
    row_count: int
    column_count: int
    fact_score: float
    grain: tuple[str, ...]
    metrics: tuple[str, ...]
    dimensions: tuple[str, ...]
    dates: tuple[str, ...]
    identifiers: tuple[str, ...]


@dataclass(frozen=True)
class AnalysisSpec:
    kind: str
    title: str
    metric: str = ""
    dimension: str = ""
    date: str = ""
    reason: str = ""


@dataclass(frozen=True)
class ChartSpec:
    kind: str
    title: str
    metric: str
    dimension: str
    reason: str


@dataclass(frozen=True)
class CompiledAnalysisPlan:
    domain_id: str
    domain_label: str
    domain_confidence: float
    primary_index: int
    primary_table: str
    fact_indices: tuple[int, ...]
    fact_tables: tuple[str, ...]
    table_profiles: tuple[TableProfile, ...]
    fields: tuple[FieldBinding, ...]
    fact_metrics: tuple[FieldBinding, ...]
    fact_dimensions: tuple[FieldBinding, ...]
    fact_dates: tuple[FieldBinding, ...]
    metrics: tuple[str, ...]
    dimensions: tuple[str, ...]
    dates: tuple[str, ...]
    identifiers: tuple[str, ...]
    intent_topics: tuple[str, ...]
    capabilities: tuple[str, ...]
    missing_evidence: tuple[str, ...]
    analyses: tuple[AnalysisSpec, ...]
    charts: tuple[ChartSpec, ...]
    warnings: tuple[str, ...]

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


_TOPICS: Mapping[str, tuple[str, ...]] = {
    "overview": ("分析", "整体", "全面", "总览", "经营", "情况", "看看"),
    "profitability": ("利润", "赚钱", "毛利", "贡献", "盈利", "成本"),
    "trend": ("趋势", "月份", "增长", "下降", "变化", "同比", "环比"),
    "ranking": ("排名", "最好", "最差", "最高", "最低", "表现", "贡献"),
    "quality": ("清洗", "重复", "缺失", "格式", "口径", "数据质量", "无效"),
    "relationships": ("关联", "合并", "整合", "串联", "匹配", "全部看"),
    "anomaly": ("异常", "风险", "关注", "问题", "预警", "失控"),
    "inventory": ("库存", "积压", "缺货", "补货", "周转", "仓库"),
    "customer": ("客户", "顾客", "会员", "满意度", "投诉", "评价"),
    "workforce": ("员工", "人员", "人工", "工时", "加班", "绩效", "薪资"),
    "channel": ("渠道", "平台", "门店", "区域", "地区", "广告", "投放"),
    "cash": ("现金", "回款", "到账", "结算", "账龄", "应收"),
    "procurement": ("采购", "供应商", "入库", "原料", "物料"),
}


_CAPABILITY_REQUIREMENTS: Mapping[str, tuple[tuple[str, ...], ...]] = {
    "profitability": (("profit",), ("margin",), ("revenue", "cost")),
    "inventory": (("inventory",),),
    "customer": (("customer", "score"), ("customer", "refund"), ("customer", "revenue")),
    "workforce": (("employee",), ("workforce",)),
    "channel": (("channel",), ("store",), ("region",)),
    "cash": (("cash",), ("balance",)),
    "procurement": (("purchase",), ("supplier",)),
}


def _domain_path(path: str | os.PathLike[str] | None = None) -> Path:
    if path:
        return Path(path)
    configured = os.getenv("ANALYSIS_DOMAIN_PACKS", "").strip()
    return Path(configured) if configured else Path(__file__).with_name("domain_packs.json")


def load_domain_packs(path: str | os.PathLike[str] | None = None) -> tuple[dict[str, Any], ...]:
    """Load and validate configurable domain vocabulary."""

    source = _domain_path(path)
    try:
        payload = json.loads(source.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"领域能力包无法读取：{source}") from exc
    if not isinstance(payload, Mapping) or payload.get("schema_version") != 1:
        raise ValueError("领域能力包 schema_version 必须为 1")
    domains = payload.get("domains")
    if not isinstance(domains, list) or not domains:
        raise ValueError("领域能力包必须包含非空 domains 数组")
    validated: list[dict[str, Any]] = []
    for index, item in enumerate(domains):
        if not isinstance(item, Mapping):
            raise ValueError(f"domains[{index}] 必须是对象")
        domain_id = item.get("id")
        label = item.get("label")
        anchors = item.get("anchors")
        identity_anchors = item.get("identity_anchors", [])
        concepts = item.get("concepts")
        if not isinstance(domain_id, str) or not domain_id.strip():
            raise ValueError(f"domains[{index}].id 必须是非空字符串")
        if not isinstance(label, str) or not label.strip():
            raise ValueError(f"domains[{index}].label 必须是非空字符串")
        if not isinstance(anchors, list) or not all(isinstance(value, str) and value.strip() for value in anchors):
            raise ValueError(f"domains[{index}].anchors 必须是字符串数组")
        if not isinstance(identity_anchors, list) or not all(
            isinstance(value, str) and value.strip() for value in identity_anchors
        ):
            raise ValueError(f"domains[{index}].identity_anchors 必须是字符串数组")
        if not isinstance(concepts, Mapping) or not concepts:
            raise ValueError(f"domains[{index}].concepts 必须是非空对象")
        normalised_concepts: dict[str, list[str]] = {}
        for concept, aliases in concepts.items():
            if not isinstance(concept, str) or not concept.strip() or not isinstance(aliases, list):
                raise ValueError(f"domains[{index}].concepts 字段结构无效")
            values = [str(value).strip() for value in aliases if isinstance(value, str) and value.strip()]
            if not values:
                raise ValueError(f"domains[{index}].concepts.{concept} 不能为空")
            normalised_concepts[concept.strip()] = values
        validated.append({
            "id": domain_id.strip(),
            "label": label.strip(),
            "anchors": [value.strip() for value in anchors],
            "identity_anchors": [value.strip() for value in identity_anchors],
            "concepts": normalised_concepts,
        })
    return tuple(validated)


def _numeric_ratio(series: pd.Series) -> float:
    if pd.api.types.is_bool_dtype(series.dtype):
        return 0.0
    if pd.api.types.is_numeric_dtype(series.dtype):
        return float(series.notna().mean())
    raw = series.astype("string").str.strip()
    cleaned = raw.str.replace(r"[¥￥$,，元件个台套箱人次天小时%％\s]", "", regex=True)
    denominator = int(raw.ne("").sum())
    return float(pd.to_numeric(cleaned, errors="coerce").notna().sum() / denominator) if denominator else 0.0


def _date_ratio(series: pd.Series) -> float:
    if pd.api.types.is_datetime64_any_dtype(series.dtype):
        return float(series.notna().mean())
    raw = series.astype("string").str.strip()
    denominator = int(raw.ne("").sum())
    if not denominator:
        return 0.0
    return float(pd.to_datetime(raw, errors="coerce", yearfirst=True, format="mixed").notna().sum() / denominator)


def _column_role(name: str, series: pd.Series) -> str:
    semantic = classify_metric(name)
    numeric_ratio = _numeric_ratio(series)
    if semantic.kind == "identifier":
        return "identifier"
    if semantic.kind == "date" and _date_ratio(series) >= 0.5:
        return "date"
    if semantic.kind in {"additive", "count", "balance", "ratio", "score"} and numeric_ratio >= 0.5:
        return "metric"
    if numeric_ratio >= 0.8:
        return "metric"
    non_empty = series.dropna().astype("string").str.strip()
    unique = int(non_empty[non_empty.ne("")].nunique())
    if 1 < unique <= min(100, max(8, int(max(len(series), 1) * 0.35))):
        return "dimension"
    if unique and unique == int(non_empty[non_empty.ne("")].shape[0]):
        return "identifier"
    return "description"


def _domain_score(domain: Mapping[str, Any], prompt: str, names: Sequence[str], frames: Sequence[pd.DataFrame]) -> float:
    prompt_text = normalise(prompt)
    source_text = normalise(" ".join(names))
    field_text = normalise(" ".join(str(column) for frame in frames for column in frame.columns))
    score = 0.0
    for anchor in domain.get("identity_anchors", []):
        token = normalise(anchor)
        if not token:
            continue
        score += 20.0 if token in prompt_text else 0.0
        score += 8.0 if token in source_text else 0.0
        score += 3.0 if token in field_text else 0.0
    for anchor in domain["anchors"]:
        token = normalise(anchor)
        if not token:
            continue
        score += 4.0 if token in prompt_text else 0.0
        score += 2.0 if token in source_text else 0.0
        score += 1.0 if token in field_text else 0.0
    for aliases in domain["concepts"].values():
        if any(normalise(alias) in field_text for alias in aliases if normalise(alias)):
            score += 0.35
    return score


def _select_domain(
    packs: Sequence[Mapping[str, Any]], prompt: str, names: Sequence[str], frames: Sequence[pd.DataFrame]
) -> tuple[Mapping[str, Any], float]:
    scores = [_domain_score(pack, prompt, names, frames) for pack in packs]
    best_index = max(range(len(scores)), key=scores.__getitem__)
    best = scores[best_index]
    runner_up = sorted(scores, reverse=True)[1] if len(scores) > 1 else 0.0
    confidence = min(0.99, max(0.15, 0.45 + (best - runner_up) / max(best + 4.0, 8.0))) if best else 0.15
    return packs[best_index], confidence


def _match_concept(field: str, concepts: Mapping[str, Sequence[str]]) -> tuple[str, float]:
    key = normalise(field)
    exact: list[tuple[str, int]] = []
    partial: list[tuple[str, int]] = []
    for concept, aliases in concepts.items():
        for alias in aliases:
            candidate = normalise(alias)
            if not candidate:
                continue
            if key == candidate:
                exact.append((concept, len(candidate)))
            elif candidate in key or key in candidate:
                partial.append((concept, min(len(candidate), len(key))))
    if exact:
        return max(exact, key=lambda item: item[1])[0], 1.0
    if partial:
        return max(partial, key=lambda item: item[1])[0], 0.78
    semantic = classify_metric(field)
    fallback = {
        "date": "date", "identifier": "identifier", "additive": "metric",
        "count": "quantity", "balance": "inventory", "ratio": "ratio", "score": "score",
    }.get(semantic.kind, "")
    return fallback, 0.35 if fallback else 0.0


def _intent_topics(prompt: str) -> tuple[str, ...]:
    folded = _PUNCTUATION.sub("", str(prompt or "")).casefold()
    topics = [topic for topic, tokens in _TOPICS.items() if any(token.casefold() in folded for token in tokens)]
    if not topics:
        topics = ["overview"]
    if any(token in folded for token in ("全面", "完整", "全部", "老板", "经营诊断")):
        for topic in ("overview", "profitability", "trend", "ranking", "quality", "relationships", "anomaly"):
            if topic not in topics:
                topics.append(topic)
    return tuple(topics)


def _has_any(concepts: set[str], groups: Sequence[Sequence[str]]) -> bool:
    return any(all(item in concepts for item in group) for group in groups)


def _metric_sort_key(binding: FieldBinding) -> tuple[int, int, str]:
    concept_priority = {
        "profit": 1, "margin": 2, "revenue": 3, "cash": 4, "cost": 5,
        "refund": 6, "inventory": 7, "purchase": 8, "quantity": 9, "score": 10,
    }
    aggregation_priority = {"sum": 1, "weighted_ratio": 2, "last": 3, "mean": 4}
    return (concept_priority.get(binding.concept, 20), aggregation_priority.get(binding.aggregation, 9), binding.field)


def compile_analysis(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    user_request: str = "",
    domain_pack_path: str | os.PathLike[str] | None = None,
) -> CompiledAnalysisPlan:
    """Create an evidence-gated plan for any non-empty workbook structure."""

    if len(frames) != len(source_names):
        raise ValueError("source_names 数量必须与输入表数量一致")
    valid = [(index, frame, str(source_names[index])) for index, frame in enumerate(frames) if isinstance(frame, pd.DataFrame) and not frame.empty]
    if not valid:
        raise ValueError("没有可编译的非空数据表")
    valid_frames = [item[1] for item in valid]
    valid_names = [item[2] for item in valid]
    packs = load_domain_packs(domain_pack_path)
    domain, domain_confidence = _select_domain(packs, user_request, valid_names, valid_frames)

    fields: list[FieldBinding] = []
    profiles: list[TableProfile] = []
    for local_index, frame in enumerate(valid_frames):
        name = valid_names[local_index]
        sheet_role = classify_sheet_role(name, frame)
        metrics: list[str] = []
        dimensions: list[str] = []
        dates: list[str] = []
        identifiers: list[str] = []
        for raw_column in frame.columns:
            column = str(raw_column)
            role = _column_role(column, frame[raw_column])
            concept, concept_confidence = _match_concept(column, domain["concepts"])
            semantic = classify_metric(column)
            components = ratio_components(column, frame.columns) if semantic.aggregation == "weighted_ratio" else None
            fields.append(FieldBinding(
                table_index=local_index, table_name=name, field=column, role=role,
                concept=concept, concept_confidence=concept_confidence,
                aggregation=semantic.aggregation if role == "metric" else "none",
                unit=semantic.unit, numerator=components[0] if components else "",
                denominator=components[1] if components else "",
            ))
            if role == "metric": metrics.append(column)
            elif role == "dimension": dimensions.append(column)
            elif role == "date": dates.append(column)
            elif role == "identifier": identifiers.append(column)
        score = (
            min(len(frame), 100_000) * 0.01 + len(frame.columns) * 1.5 + len(metrics) * 12
            + len(dates) * 12 + len(dimensions) * 7 + len(identifiers) * 4
        )
        if sheet_role == "summary": score -= 600
        elif sheet_role == "notes": score -= 350
        elif sheet_role == "dimension": score -= 35
        grain = tuple((dates + identifiers + dimensions)[:3])
        profiles.append(TableProfile(
            index=local_index, name=name, role=sheet_role, row_count=len(frame), column_count=frame.shape[1],
            fact_score=score, grain=grain, metrics=tuple(metrics), dimensions=tuple(dimensions),
            dates=tuple(dates), identifiers=tuple(identifiers),
        ))

    primary_index = max(range(len(profiles)), key=lambda index: profiles[index].fact_score)
    primary = profiles[primary_index]
    fact_indices = tuple(profile.index for profile in profiles if profile.role == "fact")
    if not fact_indices:
        fact_indices = (primary_index,)
    fact_tables = tuple(profiles[index].name for index in fact_indices)
    primary_fields = [field for field in fields if field.table_index == primary_index]
    metric_bindings = sorted((field for field in primary_fields if field.role == "metric"), key=_metric_sort_key)
    dimension_bindings = [field for field in primary_fields if field.role == "dimension"]
    date_bindings = [field for field in primary_fields if field.role == "date"]
    identifier_bindings = [field for field in primary_fields if field.role == "identifier"]
    concepts = {field.concept for field in fields if field.concept}
    topics = _intent_topics(user_request)
    fact_fields = [field for field in fields if field.table_index in fact_indices]
    fact_metric_bindings = sorted(
        (field for field in fact_fields if field.role == "metric"),
        key=lambda field: (field.table_index, _metric_sort_key(field)),
    )
    fact_dimension_bindings = [field for field in fact_fields if field.role == "dimension"]
    fact_date_bindings = [field for field in fact_fields if field.role == "date"]

    capabilities = {"overview", "quality", "anomaly"}
    if fact_metric_bindings and fact_dimension_bindings:
        capabilities.add("ranking")
    if fact_metric_bindings and fact_date_bindings:
        capabilities.add("trend")
    if len(valid_frames) > 1:
        capabilities.add("relationships")
    for capability, requirements in _CAPABILITY_REQUIREMENTS.items():
        if _has_any(concepts, requirements):
            capabilities.add(capability)
    if "profit" in concepts or "margin" in concepts or {"revenue", "cost"}.issubset(concepts):
        capabilities.add("profitability")

    missing: list[str] = []
    evidence_messages = {
        "profitability": "盈利判断缺少利润/利润率，或收入与成本的可比口径",
        "trend": "趋势分析缺少可解析日期和数值指标",
        "ranking": "排名分析缺少可分组维度或数值指标",
        "relationships": "跨表关联需要至少两张有效业务表",
        "inventory": "库存判断缺少库存数量/金额或库存状态字段",
        "customer": "客户判断缺少客户、评价或退款证据",
        "workforce": "人员判断缺少员工、部门、工时或绩效证据",
        "channel": "渠道判断缺少平台、门店、地区或渠道维度",
        "cash": "资金判断缺少回款、到账、余额或现金字段",
        "procurement": "采购判断缺少采购、供应商或入库字段",
    }
    for topic in topics:
        if topic in evidence_messages and topic not in capabilities:
            missing.append(evidence_messages[topic])

    metrics = tuple(field.field for field in metric_bindings[:12])
    dimensions = tuple(field.field for field in dimension_bindings[:8])
    dates = tuple(field.field for field in date_bindings[:4])
    identifiers = tuple(field.field for field in identifier_bindings[:6])
    analyses: list[AnalysisSpec] = [AnalysisSpec("overview", "核心经营指标", reason="按指标语义选择求和、期末、平均或加权比率")]
    for index in fact_indices:
        if index == primary_index:
            continue
        profile = profiles[index]
        safe_metric = next(
            (
                field.field
                for field in fields
                if field.table_index == index and field.role == "metric" and field.aggregation != "unknown"
            ),
            "",
        )
        analyses.append(
            AnalysisSpec(
                "fact_overview",
                f"{profile.name}事实域摘要",
                metric=safe_metric,
                reason="多事实图保留各事实表原始粒度，不把所有指标压到单一主表",
            )
        )
    if "ranking" in capabilities:
        analyses.append(AnalysisSpec("ranking", f"{dimensions[0]}表现排名", metric=metrics[0], dimension=dimensions[0], reason="存在可分组维度和可聚合指标"))
    if "trend" in capabilities:
        analyses.append(AnalysisSpec("trend", f"{metrics[0]}时间趋势", metric=metrics[0], date=dates[0], reason="存在可解析时间字段和可聚合指标"))
    if "quality" in topics or "quality" in capabilities:
        analyses.append(AnalysisSpec("quality", "数据质量与口径审计", reason="每次交付必须披露缺失、重复、粒度与人工核验边界"))
    if "relationships" in capabilities:
        analyses.append(AnalysisSpec("relationships", "跨表关系建议", reason="多表输入需要识别候选连接键和粒度风险"))
    if "anomaly" in capabilities:
        analyses.append(AnalysisSpec("anomaly", "异常与风险线索", metric=metrics[0] if metrics else "", reason="只输出可追溯的数据异常，不替代业务审批"))

    charts: list[ChartSpec] = []
    if "ranking" in capabilities:
        charts.append(ChartSpec("bar", f"{dimensions[0]}—{metrics[0]}排名", metrics[0], dimensions[0], "比较类别表现"))
        if len(dimension_bindings) and classify_metric(metrics[0]).aggregation == "sum":
            charts.append(ChartSpec("composition", f"{metrics[0]}结构占比", metrics[0], dimensions[0], "同一可加指标的结构分解"))
    if "trend" in capabilities:
        charts.append(ChartSpec("line", f"{metrics[0]}时间趋势", metrics[0], dates[0], "观察跨期变化"))
    if "anomaly" in capabilities and metrics:
        charts.append(ChartSpec("risk", "异常线索分布", metrics[0], "异常类型", "按规则聚合异常线索数量"))

    warnings: list[str] = []
    if domain_confidence < 0.55:
        warnings.append("领域识别置信度较低；仍按字段证据执行通用分析，领域术语需人工复核")
    if primary.role in {"summary", "notes"}:
        warnings.append("未找到可靠事实表，当前主表疑似汇总或说明表")
    if not metrics:
        warnings.append("主分析表未识别到可靠数值指标，仅生成结构和数据质量审计")
    return CompiledAnalysisPlan(
        domain_id=str(domain["id"]), domain_label=str(domain["label"]), domain_confidence=domain_confidence,
        primary_index=primary_index, primary_table=primary.name, fact_indices=fact_indices, fact_tables=fact_tables,
        table_profiles=tuple(profiles), fields=tuple(fields),
        fact_metrics=tuple(fact_metric_bindings), fact_dimensions=tuple(fact_dimension_bindings),
        fact_dates=tuple(fact_date_bindings),
        metrics=metrics, dimensions=dimensions, dates=dates, identifiers=identifiers,
        intent_topics=topics, capabilities=tuple(sorted(capabilities)), missing_evidence=tuple(dict.fromkeys(missing)),
        analyses=tuple(analyses), charts=tuple(charts), warnings=tuple(warnings),
    )


__all__ = [
    "AnalysisSpec", "ChartSpec", "CompiledAnalysisPlan", "FieldBinding", "TableProfile",
    "compile_analysis", "load_domain_packs",
]
