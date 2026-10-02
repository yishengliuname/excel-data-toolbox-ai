"""Local business evidence contracts and a non-joining fact relationship graph.

Names provide hypotheses, never confirmed accounting policy. A validated
physical key is not proof that two business amounts have the same population.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, replace
from itertools import combinations
import math
import re
from typing import Mapping, Sequence

import pandas as pd

from .metric_semantics import aggregate_metric, classify_metric, normalise, ratio_components


@dataclass(frozen=True)
class MetricContract:
    table_index: int
    table_name: str
    field: str
    aggregation: str
    unit: str
    grain: tuple[str, ...]
    date_field: str
    time_basis: str
    numerator: str
    denominator: str
    population: str
    valid_rows: int
    total_rows: int
    status: str
    issues: tuple[str, ...]
    origin: str = "local_inference"
    version: int = 1


@dataclass(frozen=True)
class RelationshipContract:
    left_index: int
    right_index: int
    left_table: str
    right_table: str
    left_key: str
    right_key: str
    cardinality: str
    left_row_coverage: float
    right_row_coverage: float
    left_key_coverage: float
    right_key_coverage: float
    status: str
    reason: str


@dataclass(frozen=True)
class QuestionEvidence:
    topic: str
    status: str
    sources: tuple[str, ...]
    evidence: tuple[str, ...]
    missing: tuple[str, ...]
    conclusion_boundary: str


def numeric_values(series: pd.Series) -> pd.Series:
    """Parse notation, not currencies or physical-unit conversions."""
    if pd.api.types.is_bool_dtype(series.dtype):
        return pd.Series(float("nan"), index=series.index)
    if pd.api.types.is_numeric_dtype(series.dtype):
        return pd.to_numeric(series, errors="coerce")
    raw = series.astype("string").str.strip()
    percent = raw.str.contains(r"[%％]", na=False)
    clean = raw.str.replace(r"[¥￥,，%％\s]", "", regex=True)
    values = pd.to_numeric(clean, errors="coerce")
    return values.where(~percent, values / 100.0)


def candidate_grain(frame: pd.DataFrame) -> tuple[str, ...]:
    """Return a tested physical key, not a guessed row grain."""
    identifiers = [str(c) for c in frame if classify_metric(c).kind == "identifier"]
    identifiers = identifiers[:6]
    dates = [str(c) for c in frame if classify_metric(c).kind == "date"][:2]
    candidates = list(dict.fromkeys([*identifiers, *dates]))
    for size in range(1, min(3, len(candidates)) + 1):
        for keys in combinations(candidates, size):
            work = frame[list(keys)]
            if work.isna().any(axis=None) or work.astype("string").apply(lambda s: s.str.strip().eq("")).any(axis=None):
                continue
            if not work.duplicated().any():
                return tuple(keys)
    return ()


def metric_contract(index: int, name: str, frame: pd.DataFrame, field: str) -> MetricContract:
    semantic = classify_metric(field)
    values = numeric_values(frame[field])
    valid = values.notna() & values.map(lambda v: math.isfinite(float(v)) if pd.notna(v) else False)
    dates = [str(c) for c in frame if classify_metric(c).kind == "date"]
    date = dates[0] if len(dates) == 1 else ""
    issues = []
    if semantic.aggregation in {"none", "unknown"}:
        issues.append("业务聚合语义未确认")
    if frame.duplicated().any():
        issues.append("发现完全重复记录；需确认业务键及去重规则，原记录保留")
    if any(re.search(r"币种|currency|数量单位|计量单位|采购单位", str(c), re.I) and
           (frame[c].isna().any() or frame[c].astype("string").str.strip().eq("").any() or
            frame[c].dropna().astype("string").str.strip().nunique() > 1) for c in frame):
        issues.append("存在多币种或多计量单位；缺少统一换算规则")
    if int(valid.sum()) != len(frame):
        issues.append("存在缺失、不可解析或非有限数值；不按零填补")
    if len(dates) > 1:
        issues.append("存在多个时间字段；须确认发生月或订单归属月")
    if dates and not date:
        time_basis = "待确认"
    elif date:
        time_basis = "源字段：" + date
    else:
        time_basis = "未提供期间"
    if date and pd.to_datetime(frame[date], errors="coerce", format="mixed").isna().any():
        issues.append("时间字段存在缺失或无效日期")
    components = ratio_components(field, frame.columns) if semantic.aggregation == "weighted_ratio" else None
    if semantic.aggregation == "weighted_ratio":
        if components is None:
            issues.append("缺少唯一可确认的分子与分母")
        elif frame[list(components)].apply(numeric_values).isna().any(axis=None):
            issues.append("分子分母样本不完整，禁止使用不同样本计算比例")
        else:
            component_values = frame[list(components)].apply(numeric_values)
            scales = ["万元" if "万元" in c else "千元" if "千元" in c else "元" for c in components]
            if len(set(scales)) > 1:
                issues.append("分子分母单位数量级不同，须确认换算后才能计算比例")
            if not component_values.map(lambda x: math.isfinite(float(x))).all(axis=None):
                issues.append("分子分母包含非有限数值")
            elif float(component_values[components[1]].sum()) <= 0:
                issues.append("分母合计非正，比例无可解释口径")
    if semantic.aggregation == "last" and not date:
        issues.append("余额缺少唯一时间字段，无法确定期末")
    # Unit values must not be accumulated like transactions.
    if re.search(r"单价|平均|周转天数|库存天数|每单|每人|unit.?price|average|per.?unit", field, re.I):
        issues.append("单价/平均/效率指标需要权重或专门公式，禁止直接求和")
    if re.search(r"USD|EUR|美元|欧元|港币", field, re.I):
        issues.append("外币字段需要明确币种与换算规则，不默认人民币")
    status = "unavailable" if issues else "inferred"
    population = "全部源记录（状态规则未提供）"
    if any(re.search(r"状态|status", str(c), re.I) for c in frame):
        population = "全部源记录；正常经营范围待确认状态规则"
        issues.append("包含状态字段，需确认有效记录范围；未自动过滤")
        status = "unavailable"
    unit = "万元" if "万元" in field else "千元" if "千元" in field else semantic.unit
    if not unit:
        if semantic.kind == "additive":
            unit = "金额单位未声明"
        elif semantic.kind == "count":
            unit = "小时" if re.search(r"小时|工时|hours", field, re.I) else "天" if re.search(r"天数|days", field, re.I) else "数量单位未声明"
        elif semantic.kind == "balance":
            unit = "余额单位未声明"
    contract = MetricContract(index, name, field, semantic.aggregation, unit,
                          candidate_grain(frame), date, time_basis,
                          components[0] if components else "", components[1] if components else "",
                          population, int(valid.sum()), len(frame), status, tuple(issues))
    if contract.status != "unavailable" and contract.aggregation == "last":
        value, reason = evaluate_metric(frame, contract)
        if not math.isfinite(value):
            contract = replace(contract, status="unavailable", issues=(*contract.issues, reason))
    return contract


def evaluate_metric(frame: pd.DataFrame, contract: MetricContract) -> tuple[float, str]:
    if contract.status == "unavailable":
        return float("nan"), "；".join(contract.issues)
    if contract.aggregation == "last":
        # Global last-row selection would silently lose other entity balances.
        entities = [c for c in frame if classify_metric(c).kind == "identifier"]
        work = frame.copy()
        work[contract.date_field] = pd.to_datetime(work[contract.date_field], errors="coerce", format="mixed")
        latest = work[work[contract.date_field].eq(work[contract.date_field].max())]
        if entities:
            if latest[entities].isna().any(axis=None) or latest.duplicated(subset=entities).any():
                return float("nan"), "期末实体键重复或缺失，需要人工核验"
            if len(latest[entities].drop_duplicates()) != len(work[entities].drop_duplicates()):
                return float("nan"), "最新时点未覆盖全部实体，禁止混合不同日期的余额"
        elif len(latest) != 1:
            return float("nan"), "期末存在多条记录但缺少实体键，禁止选最后一行"
        return float(numeric_values(latest[contract.field]).sum(min_count=1)), "统一最新时点的实体余额合计"
    work = frame.copy(deep=True)
    for column in (contract.field, contract.numerator, contract.denominator):
        if column:
            work[column] = numeric_values(work[column])
    value, method, _ = aggregate_metric(work, contract.field)
    return value, method


def relationship_graph(frames: Sequence[pd.DataFrame], names: Sequence[str], roles: Sequence[str]) -> tuple[RelationshipContract, ...]:
    edges = []
    for left, lf in enumerate(frames):
        if roles[left] not in {"fact", "dimension"}:
            continue
        def key_column(column) -> bool:
            kind = classify_metric(column).kind
            return kind == "identifier" or (kind not in {"additive", "count", "balance", "ratio", "score", "date"} and
                bool(re.search(r"门店|客户|商品|产品|地区|区域|仓库|员工|部门|渠道|store|customer|product|region", str(column), re.I)))
        lm = {normalise(c): str(c) for c in lf if key_column(c)}
        for right in range(left + 1, len(frames)):
            if roles[right] not in {"fact", "dimension"}:
                continue
            rf = frames[right]
            rm = {normalise(c): str(c) for c in rf if key_column(c)}
            for key in sorted(set(lm) & set(rm))[:8]:
                lc, rc = lm[key], rm[key]
                ls, rs = lf[lc].astype("string").str.strip(), rf[rc].astype("string").str.strip()
                lv, rv = ls.notna() & ls.ne(""), rs.notna() & rs.ne("")
                lkeys, rkeys = set(ls[lv]), set(rs[rv])
                if not lkeys or not rkeys:
                    continue
                lu, ru = ls[lv].is_unique, rs[rv].is_unique
                cardinality = "one_to_one" if lu and ru else "many_to_one" if ru else "one_to_many" if lu else "many_to_many"
                intersect = lkeys & rkeys
                reason = "候选键业务含义需确认；未执行跨表连接"
                status = "candidate"
                if not lv.all() or not rv.all():
                    status, reason = "blocked", "关联键缺失；禁止空键互相匹配"
                elif cardinality == "many_to_many":
                    status, reason = "blocked", "多对多关系；须明确粒度并分别汇总，禁止直接连接后求和"
                elif roles[left] == roles[right] == "fact":
                    status, reason = "review", "两个事实域；即使物理键唯一也不能证明金额口径可比"
                if not intersect:
                    status, reason = "blocked", "键值无交集；需确认数据源、编码与业务范围"
                edges.append(RelationshipContract(left, right, names[left], names[right], lc, rc, cardinality,
                    float((lv & ls.isin(rkeys)).sum() / max(len(ls), 1)),
                    float((rv & rs.isin(lkeys)).sum() / max(len(rs), 1)),
                    len(intersect) / len(lkeys), len(intersect) / len(rkeys), status, reason))
    return tuple(edges)


def question_evidence(topics: Sequence[str], contracts: Sequence[MetricContract], edges: Sequence[RelationshipContract],
                      dimensions_by_table: Mapping[int, Sequence[str]] | None = None) -> tuple[QuestionEvidence, ...]:
    patterns = {"profitability": r"利润|毛利|profit|margin", "cash": r"现金|回款|到账|应收|cash|receivable",
                "inventory": r"库存|stock|inventory", "customer": r"满意度|评价|退款|评分|rating|refund",
                "workforce": r"绩效|工资|薪资|工时|加班|salary|hours", "procurement": r"采购|入库|purchase",
                "channel": r"销售|收入|费用|sales|revenue|expense"}
    questions = []
    for topic in topics:
        supported = [c for c in contracts if c.status != "unavailable"]
        if topic in patterns:
            supported = [c for c in supported if re.search(patterns[topic], c.field, re.I)]
        elif topic == "trend":
            supported = [c for c in supported if c.date_field]
        dimensions = dimensions_by_table or {}
        if topic == "ranking":
            supported = [c for c in supported if dimensions.get(c.table_index) and c.aggregation == "sum"]
        for requirement, pattern in {"channel": r"渠道|平台|门店|地区|区域|channel|platform|store|region",
                                     "customer": r"客户|顾客|会员|customer|buyer",
                                     "workforce": r"员工|人员|部门|employee|staff|department"}.items():
            if topic == requirement:
                supported = [c for c in supported if any(re.search(pattern, d, re.I) for d in dimensions.get(c.table_index, ()))]
        missing = []
        status = "available" if supported else "unavailable"
        boundary = "仅输出已绑定来源的事实统计；推断规则仍需核验"
        if topic == "profitability":
            status = "partial" if supported else "unavailable"
            boundary = "毛利/贡献不等于净利润；无法据零散收入、采购或费用跨表推算企业净利润"
            missing.append("需确认销售成本、期间费用、税费、退款、期间及有效订单口径")
        elif topic == "inventory":
            status = "partial" if supported else "unavailable"
            missing.append("补货/积压判断需安全库存、目标库存天数及一致的数量单位")
        elif topic == "relationships":
            status = "partial" if edges else "unavailable"
            missing.append("候选关联须确认业务键和事实粒度；当前不自动跨事实表分摊")
        elif topic in {"quality", "anomaly", "overview"}:
            status = "available"
            boundary = "结构审计可执行；数值结论须另外通过逐指标契约校验"
        if not supported and topic not in {"quality", "relationships", "overview", "anomaly"}:
            missing.append("缺少已通过口径校验的指标或所需维度/时间")
        evidence = tuple(f"{c.table_name}.{c.field}（{c.time_basis}）" for c in supported[:20])
        sources = tuple(dict.fromkeys(c.table_name for c in supported))
        if topic == "relationships":
            sources = tuple(dict.fromkeys(t for e in edges for t in (e.left_table, e.right_table)))
            evidence = tuple(f"{e.left_table}.{e.left_key} ↔ {e.right_table}.{e.right_key}；{e.cardinality}；{e.reason}" for e in edges[:20])
        questions.append(QuestionEvidence(topic, status, sources, evidence, tuple(missing), boundary))
    return tuple(questions)


def contracts_frame(contracts: Sequence[MetricContract]) -> pd.DataFrame:
    return pd.DataFrame([{"来源事实表": c.table_name, "指标": c.field, "聚合方式": c.aggregation,
        "单位": c.unit, "候选物理键": "+".join(c.grain) or "未验证", "时间字段": c.date_field,
        "时间口径": c.time_basis, "分子": c.numerator, "分母": c.denominator,
        "记录范围": c.population, "有效数值记录": c.valid_rows, "源记录数": c.total_rows,
        "状态": c.status, "需确认事项": "；".join(c.issues), "规则来源": c.origin,
        "契约版本": c.version} for c in contracts])


def questions_frame(questions: Sequence[QuestionEvidence]) -> pd.DataFrame:
    return pd.DataFrame([{"分析问题": q.topic, "证据状态": q.status, "数据源": "、".join(q.sources),
                         "证据": "；".join(q.evidence), "缺少条件": "；".join(q.missing),
                         "结论边界": q.conclusion_boundary} for q in questions])


def model_dict(contracts, edges, questions):
    return {"schema_version": 1, "metrics": [asdict(c) for c in contracts],
            "relationships": [asdict(e) for e in edges], "questions": [asdict(q) for q in questions]}
