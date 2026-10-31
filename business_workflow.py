"""Local review, explicit policy reuse and evidence-gated execution boundaries.

Profiles store schemas and confirmed rules, not customer rows. Schema matching
is only a proposal: review-only profiles require fresh task-bound confirmation.
Explicit opt-in, expiring local-report grants live in trusted_automation.py.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import re
import uuid

from .analysis_compiler import compile_analysis
from .business_contracts import _money_unit, numeric_values
from .business_rules import BusinessRuleError, validate_business_rules
from .metric_semantics import classify_metric, classify_sheet_role


def review_business_data(frames, names, request):
    """Suggestions are deliberately NOT permissions or approved policies."""
    plan = compile_analysis(frames, source_names=names, user_request=request)
    tables, questions = [], []
    for name, frame in zip(names, frames):
        metrics, statuses, unknown = {}, {}, []
        for field in frame:
            semantic = classify_metric(field)
            if re.search(r"状态|status", field, re.I):
                values = frame[field].dropna().astype("string").str.strip().unique().tolist()
                if len(values) > 100:
                    raise BusinessRuleError("状态值超过100种，请先规范状态字典")
                questions.append(f"{name}.{field}：逐个确认有效/排除状态，或明确不作为经营过滤条件")
            efficiency = bool(re.search(r"单价|平均|周转天数|库存天数|每单|每人|unit.?price|average|per.?unit", field, re.I))
            if semantic.aggregation in {"sum", "mean", "last"} and not efficiency:
                unit, _ = _money_unit(frame, field)
                metrics[field] = {"aggregation": semantic.aggregation, "unit": unit or "单位待确认"}
            elif (semantic.aggregation == "unknown" or efficiency) and numeric_values(frame[field]).notna().any():
                unknown.append(field)
        dates = [str(c) for c in frame if classify_metric(c).kind == "date"]
        if len(dates) > 1:
            questions.append(f"{name}：选择一个分析日期字段；其他日期不混入月度趋势")
        if frame.duplicated().any():
            questions.append(f"{name}：存在完全重复记录，确认业务键后才能去重；冲突记录不自动删除")
        if unknown:
            questions.append(f"{name}：未知数值字段须声明聚合和单位，否则保持不可计算")
        tables.append({"source_name": name, "role": classify_sheet_role(name, frame),
                       "rename": {}, "date_field": dates[0] if len(dates) == 1 else "",
                       "statuses": statuses, "duplicate_key": [], "metrics": metrics,
                       "invalid_numeric": "block", "derived": []})
    return {"draft_rules": {"schema_version": 1, "tables": tables, "joins": [],
                            "requirements": list(plan.intent_topics)},
            "questions": questions,
            "tables": [{"source_name": name, "columns": list(frame.columns),
                        "date_candidates": [c for c in frame if classify_metric(c).kind == "date"],
                        "status_values": {c: frame[c].dropna().astype("string").str.strip().unique().tolist()
                                          for c in frame if re.search(r"状态|status", c, re.I)},
                        "numeric_candidates": [c for c in frame if numeric_values(frame[c]).notna().any()
                                               and classify_metric(c).kind not in {"identifier", "date"}],
                        "rows": len(frame)} for name, frame in zip(names, frames)],
            "contract_issues": [{"table": c.table_name, "field": c.field, "issues": list(c.issues)}
                                for c in plan.metric_contracts if c.status == "unavailable"],
            "coverage": [{"topic": q.topic, "status": q.status, "missing": list(q.missing),
                          "boundary": q.conclusion_boundary} for q in plan.question_evidence],
            "boundary": "只建议规则，不代表批准。事实域分别计算；不自动分摊退款/费用或推算企业净利润。"}


class BusinessProfileStore:
    """Customer-scoped local JSON. No executable deserialization or scanning."""

    def __init__(self, root):
        self.root = Path(root)

    def _scope(self, customer):
        if not isinstance(customer, str) or not customer.strip() or len(customer) > 100:
            raise BusinessRuleError("请明确客户标识（1至100字符）；不跨客户自动复用")
        return self.root / hashlib.sha256(customer.strip().encode("utf-8")).hexdigest()

    def save(self, customer, title, rules, frames, names):
        rules = validate_business_rules(rules)
        if not isinstance(title, str) or not title.strip() or len(title) > 100:
            raise BusinessRuleError("配置名称必须为1至100字符")
        if len(frames) != len(names) or {t["source_name"] for t in rules["tables"]} != set(names):
            raise BusinessRuleError("配置数据源范围不一致")
        folder = self._scope(customer)
        folder.mkdir(parents=True, exist_ok=True)
        profile_id = uuid.uuid4().hex
        record = {"schema_version": 1, "profile_id": profile_id, "title": title.strip(),
                  "rules": rules, "schemas": {name: list(frame.columns) for name, frame in zip(names, frames)}}
        temporary = folder / f".{profile_id}.tmp"
        try:
            temporary.write_text(json.dumps(record, ensure_ascii=False, allow_nan=False), encoding="utf-8")
            os.replace(temporary, folder / f"{profile_id}.json")
        finally:
            temporary.unlink(missing_ok=True)
        return {"profile_id": profile_id, "title": title.strip(), "schema_version": 1}

    def list(self, customer):
        folder = self._scope(customer)
        if not folder.exists():
            return []
        result = []
        for path in sorted(folder.glob("*.json"))[:200]:
            record = json.loads(path.read_text(encoding="utf-8"))
            result.append({k: record[k] for k in ("profile_id", "title", "schema_version")})
        return result

    def load_for_review(self, customer, profile_id, frames, names):
        if not isinstance(profile_id, str) or not re.fullmatch(r"[a-f0-9]{32}", profile_id):
            raise BusinessRuleError("配置标识无效")
        path = self._scope(customer) / f"{profile_id}.json"
        if not path.is_file():
            raise BusinessRuleError("当前客户没有此配置")
        record = json.loads(path.read_text(encoding="utf-8"))
        rules = validate_business_rules(record["rules"])
        current = {n: list(f.columns) for n, f in zip(names, frames)}
        if len(current) != len(names) or len(current) != len(record["schemas"]):
            raise BusinessRuleError("表数量变化，不能直接复用配置")
        mapping = {}
        for old, columns in record["schemas"].items():
            if old in current and current[old] == columns:
                candidates = [old]
            else:
                candidates = [n for n, cols in current.items() if cols == columns and n not in mapping.values()]
            if len(candidates) != 1 or candidates[0] in mapping.values():
                raise BusinessRuleError("字段漂移或同构表匹配有歧义，须重新确认映射")
            mapping[old] = candidates[0]
        for table in rules["tables"]:
            table["source_name"] = mapping[table["source_name"]]
        for join in rules.get("joins", []):
            join["left"], join["right"] = mapping[join["left"]], mapping[join["right"]]
        return rules  # Not validated against values until human confirmation.
