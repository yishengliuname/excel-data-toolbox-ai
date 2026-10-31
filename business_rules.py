"""Versioned, explicit business policy. No model-generated code or silent repairs.

The caller must obtain human approval before applying a policy. HTTP callers
are additionally guarded by task-bound approval tickets in the server.
"""
from __future__ import annotations

from dataclasses import dataclass
import json
import math
import re
from typing import Any, Mapping, Sequence

import pandas as pd

from .business_contracts import MONEY_UNITS, numeric_values, _money_unit, metric_contract
from .metric_semantics import classify_metric, normalise


class BusinessRuleError(ValueError):
    """A policy is invalid or incompatible; never downgrade this to success."""


@dataclass(frozen=True)
class PreparedBusinessData:
    frames: tuple[pd.DataFrame, ...]
    contexts: Mapping[str, Mapping[str, Any]]
    audit: pd.DataFrame
    excluded: pd.DataFrame
    originals: pd.DataFrame


def _keys(value, allowed, label):
    if not isinstance(value, Mapping) or set(value) - set(allowed):
        raise BusinessRuleError(f"{label} 结构无效或包含不支持的字段")


def _text(value, label, *, empty=False):
    if not isinstance(value, str) or (not empty and not value.strip()) or len(value) > 160:
        raise BusinessRuleError(f"{label} 必须是1至160字符的文本")
    return value.strip()


def _strings(value, label, *, empty=True):
    if not isinstance(value, (list, tuple)) or len(value) > 100 or (not empty and not value):
        raise BusinessRuleError(f"{label} 必须是长度不超过100的文本数组")
    result = [_text(v, label) for v in value]
    if len(result) != len(set(result)):
        raise BusinessRuleError(f"{label} 不能重复")
    return result


def _insert_marker(frame, label, value):
    while label in frame:
        label = "_" + label
    frame.insert(0, label, value)


def validate_business_rules(raw: Mapping[str, Any]) -> dict[str, Any]:
    """Reject unknown fields, code, malformed definitions and unsafe overrides."""
    _keys(raw, {"schema_version", "tables", "joins", "requirements"}, "业务规则")
    if type(raw.get("schema_version")) is not int or raw["schema_version"] != 1:
        raise BusinessRuleError("业务规则 schema_version 必须为1")
    tables = raw.get("tables")
    if not isinstance(tables, (list, tuple)) or not 1 <= len(tables) <= 100:
        raise BusinessRuleError("业务规则必须包含1至100张表")
    names = set()
    for table in tables:
        _keys(table, {"source_name", "rename", "date_field", "statuses", "duplicate_key", "metrics",
                      "invalid_numeric", "role", "derived", "value_maps"}, "表规则")
        name = _text(table.get("source_name"), "source_name")
        if name in names:
            raise BusinessRuleError("source_name 不能重复")
        names.add(name)
        if table.get("role", "fact") not in {"fact", "dimension", "notes", "summary"}:
            raise BusinessRuleError("表角色无效")
        if "date_field" in table:
            _text(table["date_field"], "date_field", empty=True)
        rename = table.get("rename", {})
        _keys(rename, set(rename) if isinstance(rename, Mapping) else set(), "字段映射")
        for old, new in rename.items():
            _text(old, "原字段")
            _text(new, "目标字段")
        if len(rename) > 100 or len(set(rename.values())) != len(rename):
            raise BusinessRuleError("字段映射过多或目标字段重复")
        value_maps = table.get("value_maps", {})
        _keys(value_maps, set(value_maps) if isinstance(value_maps, Mapping) else set(), "值映射")
        if len(value_maps) > 100:
            raise BusinessRuleError("值映射字段过多")
        for field, mapping in value_maps.items():
            _text(field, "值映射字段")
            _keys(mapping, set(mapping) if isinstance(mapping, Mapping) else set(), "标准值映射")
            if len(mapping) > 100:
                raise BusinessRuleError("单字段值映射超过100项")
            for old, new in mapping.items():
                _text(old, "原值")
                _text(new, "标准值")
        _strings(table.get("duplicate_key", []), "去重键")
        if table.get("invalid_numeric", "block") not in {"block", "exclude"}:
            raise BusinessRuleError("invalid_numeric 仅支持 block/exclude")
        statuses = table.get("statuses", {})
        _keys(statuses, set(statuses) if isinstance(statuses, Mapping) else set(), "状态规则")
        if len(statuses) > 30:
            raise BusinessRuleError("状态规则过多")
        for field, rule in statuses.items():
            _text(field, "状态字段")
            _keys(rule, {"mode", "valid", "invalid"}, "状态规则")
            if rule.get("mode") not in {"filter", "ignore"}:
                raise BusinessRuleError("状态规则必须明确 filter 或 ignore")
            valid = _strings(rule.get("valid", []), "有效状态", empty=rule["mode"] == "ignore")
            invalid = _strings(rule.get("invalid", []), "排除状态")
            if set(valid) & set(invalid):
                raise BusinessRuleError("有效与排除状态不能重叠")
            if rule["mode"] == "ignore" and (valid or invalid):
                raise BusinessRuleError("不作为经营过滤条件的状态不得同时配置过滤值")
        metrics = table.get("metrics", {})
        _keys(metrics, set(metrics) if isinstance(metrics, Mapping) else set(), "指标规则")
        if len(metrics) > 100:
            raise BusinessRuleError("指标规则过多")
        for field, rule in metrics.items():
            _text(field, "指标字段")
            _keys(rule, {"aggregation", "unit", "numerator", "denominator"}, "指标规则")
            if rule.get("aggregation") not in {"sum", "mean", "last", "weighted_ratio"}:
                raise BusinessRuleError("指标聚合方式无效")
            _text(rule.get("unit"), "指标单位")
            if re.search(r"待确认|未声明|unknown", rule["unit"], re.I):
                raise BusinessRuleError("确认指标单位不能仍标记待确认/未声明")
            if rule["aggregation"] == "weighted_ratio":
                if rule["unit"] != "%":
                    raise BusinessRuleError("比例单位必须为%")
                _text(rule.get("numerator"), "分子")
                _text(rule.get("denominator"), "分母")
                if rule["numerator"] == rule["denominator"]:
                    raise BusinessRuleError("分子分母不能相同")
            elif "numerator" in rule or "denominator" in rule:
                raise BusinessRuleError("非比例指标不允许配置分子分母")
        derived = table.get("derived", [])
        if not isinstance(derived, (list, tuple)) or len(derived) > 30:
            raise BusinessRuleError("派生规则必须是最多30项的数组")
        for rule in derived:
            _keys(rule, {"field", "operator", "left", "right", "unit"}, "派生规则")
            for key in ("field", "left", "right", "unit"):
                _text(rule.get(key), key)
            if rule.get("operator") not in {"add", "subtract"}:
                raise BusinessRuleError("首版派生规则仅支持同表同单位加减，不执行任意公式")
    joins = raw.get("joins", [])
    if not isinstance(joins, (list, tuple)) or len(joins) > 30:
        raise BusinessRuleError("关联规则必须是最多30项的数组")
    for rule in joins:
        _keys(rule, {"left", "right", "left_key", "right_key", "fields"}, "关联规则")
        for key in ("left", "right", "left_key", "right_key"):
            _text(rule.get(key), key)
        if rule["left"] == rule["right"] or rule["left"] not in names or rule["right"] not in names:
            raise BusinessRuleError("关联必须指定不同的已声明表")
        _strings(rule.get("fields"), "维度属性", empty=False)
    topics = _strings(raw.get("requirements", []), "需求主题")
    allowed = {"overview", "profitability", "trend", "ranking", "quality", "relationships", "anomaly",
               "inventory", "customer", "workforce", "channel", "cash", "procurement"}
    if set(topics) - allowed:
        raise BusinessRuleError("存在未支持的需求主题，不能静默忽略")
    # Convert immutable planner structures to plain JSON; no object deserialization.
    try:
        def plain(value):
            if isinstance(value, Mapping):
                return {k: plain(v) for k, v in value.items()}
            if isinstance(value, (tuple, list)):
                return [plain(v) for v in value]
            return value
        encoded = json.dumps(plain(raw), ensure_ascii=False, allow_nan=False)
    except (TypeError, ValueError) as exc:
        raise BusinessRuleError("业务规则必须为有限JSON值") from exc
    if len(encoded.encode("utf-8")) > 100_000:
        raise BusinessRuleError("业务规则超过100KB限制")
    return json.loads(encoded)


def prepare_business_data(frames: Sequence[pd.DataFrame], source_names: Sequence[str], raw) -> PreparedBusinessData:
    rules = validate_business_rules(raw)
    if len(frames) != len(source_names) or len(set(source_names)) != len(source_names):
        raise BusinessRuleError("源表范围无效")
    table_rules = {r["source_name"]: r for r in rules["tables"]}
    if set(table_rules) != set(source_names):
        raise BusinessRuleError("规则必须逐表覆盖当前输入，不能混入历史表或漏掉新表")
    prepared, contexts, audit, excluded, originals = [], {}, [], [], []
    for name, raw_frame in zip(source_names, frames):
        rule = table_rules[name]
        original = raw_frame.copy(deep=True)
        marker = "规则源行号"
        while marker in original:
            marker = "_" + marker
        original.insert(0, marker, range(2, len(original) + 2))
        _insert_marker(original, "规则来源表", name)
        originals.append(original)
        frame = raw_frame.copy(deep=True).reset_index(drop=True)
        frame.index = pd.RangeIndex(2, len(frame) + 2)  # stable source record positions
        before = len(frame)
        rename = rule.get("rename", {})
        if set(rename) - set(frame):
            raise BusinessRuleError(f"{name}：字段映射源字段不存在")
        frame = frame.rename(columns=rename)
        if not frame.columns.is_unique or len({normalise(c) for c in frame}) != len(frame.columns):
            raise BusinessRuleError(f"{name}：映射造成字段冲突")
        referenced = [rule.get("date_field", ""), *rule.get("duplicate_key", []),
                      *rule.get("statuses", {}), *rule.get("metrics", {}), *rule.get("value_maps", {})]
        if any(c and c not in frame for c in referenced):
            raise BusinessRuleError(f"{name}：规则引用了不存在的字段；请重新映射")
        for field, mapping in rule.get("value_maps", {}).items():
            if classify_metric(field).kind not in {"identifier", "numeric"} or field in rule.get("metrics", {}):
                raise BusinessRuleError(f"{name}：值映射只支持标识/分类文本，不能重写金额、日期或指标")
            if pd.api.types.is_numeric_dtype(frame[field]) and classify_metric(field).kind != "identifier":
                raise BusinessRuleError(f"{name}：不能把未知数值字段当成名称字典替换")
            # Exact, approved dictionary only. No fuzzy guessing or implicit
            # trim outside the declared mapping; nulls stay null.
            frame[field] = frame[field].map(lambda value: mapping.get(str(value), value) if pd.notna(value) else value)

        def reject(mask, reason):
            nonlocal frame
            rejected = frame.loc[mask].copy()
            if not rejected.empty:
                _insert_marker(rejected, "排除源行号", rejected.index)
                _insert_marker(rejected, "排除来源表", name)
                _insert_marker(rejected, "排除原因", reason)
                excluded.append(rejected)
                frame = frame.loc[~mask].copy()

        for field, state_rule in rule.get("statuses", {}).items():
            if state_rule["mode"] == "ignore":
                continue
            values = frame[field].astype("string").str.strip()
            known = [*state_rule["valid"], *state_rule.get("invalid", [])]
            if values.isna().any() or (~values.isin(known)).any():
                raise BusinessRuleError(f"{name}.{field}：发现未确认或缺失状态；须重新确认，不自动排除")
            reject(~values.isin(state_rule["valid"]), f"已确认状态规则：{field}")
        keys = rule.get("duplicate_key", [])
        if keys:
            if frame[keys].isna().any(axis=None) or frame[keys].astype("string").apply(lambda s: s.str.strip().eq("")).any(axis=None):
                raise BusinessRuleError(f"{name}：业务键缺失，不能安全去重")
            duplicates = frame.duplicated(keys, keep=False)
            if frame.loc[duplicates].drop_duplicates().duplicated(keys).any():
                raise BusinessRuleError(f"{name}：同业务键内容冲突，不能选择一条自动删除")
            reject(frame.duplicated(keys, keep="first"), "已确认业务键的完全重复记录")
        metrics = rule.get("metrics", {})
        metric_fields = [c for c in frame if c in metrics or classify_metric(c).kind in {"additive", "count", "balance", "ratio", "score"}]
        if rule.get("invalid_numeric", "block") == "exclude":
            invalid = pd.Series(False, index=frame.index)
            for field in metric_fields:
                ratio = metrics.get(field, {}).get("aggregation") == "weighted_ratio" or classify_metric(field).kind == "ratio"
                values = numeric_values(frame[field], allow_percent=ratio)
                invalid |= values.isna() | ~values.map(lambda v: math.isfinite(float(v)) if pd.notna(v) else False)
            reject(invalid, "已确认隔离策略：缺失/无效数值（不填零）")
        for derived in rule.get("derived", []):
            field, left, right = derived["field"], derived["left"], derived["right"]
            if field in frame or left not in frame or right not in frame:
                raise BusinessRuleError(f"{name}：派生字段已存在或操作数不存在")
            resolved = [_money_unit(frame, c, monetary=metrics.get(c, {}).get("unit") in MONEY_UNITS) for c in (left, right)]
            if any(problems for _, problems in resolved):
                raise BusinessRuleError(f"{name}：派生操作数币种声明冲突")
            if any(metrics.get(c) and unit and metrics[c]["unit"] != unit
                   for c, (unit, _) in zip((left, right), resolved)):
                raise BusinessRuleError(f"{name}：人工单位与源声明冲突")
            units = [metrics.get(c, {}).get("unit") or unit for c, (unit, _) in zip((left, right), resolved)]
            operand_context = {"metrics": metrics, "date_field": rule.get("date_field", ""), "status_confirmed": True}
            for operand in (left, right):
                operand_contract = metric_contract(0, name, frame, operand, context=operand_context)
                if any("币种" in issue or "单位" in issue for issue in operand_contract.issues):
                    raise BusinessRuleError(f"{name}：派生操作数单位或币种不一致")
            if not units[0] or units[0] != units[1] or units[0] != derived["unit"]:
                raise BusinessRuleError(f"{name}：派生计算要求同一单位，不执行隐式换算")
            operands = [numeric_values(frame[c], allow_percent=False) for c in (left, right)]
            if any(s.isna().any() or not s.map(lambda x: math.isfinite(float(x))).all() for s in operands):
                raise BusinessRuleError(f"{name}：派生计算存在无效操作数，不填零")
            frame[field] = operands[0] + operands[1] if derived["operator"] == "add" else operands[0] - operands[1]
            metrics = {**metrics, field: {"aggregation": "sum", "unit": derived["unit"]}}
        if frame.empty and rule.get("role", "fact") == "fact":
            raise BusinessRuleError(f"{name}：确认范围内没有有效记录；不能用原始总额代替")
        status_fields = [c for c in frame if re.search(r"状态|status", str(c), re.I)]
        contexts[name] = {"date_field": rule.get("date_field", ""), "metrics": metrics,
                          "status_confirmed": all(c in rule.get("statuses", {}) for c in status_fields),
                          "population": "人工确认业务规则后的记录范围", "grain": keys,
                          "role": rule.get("role", "fact")}
        prepared.append(frame.reset_index(drop=True))
        audit.append({"来源表": name, "源记录数": before, "纳入记录数": len(frame), "排除记录数": before - len(frame),
                      "确认日期字段": rule.get("date_field", ""), "确认业务键": " + ".join(keys),
                      "字段映射": json.dumps(rename, ensure_ascii=False), "规则版本": 1,
                      "标准值映射": json.dumps(rule.get("value_maps", {}), ensure_ascii=False),
                      "状态规则": json.dumps(rule.get("statuses", {}), ensure_ascii=False),
                      "指标规则": json.dumps(metrics, ensure_ascii=False),
                      "派生规则": json.dumps(rule.get("derived", []), ensure_ascii=False),
                      "无效数值策略": rule.get("invalid_numeric", "block"),
                      "处理边界": "原文件不变；排除明细保留；日期不隐式转换归属"})
    by_name = dict(zip(source_names, prepared))
    for join in rules.get("joins", []):
        left, right = join["left"], join["right"]
        if contexts[left]["role"] != "fact" or contexts[right]["role"] != "dimension":
            raise BusinessRuleError("首版规则关联仅支持确认的事实表 → 唯一维度表，不连接两个事实域后求和")
        lf, rf = by_name[left], by_name[right]
        lk, rk, attrs = join["left_key"], join["right_key"], join["fields"]
        if lk not in lf or any(c not in rf for c in [rk, *attrs]):
            raise BusinessRuleError("关联字段不存在")
        for field in attrs:
            if field in lf or field == rk or classify_metric(field).kind in {"additive", "count", "balance", "ratio", "score"}:
                raise BusinessRuleError("关联属性重复或包含度量，禁止未经分摊直接复制指标")
            if pd.api.types.is_numeric_dtype(rf[field]) and classify_metric(field).kind != "identifier":
                raise BusinessRuleError("数值维度属性须先明确含义；不自动复制为经营指标")
        lvalues, rvalues = lf[lk].astype("string").str.strip(), rf[rk].astype("string").str.strip()
        if lvalues.isna().any() or rvalues.isna().any() or lvalues.eq("").any() or rvalues.eq("").any() or not rvalues.is_unique:
            raise BusinessRuleError("关联键为空或维度键不唯一，禁止扩行")
        if not lvalues.isin(set(rvalues)).all():
            raise BusinessRuleError("事实键覆盖不完整；须核验，不静默丢掉未匹配记录")
        key = "__business_join_key__"
        while key in lf or key in rf:
            key = "_" + key
        joined = lf.assign(**{key: lvalues}).merge(rf.assign(**{key: rvalues})[[key, *attrs]],
            on=key, how="left", validate="many_to_one", sort=False).drop(columns=[key])
        if len(joined) != len(lf):
            raise BusinessRuleError("关联改变事实粒度")
        by_name[left] = joined
        contexts[left].setdefault("confirmed_joins", []).append(dict(join))
        audit.append({"来源表": left, "源记录数": len(lf), "纳入记录数": len(joined), "排除记录数": 0,
                      "处理边界": f"人工确认维度关联：{right}；many_to_one；行覆盖100%"})
    concat = lambda parts: pd.concat(parts, ignore_index=True, sort=False) if parts else pd.DataFrame()
    return PreparedBusinessData(tuple(by_name[n] for n in source_names), contexts, pd.DataFrame(audit),
                                concat(excluded), concat(originals))
