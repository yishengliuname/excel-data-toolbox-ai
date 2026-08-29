"""Conservative fuzzy matching helpers for tabular business data.

The functions in this module never mutate their input dataframes.  Fuzzy
matches are suggestions: ambiguous or low-confidence rows keep lookup payload
columns empty until a caller explicitly confirms a mapping.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from difflib import SequenceMatcher
from numbers import Real
import re
import unicodedata
from typing import Any

import pandas as pd


DEFAULT_COMPANY_SUFFIXES: tuple[str, ...] = (
    "集团股份有限公司",
    "集团有限责任公司",
    "有限责任公司",
    "股份有限公司",
    "集团有限公司",
    "集团公司",
    "有限公司",
    "公司",
    "集团",
    "co., ltd.",
    "co., ltd",
    "co. ltd.",
    "co. ltd",
    "co ltd",
    "company limited",
    "limited liability company",
    "corporation",
    "incorporated",
    "llc",
)

CLUSTER_COLUMNS: tuple[str, ...] = (
    "原值",
    "建议标准值",
    "相似度",
    "出现次数",
    "组ID",
)

_CANDIDATE_COLUMN = "候选值"
_SCORE_COLUMN = "相似度"
_SECOND_CANDIDATE_COLUMN = "次选候选值"
_SECOND_SCORE_COLUMN = "次选相似度"
_STATUS_COLUMN = "匹配状态"
_LOOKUP_METADATA_COLUMNS = (
    _CANDIDATE_COLUMN,
    _SCORE_COLUMN,
    _SECOND_CANDIDATE_COLUMN,
    _SECOND_SCORE_COLUMN,
    _STATUS_COLUMN,
)


def _require_dataframe(df: pd.DataFrame, name: str) -> None:
    if not isinstance(df, pd.DataFrame):
        raise TypeError(f"{name} 必须是 pandas DataFrame")
    if df.columns.duplicated().any():
        duplicate_labels = df.columns[df.columns.duplicated()]
        duplicates = list(dict.fromkeys(str(value) for value in duplicate_labels))
        raise ValueError(f"{name} 含重复列名，处理前请先重命名：{duplicates}")


def _require_column(df: pd.DataFrame, column: str, argument: str) -> None:
    if column not in df.columns:
        raise KeyError(f"{argument} 指定的列不存在：{column!r}")


def _validate_probability(value: float, argument: str) -> float:
    if isinstance(value, bool) or not isinstance(value, Real):
        raise TypeError(f"{argument} 必须是 0 到 1 之间的数字")
    result = float(value)
    if not 0.0 <= result <= 1.0:
        raise ValueError(f"{argument} 必须在 0 到 1 之间")
    return result


def _is_missing(value: Any) -> bool:
    if value is None or value is pd.NA:
        return True
    try:
        marker = pd.isna(value)
    except (TypeError, ValueError):
        return False
    if isinstance(marker, bool):
        return marker
    try:
        if getattr(marker, "ndim", 1) == 0:
            return bool(marker)
    except (TypeError, ValueError):
        return False
    return False


def _base_normalize(value: Any) -> str:
    if _is_missing(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).casefold().strip()
    return re.sub(r"\s+", "", text)


def _prepare_company_suffixes(
    company_suffixes: Sequence[str] | str | None,
) -> tuple[str, ...]:
    if company_suffixes is None:
        return ()
    raw_suffixes = (company_suffixes,) if isinstance(company_suffixes, str) else company_suffixes
    prepared: set[str] = set()
    for suffix in raw_suffixes:
        if not isinstance(suffix, str):
            raise TypeError("company_suffixes 中的每个后缀都必须是字符串")
        normalized = _base_normalize(suffix)
        if normalized:
            prepared.add(normalized)
    return tuple(sorted(prepared, key=lambda item: (-len(item), item)))


def _normalize_with_prepared_suffixes(value: Any, suffixes: Sequence[str]) -> str:
    text = _base_normalize(value)
    if not text:
        return ""
    while text:
        removed = False
        for suffix in suffixes:
            # Do not turn a value consisting only of a suffix into an empty key.
            if len(text) > len(suffix) and text.endswith(suffix):
                text = text[: -len(suffix)].rstrip(".,，。·-_/")
                removed = True
                break
        if not removed:
            break
    return text


def normalize_text(
    value: Any,
    company_suffixes: Sequence[str] | str | None = DEFAULT_COMPANY_SUFFIXES,
) -> str:
    """Normalize a scalar for comparison without changing the original value.

    Normalization applies Unicode NFKC conversion (including full-width to
    half-width forms), case folding, whitespace removal, and optional trailing
    company-suffix removal.  Pass an empty tuple or ``None`` to retain company
    suffixes, or pass a custom sequence for another business domain.  Missing
    values normalize to an empty string.
    """

    suffixes = _prepare_company_suffixes(company_suffixes)
    return _normalize_with_prepared_suffixes(value, suffixes)


def _similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    return SequenceMatcher(None, left, right, autojunk=False).ratio()


def _raw_identity(value: Any) -> tuple[type[Any], Any]:
    try:
        hash(value)
    except TypeError:
        return type(value), repr(value)
    return type(value), value


@dataclass
class _RawValue:
    value: Any
    normalized: str
    count: int
    first_position: int


@dataclass
class _NormalizedValue:
    normalized: str
    raw_values: list[_RawValue]
    count: int
    first_position: int


def _collect_unique_values(
    series: pd.Series,
    suffixes: Sequence[str],
) -> list[_RawValue]:
    by_raw_value: dict[tuple[type[Any], Any], _RawValue] = {}
    for position, value in enumerate(series.tolist()):
        normalized = _normalize_with_prepared_suffixes(value, suffixes)
        if not normalized:
            continue
        identity = _raw_identity(value)
        existing = by_raw_value.get(identity)
        if existing is None:
            by_raw_value[identity] = _RawValue(
                value=value,
                normalized=normalized,
                count=1,
                first_position=position,
            )
        else:
            existing.count += 1
    return sorted(by_raw_value.values(), key=lambda item: item.first_position)


def _combine_normalized_values(raw_values: Sequence[_RawValue]) -> list[_NormalizedValue]:
    combined: dict[str, _NormalizedValue] = {}
    for raw in raw_values:
        existing = combined.get(raw.normalized)
        if existing is None:
            combined[raw.normalized] = _NormalizedValue(
                normalized=raw.normalized,
                raw_values=[raw],
                count=raw.count,
                first_position=raw.first_position,
            )
        else:
            existing.raw_values.append(raw)
            existing.count += raw.count
            existing.first_position = min(existing.first_position, raw.first_position)
    return sorted(
        combined.values(),
        key=lambda item: (-item.count, item.first_position, item.normalized),
    )


def cluster_similar_values(
    df: pd.DataFrame,
    column: str,
    threshold: float = 0.85,
    max_unique: int = 1_000,
    *,
    company_suffixes: Sequence[str] | str | None = DEFAULT_COMPANY_SUFFIXES,
) -> pd.DataFrame:
    """Return reviewable groups of similar values from one column.

    Only groups containing at least two distinct original values are returned.
    The suggested standard is the most frequent original value, with first
    appearance as the deterministic tie-breaker.  A group is a review aid only;
    this function never rewrites ``df`` or applies its suggestions.
    """

    _require_dataframe(df, "df")
    _require_column(df, column, "column")
    threshold_value = _validate_probability(threshold, "threshold")
    if isinstance(max_unique, bool) or not isinstance(max_unique, int):
        raise TypeError("max_unique 必须是正整数")
    if max_unique <= 0:
        raise ValueError("max_unique 必须是正整数")

    suffixes = _prepare_company_suffixes(company_suffixes)
    raw_values = _collect_unique_values(df[column], suffixes)
    if len(raw_values) > max_unique:
        raise ValueError(
            f"列 {column!r} 有 {len(raw_values)} 个有效唯一值，超过 max_unique={max_unique}；"
            "请先筛选数据或提高限制后再生成候选组"
        )
    normalized_values = _combine_normalized_values(raw_values)

    rows: list[dict[str, Any]] = []
    remaining = list(normalized_values)
    group_number = 0
    while remaining:
        seed = remaining.pop(0)
        group = [seed]
        not_grouped: list[_NormalizedValue] = []
        for candidate in remaining:
            if _similarity(seed.normalized, candidate.normalized) >= threshold_value:
                group.append(candidate)
            else:
                not_grouped.append(candidate)
        remaining = not_grouped

        group_raw_values = [raw for item in group for raw in item.raw_values]
        if len(group_raw_values) < 2:
            continue

        suggestion = min(
            group_raw_values,
            key=lambda item: (-item.count, item.first_position),
        )
        group_number += 1
        group_id = f"G{group_number:03d}"
        ordered_raw_values = sorted(
            group_raw_values,
            key=lambda item: (
                item is not suggestion,
                -item.count,
                item.first_position,
            ),
        )
        for raw in ordered_raw_values:
            rows.append(
                {
                    "原值": raw.value,
                    "建议标准值": suggestion.value,
                    "相似度": round(
                        _similarity(raw.normalized, suggestion.normalized), 4
                    ),
                    "出现次数": raw.count,
                    "组ID": group_id,
                }
            )

    return pd.DataFrame(rows, columns=CLUSTER_COLUMNS)


def _normalise_value_columns(value_columns: str | Sequence[str]) -> list[str]:
    values = [value_columns] if isinstance(value_columns, str) else list(value_columns)
    if len(values) != len(set(values)):
        raise ValueError("value_columns 不能包含重复列名")
    return values


def _scalar_equal(left: Any, right: Any) -> bool:
    if _is_missing(left) and _is_missing(right):
        return True
    if _is_missing(left) or _is_missing(right):
        return False
    try:
        result = left == right
    except (TypeError, ValueError):
        return False
    if isinstance(result, bool):
        return result
    try:
        if getattr(result, "ndim", 1) == 0:
            return bool(result)
    except (TypeError, ValueError):
        return False
    return False


@dataclass
class _LookupCandidate:
    normalized: str
    key_value: Any
    first_position: int
    payload: tuple[Any, ...]
    conflicting_payload: bool


def _prepare_lookup_candidates(
    lookup: pd.DataFrame,
    lookup_key: str,
    value_columns: Sequence[str],
    suffixes: Sequence[str],
) -> list[_LookupCandidate]:
    positions_by_normalized: dict[str, list[int]] = {}
    for position, value in enumerate(lookup[lookup_key].tolist()):
        normalized = _normalize_with_prepared_suffixes(value, suffixes)
        if normalized:
            positions_by_normalized.setdefault(normalized, []).append(position)

    candidates: list[_LookupCandidate] = []
    for normalized, positions in positions_by_normalized.items():
        first_position = positions[0]
        first_payload = tuple(
            lookup.iloc[first_position][column] for column in value_columns
        )
        conflicting = False
        for position in positions[1:]:
            payload = tuple(lookup.iloc[position][column] for column in value_columns)
            if any(
                not _scalar_equal(current, first)
                for current, first in zip(payload, first_payload)
            ):
                conflicting = True
                break
        candidates.append(
            _LookupCandidate(
                normalized=normalized,
                key_value=lookup.iloc[first_position][lookup_key],
                first_position=first_position,
                payload=first_payload,
                conflicting_payload=conflicting,
            )
        )
    return sorted(candidates, key=lambda item: item.first_position)


@dataclass(frozen=True)
class _MatchDecision:
    candidate_value: Any
    score: float | None
    second_candidate_value: Any
    second_score: float | None
    status: str
    payload: tuple[Any, ...] | None


def _decide_match(
    source_normalized: str,
    candidates: Sequence[_LookupCandidate],
    threshold: float,
    ambiguous_gap: float,
) -> _MatchDecision:
    if not source_normalized or not candidates:
        return _MatchDecision(pd.NA, None, pd.NA, None, "未匹配", None)

    ranked = sorted(
        (
            (_similarity(source_normalized, candidate.normalized), candidate)
            for candidate in candidates
        ),
        key=lambda item: (-item[0], item[1].first_position),
    )
    top_score, top = ranked[0]
    if len(ranked) > 1:
        second_score, second = ranked[1]
        second_value: Any = second.key_value
    else:
        second_score, second_value = None, pd.NA

    if top_score < threshold:
        status = "未匹配"
    elif top.conflicting_payload:
        status = "待确认"
    elif second_score is not None and top_score - second_score <= ambiguous_gap:
        status = "待确认"
    else:
        status = "已匹配"
    payload = top.payload if status == "已匹配" else None
    return _MatchDecision(
        candidate_value=top.key_value,
        score=round(top_score, 4),
        second_candidate_value=second_value,
        second_score=None if second_score is None else round(second_score, 4),
        status=status,
        payload=payload,
    )


def _unique_lookup_output_names(
    source_columns: Sequence[Any],
    value_columns: Sequence[str],
) -> list[str]:
    occupied = {str(column) for column in source_columns} | set(_LOOKUP_METADATA_COLUMNS)
    names: list[str] = []
    for column in value_columns:
        candidate = column
        if candidate in occupied:
            candidate = f"{column}_查找"
        suffix = 2
        while candidate in occupied:
            candidate = f"{column}_查找{suffix}"
            suffix += 1
        occupied.add(candidate)
        names.append(candidate)
    return names


def fuzzy_lookup(
    source: pd.DataFrame,
    lookup: pd.DataFrame,
    source_key: str,
    lookup_key: str,
    value_columns: str | Sequence[str],
    threshold: float = 0.85,
    ambiguous_gap: float = 0.03,
    *,
    company_suffixes: Sequence[str] | str | None = DEFAULT_COMPANY_SUFFIXES,
) -> pd.DataFrame:
    """Fuzzy-match one source key against a lookup table conservatively.

    The returned dataframe preserves all source rows and adds requested lookup
    values plus ``候选值``, ``相似度``, ``次选候选值``, ``次选相似度`` and
    ``匹配状态``.  Status is ``未匹配`` below ``threshold`` and ``待确认``
    when the best two candidates are within ``ambiguous_gap`` or when identical
    normalized lookup keys carry conflicting payloads.  Lookup payload columns
    are populated only for ``已匹配`` rows.
    """

    _require_dataframe(source, "source")
    _require_dataframe(lookup, "lookup")
    _require_column(source, source_key, "source_key")
    _require_column(lookup, lookup_key, "lookup_key")
    values = _normalise_value_columns(value_columns)
    for column in values:
        _require_column(lookup, column, "value_columns")
    threshold_value = _validate_probability(threshold, "threshold")
    gap_value = _validate_probability(ambiguous_gap, "ambiguous_gap")

    metadata_collisions = [
        column for column in _LOOKUP_METADATA_COLUMNS if column in source.columns
    ]
    if metadata_collisions:
        raise ValueError(f"source 已包含模糊匹配结果列：{metadata_collisions}")

    suffixes = _prepare_company_suffixes(company_suffixes)
    candidates = _prepare_lookup_candidates(lookup, lookup_key, values, suffixes)
    output_value_names = _unique_lookup_output_names(source.columns, values)
    payload_columns: dict[str, list[Any]] = {
        output_name: [] for output_name in output_value_names
    }
    candidate_values: list[Any] = []
    scores: list[float | None] = []
    second_candidate_values: list[Any] = []
    second_scores: list[float | None] = []
    statuses: list[str] = []
    decision_cache: dict[str, _MatchDecision] = {}

    for value in source[source_key].tolist():
        normalized = _normalize_with_prepared_suffixes(value, suffixes)
        decision = decision_cache.get(normalized)
        if decision is None:
            decision = _decide_match(
                normalized,
                candidates,
                threshold_value,
                gap_value,
            )
            decision_cache[normalized] = decision
        candidate_values.append(decision.candidate_value)
        scores.append(decision.score)
        second_candidate_values.append(decision.second_candidate_value)
        second_scores.append(decision.second_score)
        statuses.append(decision.status)
        for position, output_name in enumerate(output_value_names):
            payload_columns[output_name].append(
                pd.NA if decision.payload is None else decision.payload[position]
            )

    result = source.copy(deep=True)
    for output_name in output_value_names:
        result[output_name] = pd.array(payload_columns[output_name], dtype="object")
    result[_CANDIDATE_COLUMN] = pd.array(candidate_values, dtype="object")
    result[_SCORE_COLUMN] = pd.array(scores, dtype="Float64")
    result[_SECOND_CANDIDATE_COLUMN] = pd.array(
        second_candidate_values, dtype="object"
    )
    result[_SECOND_SCORE_COLUMN] = pd.array(second_scores, dtype="Float64")
    result[_STATUS_COLUMN] = statuses
    return result


def apply_value_mapping(
    df: pd.DataFrame,
    column: str,
    mapping: Mapping[Any, Any],
    output_column: str | None = None,
) -> pd.DataFrame:
    """Apply an explicitly confirmed exact-value mapping to a copied dataframe.

    Unmapped and missing values are preserved.  This function performs no fuzzy
    inference; callers should pass only mappings a user has approved.  Set
    ``output_column`` to keep the original column and write standards elsewhere.
    """

    _require_dataframe(df, "df")
    _require_column(df, column, "column")
    if not isinstance(mapping, Mapping):
        raise TypeError("mapping 必须是键值映射")
    target_column = column if output_column is None else output_column
    if not isinstance(target_column, str) or not target_column:
        raise ValueError("output_column 必须是非空字符串")
    if target_column != column and target_column in df.columns:
        raise ValueError(f"output_column 已存在：{target_column!r}")

    def replace_if_confirmed(value: Any) -> Any:
        if _is_missing(value):
            return value
        try:
            return mapping[value] if value in mapping else value
        except (KeyError, TypeError, ValueError):
            return value

    result = df.copy(deep=True)
    result[target_column] = result[column].map(replace_if_confirmed)
    return result


__all__ = [
    "CLUSTER_COLUMNS",
    "DEFAULT_COMPANY_SUFFIXES",
    "apply_value_mapping",
    "cluster_similar_values",
    "fuzzy_lookup",
    "normalize_text",
]
