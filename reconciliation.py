"""Conservative, explainable reconciliation for spreadsheet tables.

The module deliberately prefers a review candidate over a false automatic
match.  Amounts are parsed and compared with :class:`decimal.Decimal`, source
row positions are preserved, duplicate primary keys are quarantined, and
limited one-to-two/two-to-one split suggestions are never auto-confirmed.
"""

from __future__ import annotations

from bisect import bisect_left, bisect_right
from collections import Counter, defaultdict
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from itertools import combinations
import math
from types import MappingProxyType
from typing import Any
import unicodedata

import numpy as np
import pandas as pd


_PAIR_METADATA_COLUMNS = [
    "match_type",
    "candidate_group_id",
    "left_row_position",
    "right_row_position",
    "candidate_rank",
    "candidate_count",
    "match_score",
    "match_reason",
    "left_amount_decimal",
    "right_amount_decimal",
    "amount_difference_decimal",
    "absolute_amount_difference",
    "left_date",
    "right_date",
    "date_difference_days",
    "secondary_compared",
    "secondary_matched",
    "group_left_amount_decimal",
    "group_right_amount_decimal",
    "group_amount_difference_decimal",
]

_DUPLICATE_METADATA_COLUMNS = [
    "source_side",
    "source_row_position",
    "key_values",
    "duplicate_count_left",
    "duplicate_count_right",
    "duplicate_reason",
]


def _is_missing(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    try:
        missing = pd.isna(value)
        return bool(missing) if isinstance(missing, (bool, np.bool_)) else False
    except (TypeError, ValueError):
        return False


def _json_value(value: Any) -> Any:
    if _is_missing(value):
        return None
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    if isinstance(value, (int, np.integer)):
        numeric = int(value)
        return str(numeric) if abs(numeric) > 9_007_199_254_740_991 else numeric
    if isinstance(value, (float, np.floating)):
        numeric = float(value)
        return numeric if math.isfinite(numeric) else None
    if isinstance(value, str):
        return value
    if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
        return pd.Timestamp(value).isoformat()
    if isinstance(value, Mapping):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_value(item) for item in value]
    if hasattr(value, "item"):
        try:
            return _json_value(value.item())
        except (TypeError, ValueError):
            pass
    return str(value)


def _records(frame: pd.DataFrame) -> list[dict[str, Any]]:
    return [
        {str(column): _json_value(value) for column, value in row.items()}
        for row in frame.to_dict("records")
    ]


def _frozen_mapping(values: Mapping[str, Any]) -> Mapping[str, Any]:
    return MappingProxyType(dict(values))


@dataclass(frozen=True)
class ReconciliationResult:
    """Categorised reconciliation outputs and transparent run metadata."""

    matched: pd.DataFrame
    amount_difference: pd.DataFrame
    date_difference: pd.DataFrame
    review: pd.DataFrame
    left_only: pd.DataFrame
    right_only: pd.DataFrame
    duplicates: pd.DataFrame
    summary: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "summary", _frozen_mapping(self.summary))

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": _json_value(dict(self.summary)),
            "matched": _records(self.matched),
            "amount_difference": _records(self.amount_difference),
            "date_difference": _records(self.date_difference),
            "review": _records(self.review),
            "left_only": _records(self.left_only),
            "right_only": _records(self.right_only),
            "duplicates": _records(self.duplicates),
        }


@dataclass(frozen=True)
class _PreparedRow:
    position: int
    amount: Decimal | None
    date_value: date | None
    key: tuple[str, ...] | None


@dataclass(frozen=True)
class _Candidate:
    left_position: int
    right_position: int
    amount_difference: Decimal
    date_difference_days: int | None
    secondary_compared: int
    secondary_matched: int
    score: float
    key_conflict: bool


def _require_dataframe(frame: pd.DataFrame, name: str) -> None:
    if not isinstance(frame, pd.DataFrame):
        raise TypeError(f"{name} 必须是 pandas DataFrame")
    if frame.columns.duplicated().any():
        labels = frame.columns[frame.columns.duplicated()]
        duplicates = list(dict.fromkeys(str(value) for value in labels))
        raise ValueError(f"{name} 含重复列名，请先重命名：{duplicates}")


def _column_list(
    value: str | Sequence[str] | None,
    *,
    argument: str,
) -> list[str]:
    if value is None:
        return []
    result = [value] if isinstance(value, str) else list(value)
    if not result:
        return []
    if any(not isinstance(column, str) or not column for column in result):
        raise TypeError(f"{argument} 必须是非空字段名或字段名序列")
    if len(set(result)) != len(result):
        raise ValueError(f"{argument} 不能包含重复字段")
    return result


def _paired_columns(
    left_value: str | Sequence[str] | None,
    right_value: str | Sequence[str] | None,
    *,
    left_argument: str,
    right_argument: str,
) -> tuple[list[str], list[str]]:
    left_columns = _column_list(left_value, argument=left_argument)
    right_columns = _column_list(right_value, argument=right_argument)
    if bool(left_columns) != bool(right_columns):
        raise ValueError(f"{left_argument} 与 {right_argument} 必须同时提供")
    if len(left_columns) != len(right_columns):
        raise ValueError(f"{left_argument} 与 {right_argument} 字段数量必须相同")
    return left_columns, right_columns


def _validate_columns(
    frame: pd.DataFrame,
    columns: Sequence[str],
    *,
    argument: str,
) -> None:
    missing = [column for column in columns if column not in frame.columns]
    if missing:
        raise KeyError(f"{argument} 包含不存在的列：{missing}")


def _positive_integer(value: Any, argument: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise TypeError(f"{argument} 必须是正整数")
    if value <= 0:
        raise ValueError(f"{argument} 必须是正整数")
    return value


_CURRENCY_MARKS = ("人民币", "RMB", "CNY", "¥", "￥", "$", "€", "£")


def _decimal_value(value: Any) -> Decimal | None:
    """Parse a monetary scalar without binary floating-point arithmetic."""

    if _is_missing(value) or isinstance(value, (bool, np.bool_)):
        return None
    if isinstance(value, Decimal):
        return value if value.is_finite() else None
    if isinstance(value, (int, np.integer)):
        return Decimal(int(value))
    if isinstance(value, (float, np.floating)):
        numeric = float(value)
        return Decimal(str(numeric)) if math.isfinite(numeric) else None

    text = unicodedata.normalize("NFKC", str(value)).strip()
    negative_parentheses = text.startswith("(") and text.endswith(")")
    if negative_parentheses:
        text = text[1:-1].strip()
    for mark in _CURRENCY_MARKS:
        text = text.replace(mark, "").replace(mark.lower(), "")
    text = text.replace(",", "").replace("，", "").replace(" ", "")
    if negative_parentheses:
        text = f"-{text}"
    try:
        parsed = Decimal(text)
    except (InvalidOperation, ValueError):
        return None
    return parsed if parsed.is_finite() else None


def _amount_tolerance(value: Decimal | int | float | str) -> Decimal:
    parsed = _decimal_value(value)
    if parsed is None:
        raise ValueError("amount_tolerance 必须是有效非负金额")
    if parsed < 0:
        raise ValueError("amount_tolerance 必须是有效非负金额")
    return parsed


def _date_value(value: Any) -> date | None:
    if _is_missing(value):
        return None
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(
        value, (datetime, date, pd.Timestamp, np.datetime64)
    ):
        # Bare numbers are ambiguous (Unix units versus Excel serial dates).
        return None
    try:
        if isinstance(value, (datetime, date, pd.Timestamp, np.datetime64)):
            parsed = pd.Timestamp(value)
        else:
            try:
                parsed = pd.to_datetime(value, errors="coerce", format="mixed")
            except (TypeError, ValueError):
                parsed = pd.to_datetime(value, errors="coerce")
    except (TypeError, ValueError, OverflowError):
        return None
    if pd.isna(parsed):
        return None
    timestamp = pd.Timestamp(parsed)
    if timestamp.tzinfo is not None:
        timestamp = timestamp.tz_convert("UTC").tz_localize(None)
    return timestamp.date()


def _canonical_scalar(value: Any) -> str | None:
    if _is_missing(value):
        return None
    if isinstance(value, str):
        text = unicodedata.normalize("NFKC", value).strip()
        return text or None
    if isinstance(value, (bool, np.bool_)):
        return "True" if bool(value) else "False"
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, Decimal):
        if not value.is_finite():
            return None
        return format(value, "f")
    if isinstance(value, (float, np.floating)):
        numeric = float(value)
        if not math.isfinite(numeric):
            return None
        return str(int(numeric)) if numeric.is_integer() else str(Decimal(str(numeric)))
    if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
        return pd.Timestamp(value).isoformat()
    return unicodedata.normalize("NFKC", str(value)).strip() or None


def _secondary_scalar(value: Any) -> str | None:
    token = _canonical_scalar(value)
    if token is None:
        return None
    return " ".join(token.split()).casefold()


def _prepare_rows(
    frame: pd.DataFrame,
    *,
    amount_column: str,
    date_column: str | None,
    key_columns: Sequence[str],
) -> list[_PreparedRow]:
    rows: list[_PreparedRow] = []
    for position in range(len(frame)):
        source = frame.iloc[position]
        if key_columns:
            key_values = tuple(_canonical_scalar(source[column]) for column in key_columns)
            key = (
                None
                if any(value is None for value in key_values)
                else tuple(value for value in key_values if value is not None)
            )
        else:
            key = None
        rows.append(
            _PreparedRow(
                position=position,
                amount=_decimal_value(source[amount_column]),
                date_value=(
                    None if date_column is None else _date_value(source[date_column])
                ),
                key=key,
            )
        )
    return rows


def _payload_names(
    columns: Sequence[Any], prefix: str, occupied: set[str]
) -> list[tuple[Any, str]]:
    result: list[tuple[Any, str]] = []
    for column in columns:
        base = f"{prefix}{column}"
        candidate = base
        suffix = 2
        while candidate in occupied:
            candidate = f"{base}_{suffix}"
            suffix += 1
        occupied.add(candidate)
        result.append((column, candidate))
    return result


def _pair_layout(
    left: pd.DataFrame, right: pd.DataFrame
) -> tuple[list[tuple[Any, str]], list[tuple[Any, str]], list[str]]:
    occupied = set(_PAIR_METADATA_COLUMNS) | set(_DUPLICATE_METADATA_COLUMNS)
    left_payload = _payload_names(left.columns.tolist(), "left__", occupied)
    right_payload = _payload_names(right.columns.tolist(), "right__", occupied)
    columns = [
        *_PAIR_METADATA_COLUMNS,
        *[name for _, name in left_payload],
        *[name for _, name in right_payload],
    ]
    return left_payload, right_payload, columns


def _secondary_stats(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_position: int,
    right_position: int,
    left_columns: Sequence[str],
    right_columns: Sequence[str],
) -> tuple[int, int]:
    compared = 0
    matched = 0
    left_row = left.iloc[left_position]
    right_row = right.iloc[right_position]
    for left_column, right_column in zip(left_columns, right_columns):
        left_value = _secondary_scalar(left_row[left_column])
        right_value = _secondary_scalar(right_row[right_column])
        if left_value is None or right_value is None:
            continue
        compared += 1
        matched += int(left_value == right_value)
    return compared, matched


def _pair_record(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_rows: Sequence[_PreparedRow],
    right_rows: Sequence[_PreparedRow],
    left_payload: Sequence[tuple[Any, str]],
    right_payload: Sequence[tuple[Any, str]],
    *,
    left_position: int,
    right_position: int,
    match_type: str,
    group_id: str,
    rank: int,
    candidate_count: int,
    score: float,
    reason: str,
    secondary_compared: int,
    secondary_matched: int,
    group_left_amount: Decimal | None = None,
    group_right_amount: Decimal | None = None,
) -> dict[str, Any]:
    left_prepared = left_rows[left_position]
    right_prepared = right_rows[right_position]
    amount_difference = (
        None
        if left_prepared.amount is None or right_prepared.amount is None
        else right_prepared.amount - left_prepared.amount
    )
    date_difference = (
        None
        if left_prepared.date_value is None or right_prepared.date_value is None
        else (right_prepared.date_value - left_prepared.date_value).days
    )
    effective_left_total = (
        left_prepared.amount if group_left_amount is None else group_left_amount
    )
    effective_right_total = (
        right_prepared.amount if group_right_amount is None else group_right_amount
    )
    group_difference = (
        None
        if effective_left_total is None or effective_right_total is None
        else effective_right_total - effective_left_total
    )
    record: dict[str, Any] = {
        "match_type": match_type,
        "candidate_group_id": group_id,
        "left_row_position": left_position,
        "right_row_position": right_position,
        "candidate_rank": rank,
        "candidate_count": candidate_count,
        "match_score": round(float(score), 4),
        "match_reason": reason,
        "left_amount_decimal": left_prepared.amount,
        "right_amount_decimal": right_prepared.amount,
        "amount_difference_decimal": amount_difference,
        "absolute_amount_difference": (
            None if amount_difference is None else abs(amount_difference)
        ),
        "left_date": left_prepared.date_value,
        "right_date": right_prepared.date_value,
        "date_difference_days": date_difference,
        "secondary_compared": secondary_compared,
        "secondary_matched": secondary_matched,
        "group_left_amount_decimal": effective_left_total,
        "group_right_amount_decimal": effective_right_total,
        "group_amount_difference_decimal": group_difference,
    }
    left_source = left.iloc[left_position]
    right_source = right.iloc[right_position]
    for column, output_name in left_payload:
        record[output_name] = left_source[column]
    for column, output_name in right_payload:
        record[output_name] = right_source[column]
    return record


def _pair_frame(records: Sequence[Mapping[str, Any]], columns: Sequence[str]) -> pd.DataFrame:
    return pd.DataFrame(list(records), columns=list(columns))


def _score_candidate(
    *,
    amount_difference: Decimal,
    amount_tolerance: Decimal,
    date_difference_days: int | None,
    date_tolerance_days: int,
    has_dates: bool,
    secondary_compared: int,
    secondary_matched: int,
    has_secondary: bool,
    key_conflict: bool,
) -> float:
    components: list[tuple[float, float]] = []
    if amount_difference == 0:
        amount_quality = 1.0
    elif amount_tolerance > 0:
        amount_quality = max(
            0.5,
            1.0 - float(amount_difference / amount_tolerance) * 0.5,
        )
    else:
        amount_quality = 0.0
    components.append((0.45, amount_quality))

    if has_dates:
        if date_difference_days == 0:
            date_quality = 1.0
        elif date_difference_days is not None and date_tolerance_days > 0:
            date_quality = max(
                0.5,
                1.0 - date_difference_days / date_tolerance_days * 0.5,
            )
        else:
            date_quality = 0.0
        components.append((0.25, date_quality))

    if has_secondary:
        secondary_quality = (
            secondary_matched / secondary_compared if secondary_compared else 0.0
        )
        components.append((0.30, secondary_quality))

    weight = sum(item[0] for item in components)
    score = sum(item_weight * quality for item_weight, quality in components) / weight
    if key_conflict:
        score *= 0.8
    return round(max(0.0, min(1.0, score)), 4)


def _candidate_reason(candidate: _Candidate, *, has_secondary: bool) -> str:
    reasons: list[str] = []
    if candidate.key_conflict:
        reasons.append("双方主键不一致")
    if candidate.amount_difference == 0:
        reasons.append("金额相同")
    else:
        reasons.append(f"金额差 {candidate.amount_difference}")
    if candidate.date_difference_days is not None:
        if candidate.date_difference_days == 0:
            reasons.append("日期相同")
        else:
            reasons.append(f"日期相差 {candidate.date_difference_days} 天")
    if has_secondary:
        reasons.append(
            f"次级字段匹配 {candidate.secondary_matched}/{candidate.secondary_compared}"
        )
    else:
        reasons.append("未配置次级字段")
    reasons.append("未满足保守自动匹配条件，需人工确认")
    return "；".join(reasons)


def _amount_candidates(
    left_amount: Decimal,
    sorted_right: Sequence[tuple[Decimal, int]],
    tolerance: Decimal,
) -> list[int]:
    amounts = [item[0] for item in sorted_right]
    start = bisect_left(amounts, left_amount - tolerance)
    end = bisect_right(amounts, left_amount + tolerance)
    return [position for _, position in sorted_right[start:end]]


def _key_duplicate_state(
    left_rows: Sequence[_PreparedRow], right_rows: Sequence[_PreparedRow]
) -> tuple[Counter[tuple[str, ...]], Counter[tuple[str, ...]], set[tuple[str, ...]]]:
    left_counts = Counter(row.key for row in left_rows if row.key is not None)
    right_counts = Counter(row.key for row in right_rows if row.key is not None)
    ambiguous = {
        key
        for key in set(left_counts) | set(right_counts)
        if left_counts[key] > 1 or right_counts[key] > 1
    }
    return left_counts, right_counts, ambiguous


def _duplicate_frame(
    left: pd.DataFrame,
    right: pd.DataFrame,
    left_rows: Sequence[_PreparedRow],
    right_rows: Sequence[_PreparedRow],
    left_payload: Sequence[tuple[Any, str]],
    right_payload: Sequence[tuple[Any, str]],
    pair_columns: Sequence[str],
    left_counts: Counter[tuple[str, ...]],
    right_counts: Counter[tuple[str, ...]],
    ambiguous_keys: set[tuple[str, ...]],
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []
    for side, frame, rows, payload in (
        ("left", left, left_rows, left_payload),
        ("right", right, right_rows, right_payload),
    ):
        for row in rows:
            if row.key not in ambiguous_keys or row.key is None:
                continue
            left_count = left_counts[row.key]
            right_count = right_counts[row.key]
            if (side == "left" and left_count > 1) or (
                side == "right" and right_count > 1
            ):
                reason = f"主键在{('左' if side == 'left' else '右')}侧重复，无法安全一对一匹配"
            else:
                reason = "对侧主键重复，无法安全一对一匹配"
            record: dict[str, Any] = {
                "source_side": side,
                "source_row_position": row.position,
                "key_values": "｜".join(row.key),
                "duplicate_count_left": left_count,
                "duplicate_count_right": right_count,
                "duplicate_reason": reason,
            }
            source = frame.iloc[row.position]
            for column, output_name in payload:
                record[output_name] = source[column]
            records.append(record)
    columns = [
        *_DUPLICATE_METADATA_COLUMNS,
        *[column for column in pair_columns if column not in _PAIR_METADATA_COLUMNS],
    ]
    return pd.DataFrame(records, columns=columns)


def _only_frame(
    frame: pd.DataFrame,
    rows: Sequence[_PreparedRow],
    positions: Sequence[int],
    payload: Sequence[tuple[Any, str]],
    *,
    side: str,
    has_dates: bool,
    reason_overrides: Mapping[int, str] | None = None,
) -> pd.DataFrame:
    position_column = f"{side}_row_position"
    amount_column = f"{side}_amount_decimal"
    date_column = f"{side}_date"
    columns = [
        position_column,
        "unmatched_reason",
        amount_column,
        date_column,
        *[name for _, name in payload],
    ]
    records: list[dict[str, Any]] = []
    overrides = {} if reason_overrides is None else dict(reason_overrides)
    for position in positions:
        prepared = rows[position]
        if position in overrides:
            reason = overrides[position]
        elif prepared.amount is None:
            reason = "金额为空或无法解析，未参与自动候选"
        elif has_dates and prepared.date_value is None:
            reason = "日期为空或无法解析，未参与自动候选"
        else:
            reason = "未找到满足金额、日期及次级字段条件的候选"
        record: dict[str, Any] = {
            position_column: position,
            "unmatched_reason": reason,
            amount_column: prepared.amount,
            date_column: prepared.date_value,
        }
        source = frame.iloc[position]
        for column, output_name in payload:
            record[output_name] = source[column]
        records.append(record)
    return pd.DataFrame(records, columns=columns)


def _exact_key_classification(
    left_row: _PreparedRow,
    right_row: _PreparedRow,
    *,
    amount_tolerance: Decimal,
    date_tolerance_days: int,
    has_dates: bool,
) -> tuple[str, float, str]:
    if left_row.amount is None or right_row.amount is None:
        return "review", 0.6, "主键完全一致，但金额为空或无法解析，需人工确认"
    amount_difference = abs(right_row.amount - left_row.amount)
    if has_dates and (left_row.date_value is None or right_row.date_value is None):
        return "review", 0.65, "主键完全一致，但日期为空或无法解析，需人工确认"
    date_difference = (
        0
        if not has_dates
        else abs((right_row.date_value - left_row.date_value).days)  # type: ignore[operator]
    )
    if amount_difference == 0 and date_difference == 0:
        return "matched", 1.0, "主键完全一致；金额相同；日期相同或未配置"
    if amount_difference <= amount_tolerance and date_difference == 0:
        return (
            "amount_difference",
            0.98,
            f"主键完全一致；金额差 {amount_difference} 在容差内；日期相同或未配置",
        )
    if amount_difference == 0 and date_difference <= date_tolerance_days:
        return (
            "date_difference",
            0.98,
            f"主键完全一致；金额相同；日期相差 {date_difference} 天，在容差内",
        )
    if amount_difference <= amount_tolerance and date_difference <= date_tolerance_days:
        return (
            "review",
            0.9,
            f"主键一致，但金额差 {amount_difference} 且日期相差 {date_difference} 天，需人工确认",
        )
    return (
        "review",
        0.7,
        f"主键一致，但金额差 {amount_difference} 或日期差 {date_difference} 天超出容差，需人工确认",
    )


def _split_component_positions(
    target: _PreparedRow,
    component_rows: Sequence[_PreparedRow],
    available: set[int],
    *,
    date_tolerance_days: int,
    has_dates: bool,
    amount_tolerance: Decimal,
    max_candidates: int,
) -> tuple[list[int], bool]:
    if target.amount is None or target.amount == 0:
        return [], False
    eligible: list[int] = []
    target_sign = 1 if target.amount > 0 else -1
    for position in available:
        component = component_rows[position]
        if component.amount is None or component.amount == 0:
            continue
        component_sign = 1 if component.amount > 0 else -1
        if component_sign != target_sign:
            continue
        if abs(component.amount) > abs(target.amount) + amount_tolerance:
            continue
        if has_dates:
            if target.date_value is None or component.date_value is None:
                continue
            if abs((component.date_value - target.date_value).days) > date_tolerance_days:
                continue
        eligible.append(position)
    eligible.sort(
        key=lambda position: (
            abs(abs(component_rows[position].amount) - abs(target.amount) / 2),  # type: ignore[arg-type]
            position,
        )
    )
    truncated = len(eligible) > max_candidates
    return eligible[:max_candidates], truncated


def reconcile_tables(
    left: pd.DataFrame,
    right: pd.DataFrame,
    *,
    left_amount: str,
    right_amount: str,
    left_date: str | None = None,
    right_date: str | None = None,
    left_key_columns: str | Sequence[str] | None = None,
    right_key_columns: str | Sequence[str] | None = None,
    left_secondary_columns: str | Sequence[str] | None = None,
    right_secondary_columns: str | Sequence[str] | None = None,
    amount_tolerance: Decimal | int | float | str = Decimal("0"),
    date_tolerance_days: int = 0,
    enable_split_candidates: bool = False,
    max_candidates_per_row: int = 20,
    max_candidate_pairs: int = 100_000,
    max_split_combinations: int = 10_000,
) -> ReconciliationResult:
    """Reconcile two tables into safe, explainable outcome buckets.

    Matching order is deterministic:

    1. unique, complete primary keys are paired first;
    2. remaining rows are blocked by Decimal amount and optional date tolerance;
    3. a keyless pair is auto-matched only when it is mutually unique, amount
       and date are exact, and every comparable secondary field agrees;
    4. all ambiguous, conflicting, or split candidates are returned in
       ``review`` rather than silently accepted.

    ``match_score`` is a transparent rule score, not a statistical
    probability.  All row positions are zero-based positions in the supplied
    DataFrames.  Inputs are never modified.
    """

    _require_dataframe(left, "left")
    _require_dataframe(right, "right")
    if not isinstance(left_amount, str) or not left_amount:
        raise TypeError("left_amount 必须是非空字段名")
    if not isinstance(right_amount, str) or not right_amount:
        raise TypeError("right_amount 必须是非空字段名")
    if left_date is not None and (
        not isinstance(left_date, str) or not left_date
    ):
        raise TypeError("left_date 必须是非空字段名或 None")
    if right_date is not None and (
        not isinstance(right_date, str) or not right_date
    ):
        raise TypeError("right_date 必须是非空字段名或 None")
    if (left_date is None) != (right_date is None):
        raise ValueError("left_date 与 right_date 必须同时提供")

    left_keys, right_keys = _paired_columns(
        left_key_columns,
        right_key_columns,
        left_argument="left_key_columns",
        right_argument="right_key_columns",
    )
    left_secondary, right_secondary = _paired_columns(
        left_secondary_columns,
        right_secondary_columns,
        left_argument="left_secondary_columns",
        right_argument="right_secondary_columns",
    )
    _validate_columns(
        left,
        [left_amount, *([left_date] if left_date else []), *left_keys, *left_secondary],
        argument="left columns",
    )
    _validate_columns(
        right,
        [
            right_amount,
            *([right_date] if right_date else []),
            *right_keys,
            *right_secondary,
        ],
        argument="right columns",
    )
    left_roles = [
        left_amount,
        *([left_date] if left_date else []),
        *left_keys,
        *left_secondary,
    ]
    right_roles = [
        right_amount,
        *([right_date] if right_date else []),
        *right_keys,
        *right_secondary,
    ]
    if len(left_roles) != len(set(left_roles)):
        raise ValueError("左侧金额、日期、主键和次级字段角色必须互不相同")
    if len(right_roles) != len(set(right_roles)):
        raise ValueError("右侧金额、日期、主键和次级字段角色必须互不相同")
    tolerance = _amount_tolerance(amount_tolerance)
    if isinstance(date_tolerance_days, bool) or not isinstance(date_tolerance_days, int):
        raise TypeError("date_tolerance_days 必须是非负整数")
    if date_tolerance_days < 0:
        raise ValueError("date_tolerance_days 必须是非负整数")
    if not isinstance(enable_split_candidates, bool):
        raise TypeError("enable_split_candidates 必须是布尔值")
    max_candidates = _positive_integer(max_candidates_per_row, "max_candidates_per_row")
    candidate_limit = _positive_integer(max_candidate_pairs, "max_candidate_pairs")
    split_limit = _positive_integer(max_split_combinations, "max_split_combinations")
    has_dates = left_date is not None
    has_secondary = bool(left_secondary)

    left_rows = _prepare_rows(
        left,
        amount_column=left_amount,
        date_column=left_date,
        key_columns=left_keys,
    )
    right_rows = _prepare_rows(
        right,
        amount_column=right_amount,
        date_column=right_date,
        key_columns=right_keys,
    )
    left_payload, right_payload, pair_columns = _pair_layout(left, right)
    left_counts, right_counts, ambiguous_keys = _key_duplicate_state(
        left_rows, right_rows
    )
    duplicate_left_positions = {
        row.position for row in left_rows if row.key in ambiguous_keys
    }
    duplicate_right_positions = {
        row.position for row in right_rows if row.key in ambiguous_keys
    }
    duplicates = _duplicate_frame(
        left,
        right,
        left_rows,
        right_rows,
        left_payload,
        right_payload,
        pair_columns,
        left_counts,
        right_counts,
        ambiguous_keys,
    )

    matched_records: list[dict[str, Any]] = []
    amount_records: list[dict[str, Any]] = []
    date_records: list[dict[str, Any]] = []
    review_records: list[dict[str, Any]] = []
    used_left: set[int] = set(duplicate_left_positions)
    used_right: set[int] = set(duplicate_right_positions)
    review_left: set[int] = set()
    review_right: set[int] = set()
    review_group_number = 0

    # Phase 1: complete, unique primary keys always take precedence.
    left_key_map = {
        row.key: row.position
        for row in left_rows
        if row.key is not None and row.position not in used_left
    }
    right_key_map = {
        row.key: row.position
        for row in right_rows
        if row.key is not None and row.position not in used_right
    }
    for key, left_position in left_key_map.items():
        if key not in right_key_map:
            continue
        right_position = right_key_map[key]
        category, score, reason = _exact_key_classification(
            left_rows[left_position],
            right_rows[right_position],
            amount_tolerance=tolerance,
            date_tolerance_days=date_tolerance_days,
            has_dates=has_dates,
        )
        if category == "review":
            review_group_number += 1
            group_id = f"K{review_group_number:05d}"
        else:
            group_id = ""
        compared, secondary_matches = _secondary_stats(
            left,
            right,
            left_position,
            right_position,
            left_secondary,
            right_secondary,
        )
        record = _pair_record(
            left,
            right,
            left_rows,
            right_rows,
            left_payload,
            right_payload,
            left_position=left_position,
            right_position=right_position,
            match_type="primary_key_exact",
            group_id=group_id,
            rank=1,
            candidate_count=1,
            score=score,
            reason=reason,
            secondary_compared=compared,
            secondary_matched=secondary_matches,
        )
        if category == "matched":
            matched_records.append(record)
        elif category == "amount_difference":
            amount_records.append(record)
        elif category == "date_difference":
            date_records.append(record)
        else:
            review_records.append(record)
            review_left.add(left_position)
            review_right.add(right_position)
        used_left.add(left_position)
        used_right.add(right_position)

    # Phase 2: amount/date blocking, then conservative secondary matching.
    available_left = set(range(len(left))) - used_left
    available_right = set(range(len(right))) - used_right
    sorted_right_amounts = sorted(
        (right_rows[position].amount, position)
        for position in available_right
        if right_rows[position].amount is not None
    )
    candidates: list[_Candidate] = []
    candidates_by_left: dict[int, list[_Candidate]] = defaultdict(list)
    candidates_by_right: dict[int, list[_Candidate]] = defaultdict(list)
    for left_position in sorted(available_left):
        left_row = left_rows[left_position]
        if left_row.amount is None:
            continue
        for right_position in _amount_candidates(
            left_row.amount, sorted_right_amounts, tolerance
        ):
            right_row = right_rows[right_position]
            if right_row.amount is None:
                continue
            date_difference: int | None = None
            if has_dates:
                if left_row.date_value is None or right_row.date_value is None:
                    continue
                date_difference = abs((right_row.date_value - left_row.date_value).days)
                if date_difference > date_tolerance_days:
                    continue
            compared, secondary_matches = _secondary_stats(
                left,
                right,
                left_position,
                right_position,
                left_secondary,
                right_secondary,
            )
            amount_difference = abs(right_row.amount - left_row.amount)
            key_conflict = (
                left_row.key is not None
                and right_row.key is not None
                and left_row.key != right_row.key
            )
            score = _score_candidate(
                amount_difference=amount_difference,
                amount_tolerance=tolerance,
                date_difference_days=date_difference,
                date_tolerance_days=date_tolerance_days,
                has_dates=has_dates,
                secondary_compared=compared,
                secondary_matched=secondary_matches,
                has_secondary=has_secondary,
                key_conflict=key_conflict,
            )
            candidate = _Candidate(
                left_position=left_position,
                right_position=right_position,
                amount_difference=amount_difference,
                date_difference_days=date_difference,
                secondary_compared=compared,
                secondary_matched=secondary_matches,
                score=score,
                key_conflict=key_conflict,
            )
            candidates.append(candidate)
            if len(candidates) > candidate_limit:
                raise ValueError(
                    f"候选对超过 max_candidate_pairs={candidate_limit}；"
                    "请缩小日期/金额容差、增加主键或先筛选数据"
                )
            candidates_by_left[left_position].append(candidate)
            candidates_by_right[right_position].append(candidate)

    auto_candidates: list[_Candidate] = []
    for candidate in candidates:
        exact_date = not has_dates or candidate.date_difference_days == 0
        complete_secondary = (
            has_secondary
            and candidate.secondary_compared == len(left_secondary)
            and candidate.secondary_matched == candidate.secondary_compared
        )
        if (
            len(candidates_by_left[candidate.left_position]) == 1
            and len(candidates_by_right[candidate.right_position]) == 1
            and candidate.amount_difference == 0
            and exact_date
            and complete_secondary
            and not candidate.key_conflict
        ):
            auto_candidates.append(candidate)

    auto_left = {candidate.left_position for candidate in auto_candidates}
    auto_right = {candidate.right_position for candidate in auto_candidates}
    for candidate in sorted(
        auto_candidates, key=lambda item: (item.left_position, item.right_position)
    ):
        reason = (
            "无可用共同主键；金额相同；日期相同或未配置；"
            f"次级字段匹配 {candidate.secondary_matched}/{candidate.secondary_compared}；"
            "且双方均为唯一候选"
        )
        matched_records.append(
            _pair_record(
                left,
                right,
                left_rows,
                right_rows,
                left_payload,
                right_payload,
                left_position=candidate.left_position,
                right_position=candidate.right_position,
                match_type="secondary_unique_exact",
                group_id="",
                rank=1,
                candidate_count=1,
                score=min(candidate.score, 0.95),
                reason=reason,
                secondary_compared=candidate.secondary_compared,
                secondary_matched=candidate.secondary_matched,
            )
        )
        used_left.add(candidate.left_position)
        used_right.add(candidate.right_position)

    candidate_groups_truncated = 0
    truncated_right_positions: set[int] = set()
    for left_position in sorted(candidates_by_left):
        if left_position in auto_left:
            continue
        remaining_candidates = [
            candidate
            for candidate in candidates_by_left[left_position]
            if candidate.right_position not in auto_right
        ]
        if not remaining_candidates:
            continue
        ranked = sorted(
            remaining_candidates,
            key=lambda item: (
                -item.score,
                item.amount_difference,
                item.date_difference_days if item.date_difference_days is not None else 0,
                item.right_position,
            ),
        )
        total = len(ranked)
        if total > max_candidates:
            candidate_groups_truncated += 1
            truncated_right_positions.update(
                candidate.right_position for candidate in ranked[max_candidates:]
            )
        review_group_number += 1
        group_id = f"R{review_group_number:05d}"
        for rank, candidate in enumerate(ranked[:max_candidates], start=1):
            reason = _candidate_reason(candidate, has_secondary=has_secondary)
            if total > max_candidates:
                reason += f"；仅展示前 {max_candidates}/{total} 个候选"
            review_records.append(
                _pair_record(
                    left,
                    right,
                    left_rows,
                    right_rows,
                    left_payload,
                    right_payload,
                    left_position=candidate.left_position,
                    right_position=candidate.right_position,
                    match_type="one_to_one_review",
                    group_id=group_id,
                    rank=rank,
                    candidate_count=total,
                    score=candidate.score,
                    reason=reason,
                    secondary_compared=candidate.secondary_compared,
                    secondary_matched=candidate.secondary_matched,
                )
            )
            review_left.add(candidate.left_position)
            review_right.add(candidate.right_position)
        used_left.add(left_position)
        used_right.update(
            candidate.right_position for candidate in ranked[:max_candidates]
        )

    # Phase 3: optional, bounded split suggestions.  They are always review-only.
    split_group_count = 0
    split_combinations_evaluated = 0
    split_limit_hit = False
    split_candidates_truncated = 0
    split_used_left: set[int] = set()
    split_used_right: set[int] = set()
    if enable_split_candidates:
        split_left_pool = set(range(len(left))) - used_left
        split_right_pool = set(range(len(right))) - used_right

        def append_split_group(
            left_positions: tuple[int, ...],
            right_positions: tuple[int, ...],
            match_type: str,
        ) -> None:
            nonlocal review_group_number, split_group_count
            review_group_number += 1
            split_group_count += 1
            group_id = f"S{review_group_number:05d}"
            left_total = sum(
                (left_rows[position].amount for position in left_positions),
                Decimal("0"),
            )
            right_total = sum(
                (right_rows[position].amount for position in right_positions),
                Decimal("0"),
            )
            reason = (
                f"{match_type} 拆分候选：组内金额差 {abs(right_total - left_total)} "
                "在容差内；拆分建议永远需要人工确认"
            )
            component_pairs = (
                [(left_positions[0], position) for position in right_positions]
                if len(left_positions) == 1
                else [(position, right_positions[0]) for position in left_positions]
            )
            for rank, (left_position, right_position) in enumerate(
                component_pairs, start=1
            ):
                compared, secondary_matches = _secondary_stats(
                    left,
                    right,
                    left_position,
                    right_position,
                    left_secondary,
                    right_secondary,
                )
                review_records.append(
                    _pair_record(
                        left,
                        right,
                        left_rows,
                        right_rows,
                        left_payload,
                        right_payload,
                        left_position=left_position,
                        right_position=right_position,
                        match_type=match_type,
                        group_id=group_id,
                        rank=rank,
                        candidate_count=2,
                        score=0.8 if right_total == left_total else 0.72,
                        reason=reason,
                        secondary_compared=compared,
                        secondary_matched=secondary_matches,
                        group_left_amount=left_total,
                        group_right_amount=right_total,
                    )
                )
            review_left.update(left_positions)
            review_right.update(right_positions)
            split_used_left.update(left_positions)
            split_used_right.update(right_positions)

        for left_position in sorted(split_left_pool):
            component_positions, truncated = _split_component_positions(
                left_rows[left_position],
                right_rows,
                split_right_pool,
                date_tolerance_days=date_tolerance_days,
                has_dates=has_dates,
                amount_tolerance=tolerance,
                max_candidates=max_candidates,
            )
            split_candidates_truncated += int(truncated)
            for first, second in combinations(component_positions, 2):
                split_combinations_evaluated += 1
                if split_combinations_evaluated > split_limit:
                    split_limit_hit = True
                    break
                target_amount = left_rows[left_position].amount
                combined = right_rows[first].amount + right_rows[second].amount  # type: ignore[operator]
                if target_amount is not None and abs(combined - target_amount) <= tolerance:
                    append_split_group(
                        (left_position,), (first, second), "split_1_to_2"
                    )
            if split_limit_hit:
                break

        if not split_limit_hit:
            for right_position in sorted(split_right_pool):
                component_positions, truncated = _split_component_positions(
                    right_rows[right_position],
                    left_rows,
                    split_left_pool,
                    date_tolerance_days=date_tolerance_days,
                    has_dates=has_dates,
                    amount_tolerance=tolerance,
                    max_candidates=max_candidates,
                )
                split_candidates_truncated += int(truncated)
                for first, second in combinations(component_positions, 2):
                    split_combinations_evaluated += 1
                    if split_combinations_evaluated > split_limit:
                        split_limit_hit = True
                        break
                    target_amount = right_rows[right_position].amount
                    combined = left_rows[first].amount + left_rows[second].amount  # type: ignore[operator]
                    if target_amount is not None and abs(combined - target_amount) <= tolerance:
                        append_split_group(
                            (first, second), (right_position,), "split_2_to_1"
                        )
                if split_limit_hit:
                    break

        used_left.update(split_used_left)
        used_right.update(split_used_right)

    unmatched_left_positions = sorted(set(range(len(left))) - used_left)
    unmatched_right_positions = sorted(set(range(len(right))) - used_right)
    left_only = _only_frame(
        left,
        left_rows,
        unmatched_left_positions,
        left_payload,
        side="left",
        has_dates=has_dates,
    )
    truncated_right_reasons = {
        position: (
            "存在候选，但超出 max_candidates_per_row 展示上限；"
            "请缩小容差、增加主键或提高限制后重跑"
        )
        for position in truncated_right_positions
        if position in unmatched_right_positions
    }
    right_only = _only_frame(
        right,
        right_rows,
        unmatched_right_positions,
        right_payload,
        side="right",
        has_dates=has_dates,
        reason_overrides=truncated_right_reasons,
    )

    matched = _pair_frame(matched_records, pair_columns)
    amount_difference = _pair_frame(amount_records, pair_columns)
    date_difference = _pair_frame(date_records, pair_columns)
    review = _pair_frame(review_records, pair_columns)
    review_group_ids = {
        str(value)
        for value in review.get("candidate_group_id", pd.Series(dtype="object"))
        if not _is_missing(value)
    }
    summary = {
        "left_rows": len(left),
        "right_rows": len(right),
        "matched_count": len(matched),
        "amount_difference_count": len(amount_difference),
        "date_difference_count": len(date_difference),
        "review_candidate_rows": len(review),
        "review_group_count": len(review_group_ids),
        "review_left_rows": len(review_left),
        "review_right_rows": len(review_right),
        "left_only_count": len(left_only),
        "right_only_count": len(right_only),
        "duplicate_rows_count": len(duplicates),
        "candidate_pairs_evaluated": len(candidates),
        "candidate_groups_truncated": candidate_groups_truncated,
        "split_candidate_groups": split_group_count,
        "split_combinations_evaluated": min(split_combinations_evaluated, split_limit),
        "split_limit_hit": split_limit_hit,
        "split_candidates_truncated": split_candidates_truncated,
        "amount_tolerance": str(tolerance),
        "date_tolerance_days": date_tolerance_days,
        "left_key_columns": list(left_keys),
        "right_key_columns": list(right_keys),
        "left_secondary_columns": list(left_secondary),
        "right_secondary_columns": list(right_secondary),
        "score_note": "match_score 是确定性规则分数，不是统计概率；拆分候选不能自动确认",
    }
    return ReconciliationResult(
        matched=matched,
        amount_difference=amount_difference,
        date_difference=date_difference,
        review=review,
        left_only=left_only,
        right_only=right_only,
        duplicates=duplicates,
        summary=summary,
    )
