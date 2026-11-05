"""Shared business invariants, independent of reports, HTTP and AI providers.

Adapters may identify columns; they must not redefine money, missing values,
event identity or master-data conflict policy. No guessed unit conversions.
"""
from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP
import math
import unicodedata
from collections.abc import Mapping
from typing import Sequence

import pandas as pd

class BusinessInvariantError(ValueError):
    """A business precondition failed; do not downgrade it to a completed task."""


def parse_numeric_values(series: pd.Series, *, allow_percent=True) -> pd.Series:
    """Notation only. Units, exchange rates and booleans are not numbers."""
    if pd.api.types.is_bool_dtype(series.dtype):
        return pd.Series(float('nan'), index=series.index)
    raw = series.map(lambda v: unicodedata.normalize('NFKC', v) if isinstance(v, str) else v).astype('string').str.strip()
    parentheses = raw.str.fullmatch(r'\([¥￥]?[\d,，.\s]+\)', na=False)
    raw = raw.where(~parentheses, '-' + raw.str.slice(1, -1))
    percent = raw.str.contains(r'[%％]', na=False)
    notation = raw.str.replace(r'[,，\s]', '', regex=True)
    has_grouping = raw.str.contains(',', regex=False, na=False)
    valid_grouping = raw.str.fullmatch(r'[+-]?[¥￥]?\d{1,3}(?:,\d{3})+(?:\.\d+)?%?', na=False)
    valid_percent = notation.str.fullmatch(r'[+-]?(?:\d+(?:\.\d*)?|\.\d+)[%％]', na=False)
    clean = raw.str.replace(r'[¥￥,，%％\s]', '', regex=True)
    values = pd.to_numeric(clean, errors='coerce').astype('Float64')
    values = values.mask(percent & (~valid_percent | (not allow_percent)))
    values = values.mask(has_grouping & ~valid_grouping)
    values = values.where(~percent, values / 100.0)
    return values.where(values.map(lambda v: pd.notna(v) and math.isfinite(float(v))))


def numbers(series: pd.Series) -> pd.Series:
    values = parse_numeric_values(series, allow_percent=False).astype('float64')
    return values.where(values.map(math.isfinite))


def number(value) -> float:
    return float(numbers(pd.Series([value], dtype=object)).iloc[0])


def complete_sum(values) -> float:
    """No events = zero; an event with unknown amount = unknown, not zero."""
    values = pd.Series(values, dtype='Float64')
    if values.empty:
        return 0.0
    return float(values.sum()) if values.notna().all() and values.map(math.isfinite).all() else float('nan')


def public_evidence(value):
    """Unknown quantities cross JSON boundaries as null, never NaN or zero."""
    if isinstance(value, Mapping):
        return {str(k): public_evidence(v) for k, v in value.items()}
    if isinstance(value, (list, tuple)):
        return [public_evidence(v) for v in value]
    if value is pd.NA or value is pd.NaT:
        return None
    if isinstance(value, float) and not math.isfinite(value):
        return None
    return value


def unique_master(frame: pd.DataFrame, keys: Sequence[str], *, label: str) -> pd.DataFrame:
    if frame[list(keys)].isna().any(axis=None) or frame[list(keys)].astype('string').apply(lambda s: s.str.strip().eq('')).any(axis=None):
        raise BusinessInvariantError(f'{label}：主键缺失')
    clean = frame.drop_duplicates().copy()
    if clean.duplicated(list(keys), keep=False).any():
        raise BusinessInvariantError(f'{label}：同主键内容冲突，须确认版本或生效期，禁止按行顺序选择')
    return clean


def deduplicate_events(frame: pd.DataFrame, keys: Sequence[str], *, label: str,
                       ignore: Sequence[str] = ()) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Keep separate document lines; isolate conflicting identities by refusing.

    Display comments/provenance can be ignored explicitly, business amounts and
    statuses cannot. Missing keys cannot prove a duplicate and are kept for review.
    """
    keys = list(keys)
    if not keys or any(k not in frame for k in keys):
        raise BusinessInvariantError(f'{label}：缺少已声明的业务键')
    valid_key = frame[keys].notna().all(axis=1) & frame[keys].astype('string').apply(lambda s: s.str.strip().ne('')).all(axis=1)
    business = [c for c in frame if c not in ignore]
    candidates = frame.loc[valid_key, business].drop_duplicates()
    if candidates.duplicated(keys, keep=False).any():
        raise BusinessInvariantError(f'{label}：同业务键内容冲突，须确认明细行或版本，不自动删除')
    duplicate = valid_key & frame.duplicated(keys, keep='first')
    return frame.loc[~duplicate].copy(), frame.loc[duplicate].copy()


def allocate_money(amount, weights, *, quantum=Decimal('0.01'), method='largest_remainder') -> list[Decimal]:
    """Stable largest-remainder allocation in integer minor units.

    Conserves source cents, never creates a negative rounding tail, and never
    allocates to zero-weight lines. The caller validates refundable capacity.
    """
    try:
        value = Decimal(str(amount))
        weights = [Decimal(str(w)) for w in weights]
    except (InvalidOperation, ValueError):
        raise BusinessInvariantError('金额分摊存在无效数值') from None
    if not value.is_finite() or value < 0 or any(not w.is_finite() or w < 0 for w in weights):
        raise BusinessInvariantError('金额分摊要求非负有限金额和权重')
    try:
        rounded = value.quantize(quantum, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        raise BusinessInvariantError('金额超过当前货币精度范围') from None
    if value != rounded:
        raise BusinessInvariantError('源金额精度超过最小货币单位，须先确认舍入口径')
    total = sum(weights, Decimal(0))
    if not weights or total <= 0:
        if value == 0:
            return [Decimal(0) for _ in weights]
        raise BusinessInvariantError('金额分摊缺少正数权重')
    if method == 'last_residual':
        result, remaining = [Decimal(0) for _ in weights], value
        eligible = [i for i, w in enumerate(weights) if w > 0]
        for i in eligible[:-1]:
            part = min((value * weights[i] / total).quantize(quantum, rounding=ROUND_HALF_UP), remaining)
            result[i] = part
            remaining -= part
        result[eligible[-1]] = remaining
        return result
    if method != 'largest_remainder':
        raise BusinessInvariantError('未支持的分摊舍入口径')
    units = int(value / quantum)
    exact = [Decimal(units) * w / total for w in weights]
    floors = [int(x.to_integral_value(rounding=ROUND_DOWN)) for x in exact]
    remainder = units - sum(floors)
    order = sorted(range(len(weights)), key=lambda i: (-(exact[i] - floors[i]), i))
    for i in order[:remainder]:
        floors[i] += 1
    allocated = [Decimal(x) * quantum for x in floors]
    if sum(allocated, Decimal(0)) != value:
        raise BusinessInvariantError('金额分摊不守恒')
    return allocated


def period_union(*frames: pd.DataFrame, column='月份') -> pd.DataFrame:
    """The calendar belongs to all fact domains, never only the sales fact."""
    periods = [f[[column]] for f in frames if column in f]
    return (pd.concat(periods, ignore_index=True).dropna().drop_duplicates().sort_values(column).reset_index(drop=True)
            if periods else pd.DataFrame(columns=[column]))


def through_cutoff(frame: pd.DataFrame, field: str, cutoff) -> tuple[pd.DataFrame, pd.DataFrame]:
    dates = pd.to_datetime(frame[field], errors='coerce', format='mixed')
    end = pd.Timestamp(cutoff).normalize() + pd.Timedelta(days=1)
    included = dates.notna() & dates.lt(end)
    rejected = frame.loc[~included].copy()
    if not rejected.empty:
        rejected['期间排除原因'] = dates.loc[~included].map(lambda d: '日期缺失或无效' if pd.isna(d) else '超过分析截止日')
    return frame.loc[included].copy(), rejected
