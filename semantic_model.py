"""Reusable semantic primitives for multi-table business projects.

The module deliberately separates table-role inference, field semantics and
relationship evidence.  Domain report builders consume the evidence instead
of selecting one numeric-rich "main table".
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Mapping, Sequence

import pandas as pd

_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:/.]+")


def normalise_label(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


def clean_key(series: pd.Series) -> pd.Series:
    return series.astype("string").fillna("").str.replace(r"\s+", "", regex=True).str.strip()


def find_column(frame: pd.DataFrame, aliases: Sequence[str]) -> str | None:
    columns = {normalise_label(column): str(column) for column in frame.columns}
    for alias in aliases:
        hit = columns.get(normalise_label(alias))
        if hit is not None:
            return hit
    partial = {
        column
        for key, column in columns.items()
        for alias in aliases
        if len(normalise_label(alias)) >= 2 and normalise_label(alias) in key
    }
    return next(iter(partial)) if len(partial) == 1 else None


@dataclass(frozen=True)
class RoleEvidence:
    role: str
    table_index: int
    score: int
    matched_fields: tuple[str, ...]
    table_name: str


def infer_table_roles(
    frames: Sequence[pd.DataFrame],
    source_names: Sequence[str],
    specifications: Mapping[str, Mapping[str, Sequence[str]]],
    *,
    required_roles: Sequence[str] = (),
) -> tuple[dict[str, int], list[RoleEvidence]]:
    """Assign independent roles using field evidence and optional name cues."""
    roles: dict[str, int] = {}
    evidence: list[RoleEvidence] = []
    used: set[int] = set()
    names = list(source_names) + [f"表{index + 1}" for index in range(max(0, len(frames) - len(source_names)))]
    for role, spec in specifications.items():
        aliases = spec.get("fields", ())
        required = spec.get("required", ())
        name_tokens = spec.get("name_tokens", ())
        candidates: list[tuple[int, int, int, tuple[str, ...]]] = []
        for index, frame in enumerate(frames):
            if index in used or not isinstance(frame, pd.DataFrame) or frame.empty:
                continue
            matched = tuple(alias for alias in aliases if find_column(frame, (alias,)) is not None)
            if required and not all(find_column(frame, (alias,)) is not None for alias in required):
                continue
            name = names[index] if index < len(names) else f"表{index + 1}"
            name_score = 6 if any(normalise_label(token) in normalise_label(name) for token in name_tokens) else 0
            score = len(matched) * 3 + name_score
            if score:
                candidates.append((score, len(frame), index, matched))
        if candidates:
            score, _, index, matched = max(candidates)
            roles[role] = index
            used.add(index)
            evidence.append(RoleEvidence(role, index, score, matched, names[index]))
        elif role in required_roles:
            raise ValueError(f"无法识别必需事实域：{role}")
    return roles, evidence


@dataclass(frozen=True)
class RelationshipEvidence:
    left_table: str
    left_key: str
    right_table: str
    right_key: str
    left_row_coverage: float
    left_unique_coverage: float
    right_key_unique: bool
    accepted: bool
    reason: str


def assess_relationship(
    left: pd.DataFrame,
    left_key: str,
    right: pd.DataFrame,
    right_key: str,
    *,
    left_name: str,
    right_name: str,
    require_right_unique: bool = True,
) -> RelationshipEvidence:
    """Measure a directional business-key relationship without joining facts."""
    left_values = clean_key(left[left_key])
    right_values = clean_key(right[right_key])
    left_nonblank = left_values[left_values.ne("")]
    right_nonblank = right_values[right_values.ne("")]
    right_set = set(right_nonblank)
    row_coverage = float(left_nonblank.isin(right_set).mean()) if len(left_nonblank) else 0.0
    unique_left = set(left_nonblank)
    unique_coverage = len(unique_left & right_set) / len(unique_left) if unique_left else 0.0
    right_unique = not right_nonblank.duplicated().any()
    accepted = row_coverage >= 0.8 and (right_unique or not require_right_unique)
    reason = (
        f"行覆盖{row_coverage:.1%}，唯一键覆盖{unique_coverage:.1%}，"
        f"右侧键{'唯一' if right_unique else '不唯一'}"
    )
    return RelationshipEvidence(
        left_name,
        left_key,
        right_name,
        right_key,
        row_coverage,
        unique_coverage,
        right_unique,
        accepted,
        reason,
    )


__all__ = [
    "RelationshipEvidence",
    "RoleEvidence",
    "assess_relationship",
    "clean_key",
    "find_column",
    "infer_table_roles",
    "normalise_label",
]
