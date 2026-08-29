"""High-value, side-effect-free analytics for spreadsheet data.

The functions in this module are deliberately conservative: business
identifiers, leading-zero strings, and integer-like text longer than 15 digits
are never coerced to floating-point numbers.  Every function works on a copy or
read-only view of its input and returns a new DataFrame or structured result.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
import math
import re
from types import MappingProxyType
from typing import Any, Literal

import numpy as np
import pandas as pd
from pandas.api import types as ptypes


_IDENTIFIER_NAME = re.compile(
    r"(^|[_\s])(id|uuid|code|key)($|[_\s])|编号|编码|代码|单号|证件|身份证|手机号|电话|邮编",
    re.I,
)
_LEADING_ZERO_INTEGER = re.compile(r"^[+-]?0\d+$")
_INTEGER_TEXT = re.compile(r"^[+-]?\d+$")
_DATE_HINT = re.compile(r"[-/:年月日Tt]")
_SEVERITY_ORDER = {"high": 0, "medium": 1, "low": 2, "info": 3}


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
    if isinstance(value, (bool, np.bool_)):
        return bool(value)
    if isinstance(value, (int, np.integer)):
        numeric = int(value)
        return str(numeric) if abs(numeric) > 9_007_199_254_740_991 else numeric
    if isinstance(value, str):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
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


def _require_dataframe(
    frame: pd.DataFrame,
    name: str = "df",
    *,
    allow_duplicate_columns: bool = False,
) -> None:
    if not isinstance(frame, pd.DataFrame):
        raise TypeError(f"{name} 必须是 pandas DataFrame")
    if not allow_duplicate_columns and frame.columns.duplicated().any():
        duplicate_labels = frame.columns[frame.columns.duplicated()]
        duplicates = list(dict.fromkeys(str(value) for value in duplicate_labels))
        raise ValueError(f"{name} 含重复列名，请先重命名：{duplicates}")


def _column_list(columns: str | Sequence[str], *, argument: str) -> list[str]:
    result = [columns] if isinstance(columns, str) else list(columns)
    if not result:
        raise ValueError(f"{argument} 不能为空")
    if len(set(result)) != len(result):
        raise ValueError(f"{argument} 不能包含重复字段")
    return result


def _validate_columns(
    frame: pd.DataFrame, columns: Sequence[str], *, argument: str
) -> None:
    missing = [column for column in columns if column not in frame.columns]
    if missing:
        raise KeyError(f"{argument} 包含不存在的列：{missing}")


def _missing_mask(series: pd.Series) -> pd.Series:
    return series.map(_is_missing).astype(bool)


def _normalised_for_analysis(series: pd.Series) -> pd.Series:
    result = series.copy(deep=True)
    blanks = result.map(
        lambda value: isinstance(value, str) and not value.strip(),
        na_action="ignore",
    ).fillna(False)
    if bool(blanks.any()):
        result.loc[blanks] = pd.NA
    return result


def _safe_numeric_series(
    series: pd.Series,
    column_name: str,
    *,
    minimum_success_ratio: float = 0.8,
) -> pd.Series | None:
    """Return numeric values without risking identifier precision."""

    if _IDENTIFIER_NAME.search(str(column_name)) or ptypes.is_bool_dtype(series.dtype):
        return None
    normalised = _normalised_for_analysis(series)
    non_missing = normalised.dropna()
    if non_missing.empty:
        return None
    if ptypes.is_numeric_dtype(normalised.dtype):
        if ptypes.is_integer_dtype(normalised.dtype):
            absolute = non_missing.abs()
            if bool((absolute > 9_007_199_254_740_991).any()):
                # Correlation/outlier algorithms eventually use floating point;
                # values above 2**53 are far more likely to be identifiers and
                # cannot be represented exactly as float64.
                return None
        parsed_numeric = pd.to_numeric(normalised, errors="coerce")
        finite = np.isfinite(parsed_numeric.astype(float))
        return parsed_numeric.where(finite)

    text = non_missing.astype(str).str.strip()
    integer_like = text.map(lambda value: bool(_INTEGER_TEXT.fullmatch(value)))
    if bool(text.map(lambda value: bool(_LEADING_ZERO_INTEGER.fullmatch(value))).any()):
        return None
    if bool(
        (integer_like & (text.str.lstrip("+-").str.len() > 15)).any()
    ):
        return None
    parsed = pd.to_numeric(normalised, errors="coerce")
    parsed = parsed.where(np.isfinite(parsed.astype(float)))
    success_ratio = float(parsed.notna().sum() / len(non_missing))
    if success_ratio < minimum_success_ratio or not bool(parsed.notna().any()):
        return None
    return parsed


def _safe_datetime_series(
    series: pd.Series, *, require_hint: bool = True
) -> pd.Series | None:
    normalised = _normalised_for_analysis(series)
    non_missing = normalised.dropna()
    if non_missing.empty:
        return None
    if ptypes.is_datetime64_any_dtype(normalised.dtype):
        return pd.to_datetime(normalised, errors="coerce")
    if require_hint:
        text = non_missing.astype(str)
        if not bool(text.map(lambda value: bool(_DATE_HINT.search(value))).any()):
            return None
    try:
        parsed = pd.to_datetime(normalised, errors="coerce", format="mixed")
    except (TypeError, ValueError):
        parsed = pd.to_datetime(normalised, errors="coerce")
    if not bool(parsed.notna().any()):
        return None
    return parsed


def _value_kind(value: Any) -> str:
    if _is_missing(value):
        return "missing"
    if isinstance(value, (bool, np.bool_)):
        return "boolean"
    if isinstance(value, (int, float, complex, np.number)):
        return "number"
    if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
        return "datetime"
    if isinstance(value, str):
        return "text"
    return "other"


def _looks_like_identifier_values(series: pd.Series) -> bool:
    normalised = _normalised_for_analysis(series)
    non_missing = normalised.dropna()
    if non_missing.empty:
        return False
    if ptypes.is_integer_dtype(normalised.dtype):
        return bool((non_missing.abs() > 9_007_199_254_740_991).any())
    text = non_missing.astype(str).str.strip()
    integer_like = text.map(lambda value: bool(_INTEGER_TEXT.fullmatch(value)))
    if not bool(integer_like.all()):
        return False
    return bool(
        text.map(lambda value: bool(_LEADING_ZERO_INTEGER.fullmatch(value))).any()
        or (text.str.lstrip("+-").str.len() > 15).any()
    )


def _semantic_type(series: pd.Series, column_name: str) -> str:
    if _IDENTIFIER_NAME.search(str(column_name)) or _looks_like_identifier_values(series):
        return "identifier"
    if ptypes.is_bool_dtype(series.dtype):
        return "boolean"
    if ptypes.is_datetime64_any_dtype(series.dtype):
        return "datetime"
    if ptypes.is_integer_dtype(series.dtype):
        return "integer"
    if ptypes.is_numeric_dtype(series.dtype):
        return "number"
    numeric = _safe_numeric_series(series, column_name)
    if numeric is not None:
        return "numeric_text"
    dated = _safe_datetime_series(series)
    if dated is not None and float(dated.notna().mean()) >= 0.8:
        return "datetime_text"
    non_missing = series.loc[~_missing_mask(series)]
    if non_missing.empty:
        return "empty"
    try:
        unique_count = int(non_missing.nunique(dropna=True))
    except TypeError:
        return "mixed"
    if unique_count <= max(20, int(len(non_missing) * 0.1)):
        return "category"
    return "text"


@dataclass(frozen=True)
class QualityIssue:
    """One actionable data-quality finding."""

    code: str
    severity: Literal["high", "medium", "low", "info"]
    message: str
    suggestion: str
    count: int
    ratio: float
    column: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "severity": self.severity,
            "column": self.column,
            "count": self.count,
            "ratio": self.ratio,
            "message": self.message,
            "suggestion": self.suggestion,
        }


@dataclass(frozen=True)
class DataQualityReport:
    """Overall quality score, component metrics, and issue list."""

    score: float
    grade: str
    row_count: int
    column_count: int
    issues: tuple[QualityIssue, ...]
    metrics: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "issues", tuple(self.issues))
        object.__setattr__(self, "metrics", _frozen_mapping(self.metrics))

    def to_dict(self) -> dict[str, Any]:
        return {
            "score": self.score,
            "grade": self.grade,
            "row_count": self.row_count,
            "column_count": self.column_count,
            "metrics": _json_value(dict(self.metrics)),
            "issues": [issue.to_dict() for issue in self.issues],
        }


@dataclass(frozen=True)
class OutlierDetectionResult:
    """Long-form outliers, per-column summary, and affected source rows."""

    method: Literal["iqr", "zscore"]
    outliers: pd.DataFrame
    summary: pd.DataFrame
    flagged_rows: pd.DataFrame

    def to_dict(self) -> dict[str, Any]:
        return {
            "method": self.method,
            "outliers": _records(self.outliers),
            "summary": _records(self.summary),
            "flagged_rows": _records(self.flagged_rows),
        }


@dataclass(frozen=True)
class TrendAnalysisResult:
    """Period aggregation plus rows excluded because of invalid dates/values."""

    data: pd.DataFrame
    frequency: str
    input_rows: int
    used_rows: int
    invalid_date_count: int
    invalid_value_counts: Mapping[str, int]

    def __post_init__(self) -> None:
        object.__setattr__(
            self, "invalid_value_counts", _frozen_mapping(self.invalid_value_counts)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "frequency": self.frequency,
            "input_rows": self.input_rows,
            "used_rows": self.used_rows,
            "invalid_date_count": self.invalid_date_count,
            "invalid_value_counts": dict(self.invalid_value_counts),
            "data": _records(self.data),
        }


@dataclass(frozen=True)
class CategoryContributionResult:
    """Sorted contribution table and Pareto metadata."""

    data: pd.DataFrame
    total: float
    pareto_threshold: float
    core_category_count: int
    input_rows: int
    used_rows: int
    invalid_value_count: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "total": self.total,
            "pareto_threshold": self.pareto_threshold,
            "core_category_count": self.core_category_count,
            "input_rows": self.input_rows,
            "used_rows": self.used_rows,
            "invalid_value_count": self.invalid_value_count,
            "data": _records(self.data),
        }


@dataclass(frozen=True)
class TableComparisonResult:
    """Added, removed, modified, unchanged, duplicate, and invalid key rows."""

    added: pd.DataFrame
    removed: pd.DataFrame
    modified: pd.DataFrame
    unchanged: pd.DataFrame
    duplicate_keys_old: pd.DataFrame
    duplicate_keys_new: pd.DataFrame
    invalid_keys_old: pd.DataFrame
    invalid_keys_new: pd.DataFrame
    summary: Mapping[str, Any]

    def __post_init__(self) -> None:
        object.__setattr__(self, "summary", _frozen_mapping(self.summary))

    def to_dict(self) -> dict[str, Any]:
        return {
            "summary": _json_value(dict(self.summary)),
            "added": _records(self.added),
            "removed": _records(self.removed),
            "modified": _records(self.modified),
            "unchanged": _records(self.unchanged),
            "duplicate_keys_old": _records(self.duplicate_keys_old),
            "duplicate_keys_new": _records(self.duplicate_keys_new),
            "invalid_keys_old": _records(self.invalid_keys_old),
            "invalid_keys_new": _records(self.invalid_keys_new),
        }


@dataclass(frozen=True)
class RFMAnalysisResult:
    """Customer-level RFM scores, segment summary, and excluded input rows."""

    customers: pd.DataFrame
    segment_summary: pd.DataFrame
    invalid_rows: pd.DataFrame
    reference_date: pd.Timestamp | None

    def to_dict(self) -> dict[str, Any]:
        return {
            "reference_date": _json_value(self.reference_date),
            "customers": _records(self.customers),
            "segment_summary": _records(self.segment_summary),
            "invalid_rows": _records(self.invalid_rows),
        }


def assess_data_quality(
    df: pd.DataFrame,
    *,
    key_columns: str | Sequence[str] | None = None,
) -> DataQualityReport:
    """Score completeness, uniqueness, key integrity, and structural quality.

    Blank strings count as missing. The returned score is transparent and
    deterministic: it starts at 100 and subtracts weighted issue rates.
    ``key_columns`` enables missing/duplicate business-key checks.
    """

    _require_dataframe(df, allow_duplicate_columns=True)
    row_count, column_count = df.shape
    issues: list[QualityIssue] = []
    requested_keys = (
        []
        if key_columns is None
        else _column_list(key_columns, argument="key_columns")
    )
    if requested_keys:
        _validate_columns(df, requested_keys, argument="key_columns")

    if row_count == 0 or column_count == 0:
        issues.append(
            QualityIssue(
                code="EMPTY_DATASET",
                severity="high",
                message="数据表没有可分析的数据行或字段。",
                suggestion="检查表头行、工作表选择和导入范围。",
                count=0,
                ratio=1.0,
            )
        )
        return DataQualityReport(
            score=0.0,
            grade="较差",
            row_count=row_count,
            column_count=column_count,
            issues=tuple(issues),
            metrics={
                "completeness_rate": 0.0,
                "duplicate_row_rate": 0.0,
                "key_issue_rate": 0.0,
            },
        )

    duplicate_column_labels = [
        str(value) for value in df.columns[df.columns.duplicated()].unique()
    ]
    if duplicate_column_labels:
        issues.append(
            QualityIssue(
                code="DUPLICATE_COLUMN_NAMES",
                severity="high",
                message=f"发现重复列名：{duplicate_column_labels}",
                suggestion="先为重复字段设置唯一列名，再进行分析或合并。",
                count=len(duplicate_column_labels),
                ratio=len(duplicate_column_labels) / column_count,
            )
        )

    missing_total = 0
    empty_columns = 0
    constant_columns = 0
    mixed_columns = 0
    for position in range(column_count):
        series = df.iloc[:, position]
        column_name = str(df.columns[position])
        missing_mask = _missing_mask(series)
        missing_count = int(missing_mask.sum())
        missing_total += missing_count
        missing_ratio = missing_count / row_count
        if missing_count:
            severity: Literal["high", "medium", "low", "info"]
            severity = "high" if missing_ratio >= 0.5 else "medium" if missing_ratio >= 0.1 else "low"
            issues.append(
                QualityIssue(
                    code="MISSING_VALUES",
                    severity=severity,
                    column=column_name,
                    message=f"字段“{column_name}”有 {missing_count} 个空值。",
                    suggestion="确认空值是否允许；必要时填充、删除或单独输出异常记录。",
                    count=missing_count,
                    ratio=missing_ratio,
                )
            )
        non_missing = series.loc[~missing_mask]
        if non_missing.empty:
            empty_columns += 1
            issues.append(
                QualityIssue(
                    code="EMPTY_COLUMN",
                    severity="high",
                    column=column_name,
                    message=f"字段“{column_name}”完全为空。",
                    suggestion="删除全空字段或检查导出模板是否多带了空列。",
                    count=row_count,
                    ratio=1.0,
                )
            )
            continue
        try:
            unique_count = int(non_missing.nunique(dropna=True))
        except TypeError:
            unique_count = -1
        if unique_count == 1:
            constant_columns += 1
            issues.append(
                QualityIssue(
                    code="CONSTANT_COLUMN",
                    severity="info",
                    column=column_name,
                    message=f"字段“{column_name}”只有一个非空取值。",
                    suggestion="确认该字段是否仍有分析价值。",
                    count=len(non_missing),
                    ratio=len(non_missing) / row_count,
                )
            )
        kinds = {_value_kind(value) for value in non_missing}
        if len(kinds) > 1:
            mixed_columns += 1
            issues.append(
                QualityIssue(
                    code="MIXED_TYPES",
                    severity="medium",
                    column=column_name,
                    message=f"字段“{column_name}”混合了多种数据类型：{sorted(kinds)}。",
                    suggestion="核对异常值后统一为文本、数字或日期类型。",
                    count=len(non_missing),
                    ratio=len(non_missing) / row_count,
                )
            )

    normalised = df.copy(deep=True)
    for position in range(column_count):
        series = normalised.iloc[:, position]
        blanks = series.map(
            lambda value: isinstance(value, str) and not value.strip(),
            na_action="ignore",
        ).fillna(False)
        if bool(blanks.any()):
            normalised.isetitem(position, series.mask(blanks, pd.NA))
    try:
        duplicate_rows = int(normalised.duplicated(keep=False).sum())
    except TypeError:
        duplicate_rows = 0
        issues.append(
            QualityIssue(
                code="UNHASHABLE_VALUES",
                severity="info",
                message="部分单元格包含列表或字典，无法执行整行重复检查。",
                suggestion="将嵌套对象转成稳定文本后再检查重复。",
                count=0,
                ratio=0.0,
            )
        )
    if duplicate_rows:
        issues.append(
            QualityIssue(
                code="DUPLICATE_ROWS",
                severity="medium",
                message=f"发现 {duplicate_rows} 行属于重复记录。",
                suggestion="选择业务主键并确认保留首条还是末条。",
                count=duplicate_rows,
                ratio=duplicate_rows / row_count,
            )
        )

    key_issue_rows = 0
    keys: list[str] = []
    if requested_keys:
        keys = requested_keys
        if duplicate_column_labels:
            ambiguous = [key for key in keys if list(df.columns).count(key) > 1]
            if ambiguous:
                raise ValueError(f"key_columns 引用了重复列名：{ambiguous}")
        key_frame = df.loc[:, keys].copy(deep=True)
        key_missing = pd.Series(False, index=df.index)
        for key in keys:
            key_missing |= _missing_mask(key_frame[key])
        key_duplicates = key_frame.duplicated(keep=False) & ~key_missing
        key_issue_mask = key_missing | key_duplicates
        key_issue_rows = int(key_issue_mask.sum())
        if bool(key_missing.any()):
            count = int(key_missing.sum())
            issues.append(
                QualityIssue(
                    code="MISSING_KEYS",
                    severity="high",
                    message=f"业务键 {keys} 有 {count} 行为空。",
                    suggestion="补齐业务键，或把这些行单独输出复核。",
                    count=count,
                    ratio=count / row_count,
                )
            )
        if bool(key_duplicates.any()):
            count = int(key_duplicates.sum())
            issues.append(
                QualityIssue(
                    code="DUPLICATE_KEYS",
                    severity="high",
                    message=f"业务键 {keys} 有 {count} 行处于重复键组。",
                    suggestion="连接或比对前先明确一对一、一对多关系。",
                    count=count,
                    ratio=count / row_count,
                )
            )

    missing_rate = missing_total / (row_count * column_count)
    duplicate_rate = duplicate_rows / row_count
    empty_column_rate = empty_columns / column_count
    constant_column_rate = constant_columns / column_count
    mixed_column_rate = mixed_columns / column_count
    key_issue_rate = key_issue_rows / row_count
    penalty = (
        missing_rate * 35
        + duplicate_rate * 20
        + empty_column_rate * 15
        + constant_column_rate * 5
        + mixed_column_rate * 10
        + key_issue_rate * 15
        + (10 if duplicate_column_labels else 0)
    )
    score = round(max(0.0, min(100.0, 100.0 - penalty)), 1)
    grade = "优秀" if score >= 90 else "良好" if score >= 75 else "需关注" if score >= 60 else "较差"
    issues.sort(key=lambda issue: (_SEVERITY_ORDER[issue.severity], issue.column or "", issue.code))
    metrics = {
        "completeness_rate": round(1 - missing_rate, 4),
        "duplicate_row_rate": round(duplicate_rate, 4),
        "empty_column_rate": round(empty_column_rate, 4),
        "constant_column_rate": round(constant_column_rate, 4),
        "mixed_type_column_rate": round(mixed_column_rate, 4),
        "key_issue_rate": round(key_issue_rate, 4),
        "missing_cell_count": missing_total,
        "duplicate_row_count": duplicate_rows,
    }
    return DataQualityReport(
        score=score,
        grade=grade,
        row_count=row_count,
        column_count=column_count,
        issues=tuple(issues),
        metrics=metrics,
    )


def descriptive_statistics(
    df: pd.DataFrame,
    *,
    columns: Sequence[str] | None = None,
    include_text: bool = True,
    percentiles: Sequence[float] = (0.25, 0.5, 0.75),
) -> pd.DataFrame:
    """Return one UI-friendly descriptive-statistics row per selected column."""

    _require_dataframe(df)
    selected = (
        list(df.columns)
        if columns is None
        else _column_list(columns, argument="columns")
    )
    _validate_columns(df, selected, argument="columns")
    percentile_values = tuple(float(value) for value in percentiles)
    if any(not 0 <= value <= 1 for value in percentile_values):
        raise ValueError("percentiles 必须在 [0, 1] 范围内")

    rows: list[dict[str, Any]] = []
    for column in selected:
        series = _normalised_for_analysis(df[column])
        numeric = _safe_numeric_series(series, str(column))
        dated = None if numeric is not None else _safe_datetime_series(series)
        if not include_text and numeric is None and dated is None:
            continue
        non_missing = series.dropna()
        try:
            unique_count = int(non_missing.nunique(dropna=True))
            modes = non_missing.mode(dropna=True)
        except TypeError:
            unique_count = -1
            modes = pd.Series(dtype=object)
        row: dict[str, Any] = {
            "column": str(column),
            "dtype": str(df[column].dtype),
            "semantic_type": _semantic_type(series, str(column)),
            "count": int(series.notna().sum()),
            "missing_count": int(series.isna().sum()),
            "missing_percent": round(float(series.isna().mean() * 100), 2) if len(series) else 0.0,
            "unique_count": unique_count,
            "mean": None,
            "std": None,
            "min": None,
            "max": None,
            "sum": None,
            "mode": modes.iloc[0] if not modes.empty else None,
            "mode_count": int((non_missing == modes.iloc[0]).sum()) if not modes.empty else 0,
        }
        for percentile in percentile_values:
            row[f"p{int(round(percentile * 100)):02d}"] = None
        if numeric is not None:
            valid = numeric.dropna()
            if not valid.empty:
                row.update(
                    {
                        "mean": float(valid.mean()),
                        "std": float(valid.std(ddof=1)) if len(valid) > 1 else 0.0,
                        "min": valid.min(),
                        "max": valid.max(),
                        "sum": valid.sum(),
                    }
                )
                for percentile in percentile_values:
                    row[f"p{int(round(percentile * 100)):02d}"] = valid.quantile(percentile)
        elif dated is not None:
            valid_dates = dated.dropna()
            if not valid_dates.empty:
                row["min"] = valid_dates.min()
                row["max"] = valid_dates.max()
        rows.append(row)
    return pd.DataFrame(rows)


def _inversion_count(values: list[float]) -> int:
    if len(values) < 2:
        return 0
    middle = len(values) // 2
    left, right = values[:middle], values[middle:]
    count = _inversion_count(left) + _inversion_count(right)
    left_index = right_index = output_index = 0
    while left_index < len(left) and right_index < len(right):
        if left[left_index] <= right[right_index]:
            values[output_index] = left[left_index]
            left_index += 1
        else:
            values[output_index] = right[right_index]
            right_index += 1
            count += len(left) - left_index
        output_index += 1
    while left_index < len(left):
        values[output_index] = left[left_index]
        left_index += 1
        output_index += 1
    while right_index < len(right):
        values[output_index] = right[right_index]
        right_index += 1
        output_index += 1
    return count


def _tie_pairs(values: Sequence[Any]) -> int:
    return sum(count * (count - 1) // 2 for count in Counter(values).values())


def _kendall_tau_b(left: pd.Series, right: pd.Series, min_periods: int) -> float:
    valid = left.notna() & right.notna()
    x = left.loc[valid].astype(float).tolist()
    y = right.loc[valid].astype(float).tolist()
    size = len(x)
    if size < max(2, min_periods):
        return float("nan")
    pairs = sorted(zip(x, y), key=lambda item: (item[0], item[1]))
    discordant = _inversion_count([value for _, value in pairs])
    total_pairs = size * (size - 1) // 2
    ties_x = _tie_pairs(x)
    ties_y = _tie_pairs(y)
    ties_both = _tie_pairs(pairs)
    comparable = total_pairs - ties_x - ties_y + ties_both
    numerator = comparable - 2 * discordant
    denominator = math.sqrt((total_pairs - ties_x) * (total_pairs - ties_y))
    return numerator / denominator if denominator else float("nan")


def correlation_matrix(
    df: pd.DataFrame,
    *,
    columns: Sequence[str] | None = None,
    method: Literal["pearson", "spearman", "kendall"] = "pearson",
    min_periods: int = 2,
) -> pd.DataFrame:
    """Return a correlation matrix for safe numeric and numeric-text columns."""

    _require_dataframe(df)
    if method not in {"pearson", "spearman", "kendall"}:
        raise ValueError("method 必须是 pearson、spearman 或 kendall")
    if min_periods < 1:
        raise ValueError("min_periods 必须是正整数")
    selected = (
        list(df.columns)
        if columns is None
        else _column_list(columns, argument="columns")
    )
    _validate_columns(df, selected, argument="columns")
    numeric_columns: dict[str, pd.Series] = {}
    for column in selected:
        numeric = _safe_numeric_series(df[column], str(column))
        if numeric is not None:
            numeric_columns[str(column)] = numeric
    if not numeric_columns:
        return pd.DataFrame(dtype=float)
    numeric_frame = pd.DataFrame(numeric_columns, index=df.index)
    if method == "kendall":
        names = list(numeric_frame.columns)
        matrix = pd.DataFrame(np.nan, index=names, columns=names, dtype=float)
        for row_index, left_name in enumerate(names):
            for column_index in range(row_index, len(names)):
                right_name = names[column_index]
                coefficient = _kendall_tau_b(
                    numeric_frame[left_name], numeric_frame[right_name], min_periods
                )
                matrix.loc[left_name, right_name] = coefficient
                matrix.loc[right_name, left_name] = coefficient
        return matrix
    return numeric_frame.corr(method=method, min_periods=min_periods)


def detect_outliers(
    df: pd.DataFrame,
    *,
    columns: Sequence[str] | None = None,
    method: Literal["iqr", "zscore"] = "iqr",
    iqr_multiplier: float = 1.5,
    z_threshold: float = 3.0,
) -> OutlierDetectionResult:
    """Detect numeric outliers with IQR fences or absolute z-scores."""

    _require_dataframe(df)
    if method not in {"iqr", "zscore"}:
        raise ValueError("method 必须是 iqr 或 zscore")
    if iqr_multiplier <= 0 or z_threshold <= 0:
        raise ValueError("iqr_multiplier 和 z_threshold 必须大于 0")
    selected = (
        list(df.columns)
        if columns is None
        else _column_list(columns, argument="columns")
    )
    _validate_columns(df, selected, argument="columns")
    outlier_rows: list[dict[str, Any]] = []
    summary_rows: list[dict[str, Any]] = []
    positions_by_column: dict[int, list[str]] = {}

    for column in selected:
        numeric = _safe_numeric_series(df[column], str(column))
        if numeric is None:
            continue
        valid = numeric.dropna()
        if valid.empty:
            continue
        if method == "iqr":
            q1 = float(valid.quantile(0.25))
            q3 = float(valid.quantile(0.75))
            spread = q3 - q1
            lower = q1 - iqr_multiplier * spread
            upper = q3 + iqr_multiplier * spread
            flags = numeric.notna() & ((numeric < lower) | (numeric > upper))
            center = float(valid.median())
            scale = spread
        else:
            center = float(valid.mean())
            scale = float(valid.std(ddof=0))
            if scale == 0 or math.isnan(scale):
                flags = pd.Series(False, index=df.index)
            else:
                flags = numeric.notna() & ((numeric - center).abs() / scale > z_threshold)
            lower = center - z_threshold * scale
            upper = center + z_threshold * scale

        flagged_positions = np.flatnonzero(flags.to_numpy()).tolist()
        summary_rows.append(
            {
                "column": str(column),
                "method": method,
                "valid_count": int(numeric.notna().sum()),
                "outlier_count": len(flagged_positions),
                "outlier_percent": round(len(flagged_positions) / len(df) * 100, 2) if len(df) else 0.0,
                "center": center,
                "scale": scale,
                "lower_bound": lower,
                "upper_bound": upper,
            }
        )
        for position in flagged_positions:
            value = numeric.iloc[position]
            if scale and not math.isnan(scale):
                score = abs(float(value) - center) / scale
            else:
                score = None
            outlier_rows.append(
                {
                    "row_position": position,
                    "row_index": df.index[position],
                    "column": str(column),
                    "value": value,
                    "method": method,
                    "score": score,
                    "lower_bound": lower,
                    "upper_bound": upper,
                }
            )
            positions_by_column.setdefault(position, []).append(str(column))

    outlier_columns = [
        "row_position",
        "row_index",
        "column",
        "value",
        "method",
        "score",
        "lower_bound",
        "upper_bound",
    ]
    summary_columns = [
        "column",
        "method",
        "valid_count",
        "outlier_count",
        "outlier_percent",
        "center",
        "scale",
        "lower_bound",
        "upper_bound",
    ]
    outliers = pd.DataFrame(outlier_rows, columns=outlier_columns)
    summary = pd.DataFrame(summary_rows, columns=summary_columns)
    positions = sorted(positions_by_column)
    flagged = df.iloc[positions].copy(deep=True).reset_index(drop=True)
    for reserved in ("row_position", "outlier_columns"):
        if reserved in flagged.columns:
            replacement = f"{reserved}_source"
            while replacement in flagged.columns:
                replacement += "_"
            flagged = flagged.rename(columns={reserved: replacement})
    flagged.insert(0, "row_position", positions)
    flagged["outlier_columns"] = ["；".join(positions_by_column[pos]) for pos in positions]
    return OutlierDetectionResult(
        method=method,
        outliers=outliers,
        summary=summary,
        flagged_rows=flagged,
    )


_FREQUENCY_ALIASES = {
    "day": ("day", "D"),
    "daily": ("day", "D"),
    "日": ("day", "D"),
    "week": ("week", "W"),
    "weekly": ("week", "W"),
    "周": ("week", "W"),
    "month": ("month", "M"),
    "monthly": ("month", "M"),
    "月": ("month", "M"),
    "quarter": ("quarter", "Q"),
    "quarterly": ("quarter", "Q"),
    "季度": ("quarter", "Q"),
    "year": ("year", "Y"),
    "yearly": ("year", "Y"),
    "年": ("year", "Y"),
}
_AGGREGATIONS = {"sum", "mean", "median", "min", "max", "count", "nunique"}


def aggregate_trend(
    df: pd.DataFrame,
    *,
    date_column: str,
    value_columns: str | Sequence[str],
    frequency: str = "month",
    aggregation: str | Mapping[str, str] = "sum",
    group_by: str | Sequence[str] | None = None,
    period_column: str = "period",
) -> TrendAnalysisResult:
    """Aggregate one or more metrics by day/week/month/quarter/year."""

    _require_dataframe(df)
    values = _column_list(value_columns, argument="value_columns")
    groups = [] if group_by is None else _column_list(group_by, argument="group_by")
    selected_roles = [date_column, *values, *groups]
    if len(set(selected_roles)) != len(selected_roles):
        raise ValueError("date_column、value_columns 和 group_by 不能重复使用同一字段")
    _validate_columns(df, [date_column, *values, *groups], argument="columns")
    if period_column in [*values, *groups]:
        raise ValueError("period_column 不能与指标列或分组列重名")
    frequency_key = str(frequency).casefold()
    if frequency_key not in _FREQUENCY_ALIASES:
        raise ValueError("frequency 必须是 day/week/month/quarter/year（或日/周/月/季度/年）")
    canonical_frequency, period_code = _FREQUENCY_ALIASES[frequency_key]
    if isinstance(aggregation, Mapping):
        aggregation_map = dict(aggregation)
        missing_agg = [column for column in values if column not in aggregation_map]
        if missing_agg:
            raise ValueError(f"aggregation 缺少指标配置：{missing_agg}")
        extra_agg = [column for column in aggregation_map if column not in values]
        if extra_agg:
            raise ValueError(f"aggregation 包含未选择的指标：{extra_agg}")
    else:
        aggregation_map = {column: str(aggregation) for column in values}
    invalid_aggs = {value for value in aggregation_map.values() if value not in _AGGREGATIONS}
    if invalid_aggs:
        raise ValueError(f"不支持的聚合方式：{sorted(invalid_aggs)}")
    output_columns = [period_column, *groups, *values]
    if df.empty:
        return TrendAnalysisResult(
            data=pd.DataFrame(columns=output_columns),
            frequency=canonical_frequency,
            input_rows=0,
            used_rows=0,
            invalid_date_count=0,
            invalid_value_counts={str(column): 0 for column in values},
        )

    dates = _safe_datetime_series(df[date_column], require_hint=False)
    if dates is None:
        dates = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    if getattr(dates.dt, "tz", None) is not None:
        dates = dates.dt.tz_localize(None)
    raw_date_missing = _missing_mask(df[date_column])
    invalid_date_count = int((~raw_date_missing & dates.isna()).sum())

    work = df.loc[:, groups].copy(deep=True) if groups else pd.DataFrame(index=df.index)
    numeric_values: dict[str, pd.Series] = {}
    invalid_value_counts: dict[str, int] = {}
    for column in values:
        numeric = _safe_numeric_series(
            df[column], str(column), minimum_success_ratio=0.0
        )
        if numeric is None:
            raise ValueError(
                f"指标列 {column!r} 不能安全解析为数字；标识符、前导零和超长整数不会被转换"
            )
        numeric_values[column] = numeric
        raw_missing = _missing_mask(df[column])
        invalid_value_counts[str(column)] = int((~raw_missing & numeric.isna()).sum())
        work[column] = numeric
    period_values = dates.dt.to_period(period_code).dt.start_time
    work[period_column] = period_values
    has_metric = pd.DataFrame(numeric_values).notna().any(axis=1)
    valid_rows = dates.notna() & has_metric
    work = work.loc[valid_rows].copy()

    if work.empty:
        aggregated = pd.DataFrame(columns=output_columns)
    else:
        aggregated = (
            work.groupby(
                [period_column, *groups],
                dropna=False,
                sort=True,
                observed=True,
                as_index=False,
            )
            .agg(aggregation_map)
            .sort_values([period_column, *groups], kind="stable")
            .reset_index(drop=True)
        )
    return TrendAnalysisResult(
        data=aggregated,
        frequency=canonical_frequency,
        input_rows=len(df),
        used_rows=int(valid_rows.sum()),
        invalid_date_count=invalid_date_count,
        invalid_value_counts=invalid_value_counts,
    )


def category_contribution(
    df: pd.DataFrame,
    *,
    category_columns: str | Sequence[str],
    value_column: str,
    aggregation: Literal["sum", "mean", "count", "nunique"] = "sum",
    pareto_threshold: float = 0.8,
    top_n: int | None = None,
    include_other: bool = False,
) -> CategoryContributionResult:
    """Rank category contribution and mark the Pareto core versus long tail."""

    _require_dataframe(df)
    categories = _column_list(category_columns, argument="category_columns")
    if value_column in categories:
        raise ValueError("value_column 不能同时作为 category_columns")
    _validate_columns(df, [*categories, value_column], argument="columns")
    if aggregation not in {"sum", "mean", "count", "nunique"}:
        raise ValueError("aggregation 必须是 sum、mean、count 或 nunique")
    if not 0 < pareto_threshold <= 1:
        raise ValueError("pareto_threshold 必须在 (0, 1] 范围内")
    if top_n is not None and top_n <= 0:
        raise ValueError("top_n 必须是正整数")
    if include_other and aggregation not in {"sum", "count"}:
        raise ValueError("include_other 目前仅支持 sum 或 count 聚合")
    numeric = _safe_numeric_series(
        df[value_column], value_column, minimum_success_ratio=0.0
    )
    if numeric is None:
        raise ValueError(f"value_column {value_column!r} 不能安全解析为数字")
    work = df.loc[:, categories].copy(deep=True)
    for column in categories:
        work[column] = work[column].map(
            lambda value: "（空值）" if _is_missing(value) else value
        )
    work["__value__"] = numeric
    raw_value_missing = _missing_mask(df[value_column])
    invalid_value_count = int((~raw_value_missing & numeric.isna()).sum())
    work = work.loc[work["__value__"].notna()]
    if work.empty:
        empty_columns = [
            *categories,
            "value",
            "contribution_pct",
            "cumulative_pct",
            "pareto_group",
        ]
        return CategoryContributionResult(
            data=pd.DataFrame(columns=empty_columns),
            total=0.0,
            pareto_threshold=pareto_threshold,
            core_category_count=0,
            input_rows=len(df),
            used_rows=0,
            invalid_value_count=invalid_value_count,
        )
    grouped = (
        work.groupby(categories, dropna=False, observed=True, sort=False)["__value__"]
        .agg(aggregation)
        .reset_index(name="value")
        .sort_values("value", ascending=False, kind="stable")
        .reset_index(drop=True)
    )
    if bool((grouped["value"] < 0).any()):
        raise ValueError("分类贡献/帕累托要求聚合值非负；请先处理退款或改用绝对值指标")
    total = float(grouped["value"].sum()) if not grouped.empty else 0.0
    if total <= 0:
        raise ValueError("分类贡献的汇总值必须大于 0")
    if top_n is not None and len(grouped) > top_n:
        head = grouped.iloc[:top_n].copy()
        if include_other:
            remainder = float(grouped.iloc[top_n:]["value"].sum())
            other = {column: "其他" if index == 0 else "" for index, column in enumerate(categories)}
            other["value"] = remainder
            grouped = pd.concat([head, pd.DataFrame([other])], ignore_index=True)
        else:
            grouped = head.reset_index(drop=True)
    grouped["contribution_pct"] = grouped["value"] / total
    grouped["cumulative_pct"] = grouped["contribution_pct"].cumsum()
    prior_cumulative = grouped["cumulative_pct"].shift(fill_value=0)
    grouped["pareto_group"] = np.where(
        prior_cumulative < pareto_threshold, "核心贡献", "长尾贡献"
    )
    core_count = int((grouped["pareto_group"] == "核心贡献").sum())
    return CategoryContributionResult(
        data=grouped.reset_index(drop=True),
        total=total,
        pareto_threshold=pareto_threshold,
        core_category_count=core_count,
        input_rows=len(df),
        used_rows=len(work),
        invalid_value_count=invalid_value_count,
    )


def _flatten_pivot_columns(columns: pd.Index) -> list[str]:
    flattened: list[str] = []
    for value in columns.to_flat_index():
        parts = value if isinstance(value, tuple) else (value,)
        flattened.append(" | ".join(str(part) for part in parts if str(part) not in {"", "None"}))
    return flattened


def _unique_output_columns(
    columns: Sequence[str], *, reserved: Sequence[str]
) -> list[str]:
    used = {str(value).casefold() for value in reserved}
    result: list[str] = []
    for raw_name in columns:
        base = str(raw_name) or "值"
        candidate = base
        counter = 2
        while candidate.casefold() in used:
            candidate = f"{base}_{counter}"
            counter += 1
        used.add(candidate.casefold())
        result.append(candidate)
    return result


def cross_pivot(
    df: pd.DataFrame,
    *,
    index: str | Sequence[str],
    columns: str | Sequence[str],
    values: str | None = None,
    aggregation: str = "count",
    fill_value: Any = 0,
    margins: bool = False,
    margins_name: str = "合计",
) -> pd.DataFrame:
    """Build a flat-column cross table for counts or a selected metric."""

    _require_dataframe(df)
    row_dimensions = _column_list(index, argument="index")
    column_dimensions = _column_list(columns, argument="columns")
    overlap = [column for column in row_dimensions if column in column_dimensions]
    if overlap:
        raise ValueError(f"index 和 columns 不能使用相同字段：{overlap}")
    if values is not None and values in [*row_dimensions, *column_dimensions]:
        raise ValueError("values 不能同时作为行维度或列维度")
    required = [*row_dimensions, *column_dimensions]
    if values is not None:
        required.append(values)
    _validate_columns(df, required, argument="columns")
    work = df.loc[:, required].copy(deep=True)
    for column in [*row_dimensions, *column_dimensions]:
        work[column] = work[column].map(
            lambda value: "（空值）" if _is_missing(value) else value
        )
    if values is None:
        if aggregation not in {"count", "size"}:
            raise ValueError("values=None 时 aggregation 必须是 count 或 size")
        pivot = pd.crosstab(
            index=[work[column] for column in row_dimensions],
            columns=[work[column] for column in column_dimensions],
            dropna=False,
            margins=margins,
            margins_name=margins_name,
        )
    else:
        if aggregation not in _AGGREGATIONS:
            raise ValueError(f"不支持的 aggregation：{aggregation}")
        if aggregation in {"sum", "mean", "median", "min", "max"}:
            numeric = _safe_numeric_series(
                work[values], values
            )
            if numeric is None:
                raise ValueError(f"values {values!r} 不能安全解析为数字")
            work[values] = numeric
        pivot = pd.pivot_table(
            work,
            index=row_dimensions,
            columns=column_dimensions,
            values=values,
            aggfunc=aggregation,
            fill_value=fill_value,
            observed=True,
            dropna=False,
            margins=margins,
            margins_name=margins_name,
        )
    pivot.columns = _unique_output_columns(
        _flatten_pivot_columns(pivot.columns), reserved=row_dimensions
    )
    return pivot.reset_index()


def _canonical_key_value(value: Any) -> str | None:
    if _is_missing(value):
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, (pd.Timestamp, datetime, date, np.datetime64)):
        return pd.Timestamp(value).isoformat()
    if isinstance(value, (bool, np.bool_)):
        return "True" if bool(value) else "False"
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        numeric = float(value)
        if numeric.is_integer():
            return str(int(numeric))
        return repr(numeric)
    return str(value)


def _key_tokens(
    frame: pd.DataFrame, key_columns: Sequence[str]
) -> tuple[list[tuple[str, ...] | None], list[int], list[int]]:
    tokens: list[tuple[str, ...] | None] = []
    key_values = [frame[column].tolist() for column in key_columns]
    for row_values in zip(*key_values):
        values = tuple(_canonical_key_value(value) for value in row_values)
        tokens.append(None if any(value is None for value in values) else values)  # type: ignore[arg-type]
    counts = Counter(token for token in tokens if token is not None)
    invalid_positions = [index for index, token in enumerate(tokens) if token is None]
    duplicate_positions = [
        index
        for index, token in enumerate(tokens)
        if token is not None and counts[token] > 1
    ]
    return tokens, invalid_positions, duplicate_positions


def _take_positions(frame: pd.DataFrame, positions: Sequence[int]) -> pd.DataFrame:
    return frame.iloc[list(positions)].copy(deep=True).reset_index(drop=True)


def _values_equal(old_value: Any, new_value: Any) -> bool:
    old_missing, new_missing = _is_missing(old_value), _is_missing(new_value)
    if old_missing or new_missing:
        return old_missing and new_missing
    if isinstance(old_value, (pd.Timestamp, datetime, date, np.datetime64)) or isinstance(
        new_value, (pd.Timestamp, datetime, date, np.datetime64)
    ):
        try:
            return pd.Timestamp(old_value) == pd.Timestamp(new_value)
        except (TypeError, ValueError):
            return False
    try:
        result = old_value == new_value
        return bool(result) if isinstance(result, (bool, np.bool_)) else False
    except (TypeError, ValueError):
        return False


def compare_tables(
    old: pd.DataFrame,
    new: pd.DataFrame,
    *,
    key_columns: str | Sequence[str],
    compare_columns: Sequence[str] | None = None,
    suffixes: tuple[str, str] = ("_旧", "_新"),
    include_unchanged: bool = True,
) -> TableComparisonResult:
    """Compare two tables into added/removed/modified/unchanged buckets.

    Duplicate-key and missing-key rows are returned separately and excluded
    from one-to-one comparison. Key canonicalisation matches numeric ``1`` with
    text ``"1"`` while preserving ``"001"`` and long identifiers exactly.
    """

    _require_dataframe(old, "old")
    _require_dataframe(new, "new")
    keys = _column_list(key_columns, argument="key_columns")
    _validate_columns(old, keys, argument="key_columns(old)")
    _validate_columns(new, keys, argument="key_columns(new)")
    if len(suffixes) != 2 or suffixes[0] == suffixes[1]:
        raise ValueError("suffixes 必须包含两个不同后缀")
    if compare_columns is None:
        selected = [
            column
            for column in old.columns
            if column not in keys and column in new.columns
        ]
    else:
        selected = _column_list(compare_columns, argument="compare_columns")
        _validate_columns(old, selected, argument="compare_columns(old)")
        _validate_columns(new, selected, argument="compare_columns(new)")
    overlapping_compare = [column for column in selected if column in keys]
    if overlapping_compare:
        raise ValueError(f"compare_columns 不应重复包含键字段：{overlapping_compare}")

    old_tokens, old_invalid_pos, old_duplicate_pos = _key_tokens(old, keys)
    new_tokens, new_invalid_pos, new_duplicate_pos = _key_tokens(new, keys)
    old_excluded = set(old_invalid_pos) | set(old_duplicate_pos)
    new_excluded = set(new_invalid_pos) | set(new_duplicate_pos)
    old_map = {
        token: position
        for position, token in enumerate(old_tokens)
        if token is not None and position not in old_excluded
    }
    new_map = {
        token: position
        for position, token in enumerate(new_tokens)
        if token is not None and position not in new_excluded
    }
    added_positions = [
        position for token, position in new_map.items() if token not in old_map
    ]
    removed_positions = [
        position for token, position in old_map.items() if token not in new_map
    ]
    modified_records: list[dict[str, Any]] = []
    unchanged_positions: list[int] = []
    modified_columns = [
        *keys,
        *[f"{column}{suffixes[0]}" for column in selected],
        *[f"{column}{suffixes[1]}" for column in selected],
        "changed_columns",
    ]
    if len(set(modified_columns)) != len(modified_columns):
        raise ValueError("比较结果列名会发生冲突；请调整 suffixes 或 compare_columns")
    for token, new_position in new_map.items():
        if token not in old_map:
            continue
        old_position = old_map[token]
        old_row, new_row = old.iloc[old_position], new.iloc[new_position]
        changed = [
            column
            for column in selected
            if not _values_equal(old_row[column], new_row[column])
        ]
        if not changed:
            unchanged_positions.append(new_position)
            continue
        record: dict[str, Any] = {key: new_row[key] for key in keys}
        for column in selected:
            record[f"{column}{suffixes[0]}"] = old_row[column]
            record[f"{column}{suffixes[1]}"] = new_row[column]
        record["changed_columns"] = "；".join(str(column) for column in changed)
        modified_records.append(record)

    modified = pd.DataFrame(modified_records, columns=modified_columns)
    unchanged = (
        _take_positions(new, unchanged_positions)
        if include_unchanged
        else pd.DataFrame(columns=new.columns)
    )
    summary = {
        "old_rows": len(old),
        "new_rows": len(new),
        "added_count": len(added_positions),
        "removed_count": len(removed_positions),
        "modified_count": len(modified_records),
        "unchanged_count": len(unchanged_positions),
        "duplicate_key_rows_old": len(old_duplicate_pos),
        "duplicate_key_rows_new": len(new_duplicate_pos),
        "invalid_key_rows_old": len(old_invalid_pos),
        "invalid_key_rows_new": len(new_invalid_pos),
        "key_columns": list(keys),
        "compare_columns": list(selected),
        "columns_only_old": [column for column in old.columns if column not in new.columns],
        "columns_only_new": [column for column in new.columns if column not in old.columns],
    }
    return TableComparisonResult(
        added=_take_positions(new, added_positions),
        removed=_take_positions(old, removed_positions),
        modified=modified,
        unchanged=unchanged,
        duplicate_keys_old=_take_positions(old, old_duplicate_pos),
        duplicate_keys_new=_take_positions(new, new_duplicate_pos),
        invalid_keys_old=_take_positions(old, old_invalid_pos),
        invalid_keys_new=_take_positions(new, new_invalid_pos),
        summary=summary,
    )


def _rfm_score(series: pd.Series, quantiles: int, *, higher_is_better: bool) -> pd.Series:
    if len(series) == 1:
        return pd.Series([quantiles], index=series.index, dtype="Int64")
    ranks = series.rank(
        method="average", ascending=higher_is_better, pct=True
    )
    scores = np.ceil(ranks * quantiles).clip(1, quantiles)
    return scores.astype("Int64")


def _rfm_segment(row: pd.Series, quantiles: int) -> str:
    high = max(2, math.ceil(quantiles * 0.8))
    middle = max(2, math.ceil(quantiles * 0.6))
    low = max(1, math.floor(quantiles * 0.4))
    r, f, m = int(row["r_score"]), int(row["f_score"]), int(row["m_score"])
    if r >= high and f >= high and m >= high:
        return "高价值客户"
    if r >= middle and f >= middle:
        return "忠诚客户"
    if r >= high and f <= low:
        return "新客/潜力客户"
    if r <= low and f >= middle:
        return "需唤回客户"
    if r <= low and f <= low:
        return "流失风险客户"
    return "一般客户"


def rfm_segmentation(
    df: pd.DataFrame,
    *,
    customer_column: str,
    date_column: str,
    amount_column: str,
    transaction_column: str | None = None,
    reference_date: str | date | datetime | pd.Timestamp | None = None,
    quantiles: int = 5,
) -> RFMAnalysisResult:
    """Calculate recency, frequency, monetary scores and customer segments."""

    _require_dataframe(df)
    required = [customer_column, date_column, amount_column]
    if transaction_column is not None:
        required.append(transaction_column)
    _validate_columns(df, required, argument="columns")
    if len(set(required)) != len(required):
        raise ValueError("客户、日期、金额和交易字段必须互不相同")
    if not 2 <= quantiles <= 10:
        raise ValueError("quantiles 必须在 2 到 10 之间")

    customer_output_columns = [
        customer_column,
        "last_purchase_date",
        "recency_days",
        "frequency",
        "monetary",
        "r_score",
        "f_score",
        "m_score",
        "rfm_score",
        "segment",
    ]
    if df.empty:
        invalid_rows = df.copy(deep=True)
        if "invalid_reason" in invalid_rows.columns:
            replacement = "invalid_reason_source"
            while replacement in invalid_rows.columns:
                replacement += "_"
            invalid_rows = invalid_rows.rename(columns={"invalid_reason": replacement})
        invalid_rows["invalid_reason"] = pd.Series(dtype=str)
        return RFMAnalysisResult(
            customers=pd.DataFrame(columns=customer_output_columns),
            segment_summary=pd.DataFrame(
                columns=["segment", "customer_count", "customer_pct", "monetary"]
            ),
            invalid_rows=invalid_rows,
            reference_date=None,
        )

    dates = _safe_datetime_series(df[date_column], require_hint=False)
    if dates is None:
        dates = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
    if getattr(dates.dt, "tz", None) is not None:
        dates = dates.dt.tz_localize(None)
    amounts = _safe_numeric_series(
        df[amount_column], amount_column, minimum_success_ratio=0.0
    )
    if amounts is None:
        raise ValueError(f"amount_column {amount_column!r} 不能安全解析为数字")
    customer_missing = _missing_mask(df[customer_column])
    date_missing = dates.isna()
    amount_missing = amounts.isna()
    valid = ~customer_missing & ~date_missing & ~amount_missing

    invalid_positions = np.flatnonzero((~valid).to_numpy()).tolist()
    invalid_rows = _take_positions(df, invalid_positions)
    if "invalid_reason" in invalid_rows.columns:
        replacement = "invalid_reason_source"
        while replacement in invalid_rows.columns:
            replacement += "_"
        invalid_rows = invalid_rows.rename(columns={"invalid_reason": replacement})
    if invalid_positions:
        reasons = []
        for position in invalid_positions:
            row_reasons = []
            if customer_missing.iloc[position]:
                row_reasons.append("客户标识为空")
            if date_missing.iloc[position]:
                row_reasons.append("日期为空或无法解析")
            if amount_missing.iloc[position]:
                row_reasons.append("金额为空或无法解析")
            reasons.append("；".join(row_reasons))
        invalid_rows["invalid_reason"] = reasons
    else:
        invalid_rows["invalid_reason"] = pd.Series(dtype=str)

    if not bool(valid.any()):
        return RFMAnalysisResult(
            customers=pd.DataFrame(columns=customer_output_columns),
            segment_summary=pd.DataFrame(
                columns=["segment", "customer_count", "customer_pct", "monetary"]
            ),
            invalid_rows=invalid_rows,
            reference_date=None,
        )

    work = pd.DataFrame(
        {
            customer_column: df.loc[valid, customer_column].to_numpy(),
            "__date__": dates.loc[valid].to_numpy(),
            "__amount__": amounts.loc[valid].to_numpy(),
        }
    )
    if transaction_column is not None:
        transactions = _normalised_for_analysis(df[transaction_column])
        work["__transaction__"] = transactions.loc[valid].to_numpy()
    grouped = work.groupby(customer_column, sort=False, observed=True, dropna=False)
    customer_data = grouped.agg(
        last_purchase_date=("__date__", "max"),
        monetary=("__amount__", "sum"),
    )
    if transaction_column is None:
        customer_data["frequency"] = grouped.size()
    else:
        customer_data["frequency"] = grouped["__transaction__"].nunique(dropna=True)
    customer_data = customer_data.reset_index()

    latest_purchase = pd.Timestamp(customer_data["last_purchase_date"].max())
    if reference_date is None:
        reference = latest_purchase.normalize() + pd.Timedelta(days=1)
    else:
        reference = pd.Timestamp(reference_date)
        if reference.tzinfo is not None:
            reference = reference.tz_localize(None)
        reference = reference.normalize()
        if reference < latest_purchase.normalize():
            raise ValueError("reference_date 不能早于最后一笔有效交易日期")
    customer_data["recency_days"] = (
        reference - customer_data["last_purchase_date"].dt.normalize()
    ).dt.days.astype("Int64")
    customer_data["r_score"] = _rfm_score(
        customer_data["recency_days"], quantiles, higher_is_better=False
    )
    customer_data["f_score"] = _rfm_score(
        customer_data["frequency"], quantiles, higher_is_better=True
    )
    customer_data["m_score"] = _rfm_score(
        customer_data["monetary"], quantiles, higher_is_better=True
    )
    customer_data["rfm_score"] = (
        customer_data["r_score"]
        + customer_data["f_score"]
        + customer_data["m_score"]
    ).astype("Int64")
    customer_data["segment"] = customer_data.apply(
        lambda row: _rfm_segment(row, quantiles), axis=1
    )
    customer_data = customer_data.loc[:, customer_output_columns].sort_values(
        ["rfm_score", "monetary"], ascending=[False, False], kind="stable"
    ).reset_index(drop=True)
    segment_summary = (
        customer_data.groupby("segment", observed=True, sort=False, as_index=False)
        .agg(customer_count=(customer_column, "size"), monetary=("monetary", "sum"))
        .sort_values("monetary", ascending=False, kind="stable")
        .reset_index(drop=True)
    )
    segment_summary["customer_pct"] = (
        segment_summary["customer_count"] / len(customer_data)
    )
    segment_summary = segment_summary.loc[
        :, ["segment", "customer_count", "customer_pct", "monetary"]
    ]
    return RFMAnalysisResult(
        customers=customer_data,
        segment_summary=segment_summary,
        invalid_rows=invalid_rows,
        reference_date=reference,
    )


__all__ = [
    "CategoryContributionResult",
    "DataQualityReport",
    "OutlierDetectionResult",
    "QualityIssue",
    "RFMAnalysisResult",
    "TableComparisonResult",
    "TrendAnalysisResult",
    "aggregate_trend",
    "assess_data_quality",
    "category_contribution",
    "compare_tables",
    "correlation_matrix",
    "cross_pivot",
    "descriptive_statistics",
    "detect_outliers",
    "rfm_segmentation",
]
