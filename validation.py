"""Declarative, non-mutating validation rules for pandas DataFrames."""

from __future__ import annotations

from collections import Counter
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import json
import math
import re
from types import MappingProxyType
from typing import Any, Literal

import numpy as np
import pandas as pd


RuleType = Literal[
    "not_null",
    "unique",
    "range",
    "regex",
    "allowed_values",
    "numeric",
    "date",
    "column_compare",
]
Severity = Literal["error", "warning", "info"]

RULE_TYPES: frozenset[str] = frozenset(
    {
        "not_null",
        "unique",
        "range",
        "regex",
        "allowed_values",
        "numeric",
        "date",
        "column_compare",
    }
)
SEVERITIES: frozenset[str] = frozenset({"error", "warning", "info"})
COLUMN_COMPARE_OPERATORS: frozenset[str] = frozenset(
    {"eq", "ne", "gt", "gte", "lt", "lte"}
)
FAILURE_COLUMNS = [
    "rule_id",
    "rule_type",
    "severity",
    "column",
    "other_column",
    "row_position",
    "code",
    "message",
    "value_preview",
    "other_value_preview",
]
MAX_REGEX_LENGTH = 256
MAX_VALUE_PREVIEW_CHARS = 500

_RULE_PARAM_KEYS: Mapping[str, frozenset[str]] = MappingProxyType(
    {
        "not_null": frozenset({"blank_as_null"}),
        "unique": frozenset({"ignore_nulls", "blank_as_null"}),
        "range": frozenset(
            {"min", "max", "inclusive_min", "inclusive_max", "ignore_nulls", "value_type"}
        ),
        "regex": frozenset({"pattern", "mode", "ignore_nulls"}),
        "allowed_values": frozenset({"values", "ignore_nulls"}),
        "numeric": frozenset({"ignore_nulls", "integer_only", "allow_infinite"}),
        "date": frozenset({"ignore_nulls", "format", "dayfirst", "min", "max"}),
        "column_compare": frozenset(
            {"other_column", "operator", "ignore_nulls", "value_type"}
        ),
    }
)


def _require_string(value: Any, *, name: str, allow_empty: bool = False) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{name} 必须是字符串")
    if not allow_empty and not value.strip():
        raise ValueError(f"{name} 不能为空")
    if any(ord(char) < 32 for char in value):
        raise ValueError(f"{name} 不能包含控制字符")
    return value


def _require_bool(value: Any, *, name: str) -> bool:
    if not isinstance(value, bool):
        raise TypeError(f"{name} 必须是布尔值")
    return value


def _is_scalar(value: Any) -> bool:
    return value is None or isinstance(value, (str, bool, int, float))


def _validate_scalar(value: Any, *, name: str) -> None:
    if not _is_scalar(value):
        raise TypeError(f"{name} 必须是 JSON 标量")
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError(f"{name} 不允许 NaN 或无穷大")


def _freeze_json(value: Any, *, path: str = "params", depth: int = 0) -> Any:
    if depth > 20:
        raise ValueError(f"{path} 嵌套过深")
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(f"{path} 不允许 NaN 或无穷大")
        return value
    if isinstance(value, Mapping):
        frozen: dict[str, Any] = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise TypeError(f"{path} 的 JSON 对象键必须是字符串")
            frozen[key] = _freeze_json(item, path=f"{path}.{key}", depth=depth + 1)
        return MappingProxyType(frozen)
    if isinstance(value, (list, tuple)):
        return tuple(
            _freeze_json(item, path=f"{path}[{index}]", depth=depth + 1)
            for index, item in enumerate(value)
        )
    raise TypeError(f"{path} 只能包含 JSON 值，不能包含 {type(value).__name__}")


def _thaw_json(value: Any) -> Any:
    if isinstance(value, Mapping):
        return {str(key): _thaw_json(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [_thaw_json(item) for item in value]
    return value


def _safe_regex(pattern: str) -> re.Pattern[str]:
    _require_string(pattern, name="regex.pattern", allow_empty=True)
    if len(pattern) > MAX_REGEX_LENGTH:
        raise ValueError(f"regex.pattern 最长 {MAX_REGEX_LENGTH} 个字符")
    # Python's stdlib regex engine has no timeout.  Reject the most common
    # constructs used for catastrophic backtracking and advanced assertions.
    if "(?" in pattern:
        raise ValueError("regex.pattern 不支持扩展分组或环视")
    if re.search(r"\\[1-9]", pattern):
        raise ValueError("regex.pattern 不支持反向引用")
    if re.search(r"\([^)]*[+*{][^)]*\)[+*{]", pattern):
        raise ValueError("regex.pattern 不支持嵌套重复量词")
    try:
        return re.compile(pattern)
    except re.error as exc:
        raise ValueError("regex.pattern 不是有效的正则表达式") from exc


def _validate_rule_params(rule_type: str, params: Mapping[str, Any]) -> None:
    unknown = sorted(set(params) - _RULE_PARAM_KEYS[rule_type])
    if unknown:
        raise ValueError(f"{rule_type} 包含不支持的参数：{unknown}")

    bool_defaults: dict[str, dict[str, bool]] = {
        "not_null": {"blank_as_null": True},
        "unique": {"ignore_nulls": True, "blank_as_null": False},
        "range": {"inclusive_min": True, "inclusive_max": True, "ignore_nulls": True},
        "regex": {"ignore_nulls": True},
        "allowed_values": {"ignore_nulls": True},
        "numeric": {"ignore_nulls": True, "integer_only": False, "allow_infinite": False},
        "date": {"ignore_nulls": True, "dayfirst": False},
        "column_compare": {"ignore_nulls": True},
    }
    for key, default in bool_defaults.get(rule_type, {}).items():
        _require_bool(params.get(key, default), name=f"{rule_type}.{key}")

    if rule_type == "range":
        if "min" not in params and "max" not in params:
            raise ValueError("range 至少需要 min 或 max")
        value_type = params.get("value_type", "numeric")
        if value_type not in {"numeric", "date"}:
            raise ValueError("range.value_type 必须是 numeric 或 date")
        parsed_bounds: dict[str, Decimal | pd.Timestamp] = {}
        for key in ("min", "max"):
            if key in params:
                _validate_scalar(params[key], name=f"range.{key}")
                if params[key] is None:
                    raise ValueError(f"range.{key} 不能是空值")
                parsed = (
                    _to_decimal(params[key])
                    if value_type == "numeric"
                    else _parse_date(params[key])
                )
                if parsed is None:
                    label = "数字" if value_type == "numeric" else "日期"
                    raise ValueError(f"range.{key} 不是有效{label}")
                parsed_bounds[key] = parsed
        if (
            "min" in parsed_bounds
            and "max" in parsed_bounds
            and parsed_bounds["min"] > parsed_bounds["max"]
        ):
            raise ValueError("range.min 不能大于 range.max")
        return

    if rule_type == "regex":
        if "pattern" not in params:
            raise ValueError("regex 缺少 pattern")
        _safe_regex(params["pattern"])
        if params.get("mode", "fullmatch") not in {"fullmatch", "match", "search"}:
            raise ValueError("regex.mode 必须是 fullmatch、match 或 search")
        return

    if rule_type == "allowed_values":
        values = params.get("values")
        if not isinstance(values, (list, tuple)) or not values:
            raise ValueError("allowed_values.values 必须是非空列表")
        for index, value in enumerate(values):
            _validate_scalar(value, name=f"allowed_values.values[{index}]")
        return

    if rule_type == "date":
        if "format" in params and params["format"] is not None:
            value = _require_string(params["format"], name="date.format")
            if len(value) > 100:
                raise ValueError("date.format 过长")
        parsed_bounds: dict[str, pd.Timestamp] = {}
        for key in ("min", "max"):
            if key in params:
                _validate_scalar(params[key], name=f"date.{key}")
                parsed = _parse_date(
                    params[key],
                    date_format=params.get("format"),
                    dayfirst=params.get("dayfirst", False),
                )
                if parsed is None:
                    raise ValueError(f"date.{key} 不是有效日期")
                parsed_bounds[key] = parsed
        if (
            "min" in parsed_bounds
            and "max" in parsed_bounds
            and parsed_bounds["min"] > parsed_bounds["max"]
        ):
            raise ValueError("date.min 不能大于 date.max")
        return

    if rule_type == "column_compare":
        if "other_column" not in params or "operator" not in params:
            raise ValueError("column_compare 必须包含 other_column 和 operator")
        _require_string(params["other_column"], name="column_compare.other_column")
        if params["operator"] not in COLUMN_COMPARE_OPERATORS:
            raise ValueError(f"不支持的列比较符：{params['operator']}")
        if params.get("value_type", "native") not in {"native", "numeric", "date"}:
            raise ValueError("column_compare.value_type 必须是 native、numeric 或 date")


@dataclass(frozen=True)
class ValidationRule:
    """One safe validation rule with a stable identifier and severity."""

    rule_id: str
    rule_type: RuleType | str
    column: str
    severity: Severity | str = "error"
    params: Mapping[str, Any] = field(default_factory=dict)
    message: str | None = None
    enabled: bool = True

    def __post_init__(self) -> None:
        _require_string(self.rule_id, name="rule_id")
        if len(self.rule_id) > 128:
            raise ValueError("rule_id 最长 128 个字符")
        rule_type = _require_string(self.rule_type, name="rule_type")
        if rule_type not in RULE_TYPES:
            raise ValueError(f"不支持的验证规则：{rule_type}")
        _require_string(self.column, name="column")
        severity = _require_string(self.severity, name="severity")
        if severity not in SEVERITIES:
            raise ValueError("severity 必须是 error、warning 或 info")
        if self.message is not None:
            _require_string(self.message, name="message")
            if len(self.message) > 500:
                raise ValueError("message 最长 500 个字符")
        _require_bool(self.enabled, name="enabled")
        frozen = _freeze_json(self.params)
        if not isinstance(frozen, Mapping):
            raise TypeError("params 必须是 JSON 对象")
        _validate_rule_params(rule_type, frozen)
        object.__setattr__(self, "rule_type", rule_type)
        object.__setattr__(self, "severity", severity)
        object.__setattr__(self, "params", frozen)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "rule_id": self.rule_id,
            "rule_type": self.rule_type,
            "column": self.column,
            "severity": self.severity,
            "params": _thaw_json(self.params),
            "enabled": self.enabled,
        }
        if self.message is not None:
            result["message"] = self.message
        return result

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "ValidationRule":
        if not isinstance(data, Mapping):
            raise TypeError("ValidationRule 必须是 JSON 对象")
        allowed = {"rule_id", "rule_type", "column", "severity", "params", "message", "enabled"}
        unknown = sorted(set(data) - allowed)
        if unknown:
            raise ValueError(f"ValidationRule 包含未知字段：{unknown}")
        missing = sorted({"rule_id", "rule_type", "column"} - set(data))
        if missing:
            raise ValueError(f"ValidationRule 缺少字段：{missing}")
        return cls(
            rule_id=data["rule_id"],
            rule_type=data["rule_type"],
            column=data["column"],
            severity=data.get("severity", "error"),
            params=data.get("params", {}),
            message=data.get("message"),
            enabled=data.get("enabled", True),
        )


@dataclass(frozen=True)
class RuleValidationResult:
    """Aggregated outcome of one validation rule."""

    rule_id: str
    rule_type: str
    severity: str
    column: str
    passed: bool
    checked_count: int
    failure_count: int
    failure_codes: Mapping[str, int] = field(default_factory=dict)
    skipped: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "failure_codes", MappingProxyType(dict(self.failure_codes)))

    def to_dict(self) -> dict[str, Any]:
        return {
            "rule_id": self.rule_id,
            "rule_type": self.rule_type,
            "severity": self.severity,
            "column": self.column,
            "passed": self.passed,
            "checked_count": self.checked_count,
            "failure_count": self.failure_count,
            "failure_codes": dict(self.failure_codes),
            "skipped": self.skipped,
        }


@dataclass(frozen=True)
class ValidationReport:
    """Validation overview plus a long-form, one-row-per-violation table."""

    row_count: int
    column_count: int
    rule_results: tuple[RuleValidationResult, ...]
    failures: pd.DataFrame
    blocking_failure_count: int
    severity_counts: Mapping[str, int]

    def __post_init__(self) -> None:
        object.__setattr__(self, "rule_results", tuple(self.rule_results))
        object.__setattr__(self, "failures", self.failures.copy(deep=True))
        object.__setattr__(self, "severity_counts", MappingProxyType(dict(self.severity_counts)))

    @property
    def rule_count(self) -> int:
        return len(self.rule_results)

    @property
    def passed_rule_count(self) -> int:
        return sum(result.passed for result in self.rule_results)

    @property
    def failed_rule_count(self) -> int:
        return sum(not result.passed and not result.skipped for result in self.rule_results)

    @property
    def failure_count(self) -> int:
        return len(self.failures)

    @property
    def passed(self) -> bool:
        """Warnings and informational failures do not block the report."""

        return self.blocking_failure_count == 0

    def failures_frame(self) -> pd.DataFrame:
        return self.failures.copy(deep=True)

    def rule_results_frame(self) -> pd.DataFrame:
        columns = [
            "rule_id",
            "rule_type",
            "severity",
            "column",
            "passed",
            "checked_count",
            "failure_count",
            "failure_codes",
            "skipped",
        ]
        rows = []
        for result in self.rule_results:
            row = result.to_dict()
            row["failure_codes"] = json.dumps(
                row["failure_codes"], ensure_ascii=False, sort_keys=True
            )
            rows.append(row)
        return pd.DataFrame(rows, columns=columns)

    def to_dict(self, *, include_failures: bool = False) -> dict[str, Any]:
        result: dict[str, Any] = {
            "passed": self.passed,
            "row_count": self.row_count,
            "column_count": self.column_count,
            "rule_count": self.rule_count,
            "passed_rule_count": self.passed_rule_count,
            "failed_rule_count": self.failed_rule_count,
            "failure_count": self.failure_count,
            "blocking_failure_count": self.blocking_failure_count,
            "severity_counts": dict(self.severity_counts),
            "rule_results": [item.to_dict() for item in self.rule_results],
        }
        if include_failures:
            result["failures"] = self.failures.astype(object).where(
                self.failures.notna(), None
            ).to_dict(orient="records")
        return result


def _is_missing(value: Any, *, blank_as_null: bool = False) -> bool:
    if blank_as_null and isinstance(value, str) and not value.strip():
        return True
    try:
        result = pd.isna(value)
        return isinstance(result, (bool, np.bool_)) and bool(result)
    except (TypeError, ValueError):
        return False


def _missing_mask(series: pd.Series, *, blank_as_null: bool = False) -> pd.Series:
    return series.map(lambda value: _is_missing(value, blank_as_null=blank_as_null)).astype(bool)


def _to_decimal(value: Any, *, allow_infinite: bool = False) -> Decimal | None:
    if isinstance(value, (bool, np.bool_)) or _is_missing(value):
        return None
    if not isinstance(value, (str, int, float, Decimal, np.integer, np.floating)):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        result = Decimal(text)
    except (InvalidOperation, ValueError):
        return None
    if not allow_infinite and not result.is_finite():
        return None
    return result


def _parse_date(
    value: Any,
    *,
    date_format: str | None = None,
    dayfirst: bool = False,
) -> pd.Timestamp | None:
    if _is_missing(value) or isinstance(value, (bool, int, float, np.number)):
        return None
    try:
        if isinstance(value, (pd.Timestamp, datetime, date)):
            parsed = pd.Timestamp(value)
        elif date_format:
            parsed = pd.Timestamp(datetime.strptime(str(value), date_format))
        else:
            parsed = pd.to_datetime(str(value), errors="coerce", dayfirst=dayfirst)
        if pd.isna(parsed):
            return None
        parsed = pd.Timestamp(parsed)
        if parsed.tzinfo is not None:
            parsed = parsed.tz_convert("UTC").tz_localize(None)
        return parsed
    except (TypeError, ValueError, OverflowError):
        return None


def _strict_key(value: Any) -> tuple[str, Any]:
    if _is_missing(value):
        return ("null", None)
    if isinstance(value, (bool, np.bool_)):
        return ("bool", bool(value))
    if isinstance(value, str):
        return ("str", value)
    if isinstance(value, (int, float, Decimal, np.integer, np.floating)):
        number = _to_decimal(value, allow_infinite=True)
        if number is not None and number.is_finite():
            number = Decimal(0) if number == 0 else number.normalize()
        return ("number", str(number) if number is not None else str(value))
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return ("date", pd.Timestamp(value).isoformat())
    return (f"object:{type(value).__qualname__}", str(value))


def _strict_equal(left: Any, right: Any) -> bool:
    return _strict_key(left) == _strict_key(right)


def _preview(value: Any, *, include_values: bool, max_chars: int) -> str | None:
    if not include_values:
        return None
    if _is_missing(value):
        return "<空值>"
    if isinstance(value, (pd.Timestamp, datetime, date)):
        text = pd.Timestamp(value).isoformat()
    else:
        text = str(value)
    text = text.replace("\r", "\\r").replace("\n", "\\n")
    return text if len(text) <= max_chars else f"{text[: max_chars - 1]}…"


_DEFAULT_MESSAGES: Mapping[str, str] = MappingProxyType(
    {
        "missing_column": "规则引用的列不存在",
        "null_value": "值不能为空",
        "duplicate_value": "值不唯一",
        "not_numeric": "值不是有效数字",
        "not_integer": "值不是整数",
        "out_of_range": "值超出允许范围",
        "regex_mismatch": "值不符合格式规则",
        "type_mismatch": "值类型不符合规则要求",
        "not_allowed": "值不在允许集合中",
        "not_date": "值不是有效日期",
        "date_out_of_range": "日期超出允许范围",
        "comparison_failed": "两列值不满足比较关系",
        "incomparable": "两列值无法按指定类型比较",
    }
)


def _failure_row(
    rule: ValidationRule,
    *,
    row_position: int | None,
    code: str,
    value: Any = None,
    other_column: str | None = None,
    other_value: Any = None,
    include_values: bool,
    max_chars: int,
) -> dict[str, Any]:
    return {
        "rule_id": rule.rule_id,
        "rule_type": rule.rule_type,
        "severity": rule.severity,
        "column": rule.column,
        "other_column": other_column,
        "row_position": row_position,
        "code": code,
        "message": rule.message or _DEFAULT_MESSAGES.get(code, "验证失败"),
        "value_preview": _preview(value, include_values=include_values, max_chars=max_chars),
        "other_value_preview": _preview(
            other_value, include_values=include_values, max_chars=max_chars
        ) if other_column is not None else None,
    }


def _range_failures(
    series: pd.Series,
    params: Mapping[str, Any],
) -> list[tuple[int, str]]:
    value_type = params.get("value_type", "numeric")
    ignore_nulls = params.get("ignore_nulls", True)
    minimum = params.get("min")
    maximum = params.get("max")
    if value_type == "numeric":
        parsed_min = _to_decimal(minimum) if minimum is not None else None
        parsed_max = _to_decimal(maximum) if maximum is not None else None
    else:
        parsed_min = _parse_date(minimum) if minimum is not None else None
        parsed_max = _parse_date(maximum) if maximum is not None else None
    failures: list[tuple[int, str]] = []
    for position, value in enumerate(series.array):
        if _is_missing(value):
            if not ignore_nulls:
                failures.append((position, "null_value"))
            continue
        parsed = _to_decimal(value) if value_type == "numeric" else _parse_date(value)
        if parsed is None:
            failures.append((position, "not_numeric" if value_type == "numeric" else "not_date"))
            continue
        too_low = parsed_min is not None and (
            parsed < parsed_min if params.get("inclusive_min", True) else parsed <= parsed_min
        )
        too_high = parsed_max is not None and (
            parsed > parsed_max if params.get("inclusive_max", True) else parsed >= parsed_max
        )
        if too_low or too_high:
            failures.append((position, "out_of_range" if value_type == "numeric" else "date_out_of_range"))
    return failures


def _compare_values(left: Any, right: Any, *, operator: str, value_type: str) -> bool | None:
    if value_type == "numeric":
        left_value, right_value = _to_decimal(left), _to_decimal(right)
    elif value_type == "date":
        left_value, right_value = _parse_date(left), _parse_date(right)
    else:
        if operator in {"eq", "ne"}:
            equal = _strict_equal(left, right)
            return equal if operator == "eq" else not equal
        left_value, right_value = left, right
    if left_value is None or right_value is None:
        return None
    try:
        return {
            "eq": left_value == right_value,
            "ne": left_value != right_value,
            "gt": left_value > right_value,
            "gte": left_value >= right_value,
            "lt": left_value < right_value,
            "lte": left_value <= right_value,
        }[operator]
    except (TypeError, ValueError):
        return None


def _evaluate_rule(
    df: pd.DataFrame,
    rule: ValidationRule,
    *,
    include_values: bool,
    max_chars: int,
) -> tuple[RuleValidationResult, list[dict[str, Any]]]:
    if not rule.enabled:
        return (
            RuleValidationResult(
                rule_id=rule.rule_id,
                rule_type=rule.rule_type,
                severity=rule.severity,
                column=rule.column,
                passed=True,
                checked_count=0,
                failure_count=0,
                skipped=True,
            ),
            [],
        )

    required = [rule.column]
    other_column = rule.params.get("other_column") if rule.rule_type == "column_compare" else None
    if other_column is not None:
        required.append(other_column)
    missing_columns = [column for column in required if column not in df.columns]
    if missing_columns:
        failures = [
            _failure_row(
                rule,
                row_position=None,
                code="missing_column",
                other_column=other_column,
                include_values=include_values,
                max_chars=max_chars,
            )
        ]
        return (
            RuleValidationResult(
                rule_id=rule.rule_id,
                rule_type=rule.rule_type,
                severity=rule.severity,
                column=rule.column,
                passed=False,
                checked_count=0,
                failure_count=1,
                failure_codes={"missing_column": 1},
            ),
            failures,
        )

    series = df[rule.column]
    positions_and_codes: list[tuple[int, str]] = []
    params = rule.params

    if rule.rule_type == "not_null":
        mask = _missing_mask(series, blank_as_null=params.get("blank_as_null", True))
        positions_and_codes = [(position, "null_value") for position in np.flatnonzero(mask)]

    elif rule.rule_type == "unique":
        missing = _missing_mask(series, blank_as_null=params.get("blank_as_null", False))
        # When blank strings are configured as null, normalise all such values
        # to the same key before duplicate detection.  This makes the
        # ``ignore_nulls=False`` behaviour explicit and deterministic.
        keys = pd.Series(
            [
                ("null", None) if bool(missing.iloc[position]) else _strict_key(value)
                for position, value in enumerate(series.array)
            ],
            index=series.index,
            dtype=object,
        )
        duplicated = keys.duplicated(keep=False)
        if params.get("ignore_nulls", True):
            duplicated &= ~missing
        positions_and_codes = [(position, "duplicate_value") for position in np.flatnonzero(duplicated)]

    elif rule.rule_type == "range":
        positions_and_codes = _range_failures(series, params)

    elif rule.rule_type == "regex":
        compiled = _safe_regex(params["pattern"])
        mode = params.get("mode", "fullmatch")
        matcher = getattr(compiled, mode)
        for position, value in enumerate(series.array):
            if _is_missing(value):
                if not params.get("ignore_nulls", True):
                    positions_and_codes.append((position, "null_value"))
            elif not isinstance(value, str):
                positions_and_codes.append((position, "type_mismatch"))
            elif matcher(value) is None:
                positions_and_codes.append((position, "regex_mismatch"))

    elif rule.rule_type == "allowed_values":
        allowed = list(params["values"])
        for position, value in enumerate(series.array):
            if _is_missing(value):
                if not params.get("ignore_nulls", True):
                    positions_and_codes.append((position, "null_value"))
            elif not any(_strict_equal(value, candidate) for candidate in allowed):
                positions_and_codes.append((position, "not_allowed"))

    elif rule.rule_type == "numeric":
        for position, value in enumerate(series.array):
            if _is_missing(value):
                if not params.get("ignore_nulls", True):
                    positions_and_codes.append((position, "null_value"))
                continue
            parsed = _to_decimal(value, allow_infinite=params.get("allow_infinite", False))
            if parsed is None:
                positions_and_codes.append((position, "not_numeric"))
            elif params.get("integer_only", False) and parsed != parsed.to_integral_value():
                positions_and_codes.append((position, "not_integer"))

    elif rule.rule_type == "date":
        date_format = params.get("format")
        dayfirst = params.get("dayfirst", False)
        minimum = _parse_date(params.get("min"), date_format=date_format, dayfirst=dayfirst) if "min" in params else None
        maximum = _parse_date(params.get("max"), date_format=date_format, dayfirst=dayfirst) if "max" in params else None
        for position, value in enumerate(series.array):
            if _is_missing(value):
                if not params.get("ignore_nulls", True):
                    positions_and_codes.append((position, "null_value"))
                continue
            parsed = _parse_date(value, date_format=date_format, dayfirst=dayfirst)
            if parsed is None:
                positions_and_codes.append((position, "not_date"))
            elif (minimum is not None and parsed < minimum) or (maximum is not None and parsed > maximum):
                positions_and_codes.append((position, "date_out_of_range"))

    elif rule.rule_type == "column_compare":
        other = df[other_column]
        for position, (left, right) in enumerate(zip(series.array, other.array)):
            if _is_missing(left) or _is_missing(right):
                if not params.get("ignore_nulls", True):
                    positions_and_codes.append((position, "null_value"))
                continue
            comparison = _compare_values(
                left,
                right,
                operator=params["operator"],
                value_type=params.get("value_type", "native"),
            )
            positions_and_codes.append(
                (position, "incomparable" if comparison is None else "comparison_failed")
            ) if not comparison else None

    failures = [
        _failure_row(
            rule,
            row_position=position,
            code=code,
            value=series.iloc[position],
            other_column=other_column,
            other_value=df[other_column].iloc[position] if other_column is not None else None,
            include_values=include_values,
            max_chars=max_chars,
        )
        for position, code in positions_and_codes
    ]
    code_counts = Counter(code for _, code in positions_and_codes)
    result = RuleValidationResult(
        rule_id=rule.rule_id,
        rule_type=rule.rule_type,
        severity=rule.severity,
        column=rule.column,
        passed=not failures,
        checked_count=len(series),
        failure_count=len(failures),
        failure_codes=dict(code_counts),
    )
    return result, failures


def validate_dataframe(
    df: pd.DataFrame,
    rules: Sequence[ValidationRule | Mapping[str, Any]],
    *,
    include_values: bool = False,
    max_value_chars: int = 80,
) -> ValidationReport:
    """Validate ``df`` without mutation and return overview plus failures.

    ``report.failures`` contains one row per failing source row.  It uses
    zero-based ``row_position`` rather than the DataFrame index, which may be a
    customer identifier.  Values are hidden by default; opt in with
    ``include_values=True`` for a bounded preview.
    """

    if not isinstance(df, pd.DataFrame):
        raise TypeError("df 必须是 pandas DataFrame")
    if isinstance(rules, (str, bytes)) or not isinstance(rules, Sequence):
        raise TypeError("rules 必须是 ValidationRule 序列")
    _require_bool(include_values, name="include_values")
    if isinstance(max_value_chars, bool) or not isinstance(max_value_chars, int):
        raise TypeError("max_value_chars 必须是整数")
    if not 1 <= max_value_chars <= MAX_VALUE_PREVIEW_CHARS:
        raise ValueError(f"max_value_chars 必须在 1 到 {MAX_VALUE_PREVIEW_CHARS} 之间")
    if df.columns.duplicated().any():
        raise ValueError("验证不支持重复列名")

    prepared = tuple(
        rule if isinstance(rule, ValidationRule) else ValidationRule.from_dict(rule)
        for rule in rules
    )
    identifiers = [rule.rule_id for rule in prepared]
    duplicate_ids = sorted({item for item in identifiers if identifiers.count(item) > 1})
    if duplicate_ids:
        raise ValueError(f"rule_id 不能重复：{duplicate_ids}")

    results: list[RuleValidationResult] = []
    failure_rows: list[dict[str, Any]] = []
    for rule in prepared:
        result, failures = _evaluate_rule(
            df,
            rule,
            include_values=include_values,
            max_chars=max_value_chars,
        )
        results.append(result)
        failure_rows.extend(failures)

    failure_frame = pd.DataFrame(failure_rows, columns=FAILURE_COLUMNS)
    if not failure_frame.empty:
        failure_frame["row_position"] = pd.array(
            failure_frame["row_position"], dtype="Int64"
        )
    else:
        failure_frame = failure_frame.astype(
            {
                "rule_id": "string",
                "rule_type": "string",
                "severity": "string",
                "column": "string",
                "other_column": "string",
                "row_position": "Int64",
                "code": "string",
                "message": "string",
                "value_preview": "string",
                "other_value_preview": "string",
            }
        )
    severity_counts = Counter(row["severity"] for row in failure_rows)
    blocking = sum(1 for row in failure_rows if row["severity"] == "error")
    return ValidationReport(
        row_count=len(df),
        column_count=df.shape[1],
        rule_results=tuple(results),
        failures=failure_frame,
        blocking_failure_count=blocking,
        severity_counts={severity: severity_counts.get(severity, 0) for severity in ("error", "warning", "info")},
    )


__all__ = [
    "COLUMN_COMPARE_OPERATORS",
    "FAILURE_COLUMNS",
    "RULE_TYPES",
    "RuleValidationResult",
    "ValidationReport",
    "ValidationRule",
    "validate_dataframe",
]
