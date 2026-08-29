"""Safe, reusable processing recipes for one pandas DataFrame.

Recipes are deliberately declarative.  They contain JSON values and a small
registry of built-in operations; Python expressions, import paths and
callables are never accepted.  All operations return new DataFrames.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
import hashlib
import json
import math
from types import MappingProxyType
from typing import Any, Literal

import numpy as np
import pandas as pd

from .core import group_summary, select_rename_sort, smart_clean
from .models import CleaningConfig


RECIPE_SCHEMA_VERSION = 1
MAX_RECIPE_BYTES = 1_000_000
MAX_RECIPE_STEPS = 100
MAX_JSON_DEPTH = 20
MAX_COLLECTION_ITEMS = 10_000

RecipeOperation = Literal[
    "clean",
    "replace",
    "select_rename_sort",
    "fill_missing",
    "drop_duplicates",
    "filter",
    "summary",
]

ALLOWED_OPERATIONS: frozenset[str] = frozenset(
    {
        "clean",
        "replace",
        "select_rename_sort",
        "fill_missing",
        "drop_duplicates",
        "filter",
        "summary",
    }
)
FILTER_OPERATORS: frozenset[str] = frozenset(
    {
        "eq",
        "ne",
        "gt",
        "gte",
        "lt",
        "lte",
        "in",
        "not_in",
        "between",
        "contains",
        "starts_with",
        "ends_with",
        "is_null",
        "not_null",
    }
)
SUMMARY_AGGREGATIONS: frozenset[str] = frozenset(
    {"count", "size", "sum", "mean", "min", "max", "median", "nunique", "first", "last"}
)

_CLEAN_KEYS = frozenset(
    {
        "trim_whitespace",
        "normalize_blank_strings",
        "drop_empty_rows",
        "drop_empty_columns",
        "drop_duplicates",
        "duplicate_subset",
        "keep_duplicate",
        "infer_types",
        "type_inference_threshold",
        "missing_strategy",
        "missing_subset",
        "drop_missing_how",
        "fill_values",
        "fill_numeric_with",
        "fill_text_with",
        "fill_boolean_with",
        "reset_index",
    }
)
_STEP_PARAM_KEYS: Mapping[str, frozenset[str]] = MappingProxyType(
    {
        "clean": _CLEAN_KEYS,
        "replace": frozenset({"replacements"}),
        "select_rename_sort": frozenset(
            {"columns", "rename", "sort_by", "ascending", "na_position", "reset_index"}
        ),
        "fill_missing": frozenset({"values", "default", "columns"}),
        "drop_duplicates": frozenset({"subset", "keep", "reset_index"}),
        "filter": frozenset({"conditions", "combine", "reset_index"}),
        "summary": frozenset({"by", "aggregations", "dropna", "sort"}),
    }
)


def _is_json_scalar(value: Any) -> bool:
    return value is None or isinstance(value, (str, bool, int, float))


def _freeze_json(value: Any, *, path: str = "value", depth: int = 0) -> Any:
    """Validate a JSON-compatible value and return an immutable copy."""

    if depth > MAX_JSON_DEPTH:
        raise ValueError(f"{path} 嵌套层级过深")
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(f"{path} 不允许 NaN 或无穷大")
        return value
    if isinstance(value, Mapping):
        if len(value) > MAX_COLLECTION_ITEMS:
            raise ValueError(f"{path} 项目过多")
        frozen: dict[str, Any] = {}
        for key, item in value.items():
            if not isinstance(key, str):
                raise TypeError(f"{path} 的 JSON 对象键必须是字符串")
            if any(ord(char) < 32 for char in key):
                raise ValueError(f"{path} 的键不能包含控制字符")
            frozen[key] = _freeze_json(item, path=f"{path}.{key}", depth=depth + 1)
        return MappingProxyType(frozen)
    if isinstance(value, (list, tuple)):
        if len(value) > MAX_COLLECTION_ITEMS:
            raise ValueError(f"{path} 项目过多")
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


def _require_mapping(value: Any, *, name: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise TypeError(f"{name} 必须是 JSON 对象")
    return value


def _require_bool(value: Any, *, name: str) -> bool:
    if not isinstance(value, bool):
        raise TypeError(f"{name} 必须是布尔值")
    return value


def _require_string(value: Any, *, name: str, allow_empty: bool = False) -> str:
    if not isinstance(value, str):
        raise TypeError(f"{name} 必须是字符串")
    if not allow_empty and not value.strip():
        raise ValueError(f"{name} 不能为空")
    if any(ord(char) < 32 for char in value):
        raise ValueError(f"{name} 不能包含控制字符")
    return value


def _string_list(value: Any, *, name: str, allow_empty: bool = True) -> list[str]:
    if not isinstance(value, (list, tuple)) or isinstance(value, (str, bytes)):
        raise TypeError(f"{name} 必须是字符串列表")
    result = [_require_string(item, name=f"{name}[{index}]") for index, item in enumerate(value)]
    if not allow_empty and not result:
        raise ValueError(f"{name} 不能为空")
    if len(set(result)) != len(result):
        raise ValueError(f"{name} 不能包含重复列名")
    return result


def _validate_scalar(value: Any, *, name: str) -> None:
    if not _is_json_scalar(value):
        raise TypeError(f"{name} 必须是 JSON 标量")
    if isinstance(value, float) and not math.isfinite(value):
        raise ValueError(f"{name} 不允许 NaN 或无穷大")


def _validate_step_params(operation: str, params: Mapping[str, Any]) -> None:
    allowed = _STEP_PARAM_KEYS[operation]
    unknown = sorted(set(params) - allowed)
    if unknown:
        raise ValueError(f"{operation} 包含不支持的参数：{unknown}")

    if operation == "clean":
        bool_keys = {
            "trim_whitespace",
            "normalize_blank_strings",
            "drop_empty_rows",
            "drop_empty_columns",
            "drop_duplicates",
            "infer_types",
            "reset_index",
        }
        for key in bool_keys:
            if key in params:
                _require_bool(params[key], name=f"clean.{key}")
        for key in ("duplicate_subset", "missing_subset"):
            if key in params and params[key] is not None:
                _string_list(params[key], name=f"clean.{key}", allow_empty=False)
        if params.get("keep_duplicate", "first") not in {"first", "last", False}:
            raise ValueError("clean.keep_duplicate 必须是 first、last 或 false")
        if "type_inference_threshold" in params:
            threshold = params["type_inference_threshold"]
            if isinstance(threshold, bool) or not isinstance(threshold, (int, float)):
                raise TypeError("clean.type_inference_threshold 必须是数字")
        if params.get("missing_strategy", "keep") not in {"keep", "drop_rows", "fill"}:
            raise ValueError("clean.missing_strategy 必须是 keep、drop_rows 或 fill")
        if params.get("drop_missing_how", "any") not in {"any", "all"}:
            raise ValueError("clean.drop_missing_how 必须是 any 或 all")
        if "fill_values" in params:
            values = _require_mapping(params["fill_values"], name="clean.fill_values")
            for column, value in values.items():
                _require_string(column, name="clean.fill_values 的列名")
                _validate_scalar(value, name=f"clean.fill_values[{column!r}]")
        if "fill_numeric_with" in params and params["fill_numeric_with"] is not None:
            value = params["fill_numeric_with"]
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                raise TypeError("clean.fill_numeric_with 必须是数字或 null")
        if "fill_text_with" in params and params["fill_text_with"] is not None:
            if not isinstance(params["fill_text_with"], str):
                raise TypeError("clean.fill_text_with 必须是字符串或 null")
        if "fill_boolean_with" in params and params["fill_boolean_with"] is not None:
            _require_bool(params["fill_boolean_with"], name="clean.fill_boolean_with")
        prepared = _cleaning_config_kwargs(params)
        CleaningConfig(**prepared)
        return

    if operation == "replace":
        replacements = params.get("replacements")
        if not isinstance(replacements, (list, tuple)) or not replacements:
            raise ValueError("replace.replacements 必须是非空列表")
        for index, item in enumerate(replacements):
            rule = _require_mapping(item, name=f"replace.replacements[{index}]")
            unknown_rule = sorted(set(rule) - {"column", "old", "new"})
            if unknown_rule or set(rule) != {"column", "old", "new"}:
                raise ValueError(
                    "每条 replace 规则必须且只能包含 column、old、new"
                )
            _require_string(rule["column"], name=f"replace.replacements[{index}].column")
            _validate_scalar(rule["old"], name=f"replace.replacements[{index}].old")
            _validate_scalar(rule["new"], name=f"replace.replacements[{index}].new")
        return

    if operation == "select_rename_sort":
        if "columns" in params and params["columns"] is not None:
            _string_list(params["columns"], name="select_rename_sort.columns")
        if "rename" in params and params["rename"] is not None:
            rename = _require_mapping(params["rename"], name="select_rename_sort.rename")
            values = []
            for source, target in rename.items():
                _require_string(source, name="select_rename_sort.rename 的源列")
                values.append(_require_string(target, name=f"rename[{source!r}]") )
            if len(values) != len(set(values)):
                raise ValueError("rename 的目标列名不能重复")
        if "sort_by" in params and params["sort_by"] is not None:
            sort_by = params["sort_by"]
            if isinstance(sort_by, str):
                _require_string(sort_by, name="select_rename_sort.sort_by")
            else:
                _string_list(sort_by, name="select_rename_sort.sort_by", allow_empty=False)
        if "ascending" in params:
            ascending = params["ascending"]
            if isinstance(ascending, (list, tuple)):
                if not ascending or not all(isinstance(item, bool) for item in ascending):
                    raise TypeError("ascending 必须是布尔值或非空布尔值列表")
            else:
                _require_bool(ascending, name="ascending")
        if params.get("na_position", "last") not in {"first", "last"}:
            raise ValueError("na_position 必须是 first 或 last")
        if "reset_index" in params:
            _require_bool(params["reset_index"], name="reset_index")
        return

    if operation == "fill_missing":
        has_values = "values" in params and params["values"] is not None
        has_default = "default" in params
        if not has_values and not has_default:
            raise ValueError("fill_missing 至少需要 values 或 default")
        if has_values:
            values = _require_mapping(params["values"], name="fill_missing.values")
            if not values:
                raise ValueError("fill_missing.values 不能为空")
            for column, value in values.items():
                _require_string(column, name="fill_missing.values 的列名")
                _validate_scalar(value, name=f"fill_missing.values[{column!r}]")
        if has_default:
            _validate_scalar(params["default"], name="fill_missing.default")
        if "columns" in params and params["columns"] is not None:
            _string_list(params["columns"], name="fill_missing.columns", allow_empty=False)
        return

    if operation == "drop_duplicates":
        if "subset" in params and params["subset"] is not None:
            _string_list(params["subset"], name="drop_duplicates.subset", allow_empty=False)
        if params.get("keep", "first") not in {"first", "last", False}:
            raise ValueError("drop_duplicates.keep 必须是 first、last 或 false")
        if "reset_index" in params:
            _require_bool(params["reset_index"], name="drop_duplicates.reset_index")
        return

    if operation == "filter":
        conditions = params.get("conditions")
        if not isinstance(conditions, (list, tuple)) or not conditions:
            raise ValueError("filter.conditions 必须是非空列表")
        if params.get("combine", "and") not in {"and", "or"}:
            raise ValueError("filter.combine 必须是 and 或 or")
        if "reset_index" in params:
            _require_bool(params["reset_index"], name="filter.reset_index")
        for index, item in enumerate(conditions):
            condition = _require_mapping(item, name=f"filter.conditions[{index}]")
            unknown_condition = sorted(set(condition) - {"column", "operator", "value"})
            if unknown_condition:
                raise ValueError(f"筛选条件包含不支持的字段：{unknown_condition}")
            if "column" not in condition or "operator" not in condition:
                raise ValueError("筛选条件必须包含 column 和 operator")
            _require_string(condition["column"], name=f"filter.conditions[{index}].column")
            operator = _require_string(
                condition["operator"], name=f"filter.conditions[{index}].operator"
            )
            if operator not in FILTER_OPERATORS:
                raise ValueError(f"不支持的筛选比较符：{operator}")
            needs_value = operator not in {"is_null", "not_null"}
            if needs_value != ("value" in condition):
                if needs_value:
                    raise ValueError(f"筛选比较符 {operator} 需要 value")
                raise ValueError(f"筛选比较符 {operator} 不应提供 value")
            if not needs_value:
                continue
            value = condition["value"]
            if operator in {"in", "not_in", "between"}:
                if not isinstance(value, (list, tuple)):
                    raise TypeError(f"filter {operator} 的 value 必须是列表")
                expected = 2 if operator == "between" else None
                if expected is not None and len(value) != expected:
                    raise ValueError("between 的 value 必须恰好有两个边界")
                if operator in {"in", "not_in"} and not value:
                    raise ValueError(f"{operator} 的 value 不能为空")
                for value_index, member in enumerate(value):
                    _validate_scalar(member, name=f"filter.value[{value_index}]")
            else:
                _validate_scalar(value, name="filter.value")
                if value is None and operator in {"eq", "ne", "gt", "gte", "lt", "lte"}:
                    raise ValueError("空值筛选请使用 is_null 或 not_null")
        return

    if operation == "summary":
        if "by" not in params or "aggregations" not in params:
            raise ValueError("summary 必须包含 by 和 aggregations")
        by = params["by"]
        if isinstance(by, str):
            _require_string(by, name="summary.by")
        else:
            _string_list(by, name="summary.by", allow_empty=False)
        aggregations = _require_mapping(params["aggregations"], name="summary.aggregations")
        if not aggregations:
            raise ValueError("summary.aggregations 不能为空")
        for column, names in aggregations.items():
            _require_string(column, name="summary.aggregations 的列名")
            values = [names] if isinstance(names, str) else list(names) if isinstance(names, (list, tuple)) else []
            if not values or not all(isinstance(item, str) for item in values):
                raise TypeError(f"summary.aggregations[{column!r}] 必须是聚合名或聚合名列表")
            unknown_aggregations = sorted(set(values) - SUMMARY_AGGREGATIONS)
            if unknown_aggregations:
                raise ValueError(f"不支持的汇总聚合：{unknown_aggregations}")
        for key in ("dropna", "sort"):
            if key in params:
                _require_bool(params[key], name=f"summary.{key}")


def _cleaning_config_kwargs(params: Mapping[str, Any]) -> dict[str, Any]:
    prepared = _thaw_json(params)
    for key in ("duplicate_subset", "missing_subset"):
        if key in prepared and prepared[key] is not None:
            prepared[key] = tuple(_string_list(prepared[key], name=f"clean.{key}"))
    if "fill_values" in prepared and prepared["fill_values"] is not None:
        prepared["fill_values"] = dict(
            _require_mapping(prepared["fill_values"], name="clean.fill_values")
        )
    return prepared


@dataclass(frozen=True)
class RecipeStep:
    """One versioned, declarative recipe operation."""

    operation: RecipeOperation | str
    params: Mapping[str, Any] = field(default_factory=dict)
    name: str | None = None
    enabled: bool = True

    def __post_init__(self) -> None:
        operation = _require_string(self.operation, name="operation")
        if operation not in ALLOWED_OPERATIONS:
            raise ValueError(f"不支持的配方操作：{operation}")
        if self.name is not None:
            _require_string(self.name, name="name")
        _require_bool(self.enabled, name="enabled")
        if not isinstance(self.params, Mapping):
            raise TypeError(f"{operation}.params 必须是 JSON 对象")
        frozen_params = _freeze_json(self.params, path=f"{operation}.params")
        assert isinstance(frozen_params, Mapping)
        _validate_step_params(operation, frozen_params)
        object.__setattr__(self, "operation", operation)
        object.__setattr__(self, "params", frozen_params)

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "operation": self.operation,
            "params": _thaw_json(self.params),
            "enabled": self.enabled,
        }
        if self.name is not None:
            result["name"] = self.name
        return result

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "RecipeStep":
        value = _require_mapping(data, name="RecipeStep")
        unknown = sorted(set(value) - {"operation", "params", "name", "enabled"})
        if unknown:
            raise ValueError(f"RecipeStep 包含不支持的字段：{unknown}")
        if "operation" not in value:
            raise ValueError("RecipeStep 缺少 operation")
        return cls(
            operation=value["operation"],
            params=value.get("params", {}),
            name=value.get("name"),
            enabled=value.get("enabled", True),
        )


@dataclass(frozen=True)
class ProcessingRecipe:
    """A JSON-serialisable sequence of safe DataFrame operations."""

    name: str
    steps: Sequence[RecipeStep]
    description: str = ""
    schema_version: int = RECIPE_SCHEMA_VERSION

    def __post_init__(self) -> None:
        _require_string(self.name, name="recipe.name")
        _require_string(self.description, name="recipe.description", allow_empty=True)
        if isinstance(self.schema_version, bool) or self.schema_version != RECIPE_SCHEMA_VERSION:
            raise ValueError(f"仅支持配方 schema_version={RECIPE_SCHEMA_VERSION}")
        if isinstance(self.steps, (str, bytes)) or not isinstance(self.steps, Sequence):
            raise TypeError("recipe.steps 必须是 RecipeStep 序列")
        if len(self.steps) > MAX_RECIPE_STEPS:
            raise ValueError(f"配方最多允许 {MAX_RECIPE_STEPS} 个步骤")
        converted = tuple(
            step if isinstance(step, RecipeStep) else RecipeStep.from_dict(step)
            for step in self.steps
        )
        object.__setattr__(self, "steps", converted)

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "name": self.name,
            "description": self.description,
            "steps": [step.to_dict() for step in self.steps],
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]) -> "ProcessingRecipe":
        value = _require_mapping(data, name="ProcessingRecipe")
        allowed = {"schema_version", "name", "description", "steps"}
        unknown = sorted(set(value) - allowed)
        if unknown:
            raise ValueError(f"ProcessingRecipe 包含未知字段：{unknown}")
        missing = sorted({"name", "steps"} - set(value))
        if missing:
            raise ValueError(f"ProcessingRecipe 缺少字段：{missing}")
        steps = value["steps"]
        if not isinstance(steps, (list, tuple)):
            raise TypeError("ProcessingRecipe.steps 必须是列表")
        return cls(
            name=value["name"],
            description=value.get("description", ""),
            steps=tuple(RecipeStep.from_dict(step) for step in steps),
            schema_version=value.get("schema_version", RECIPE_SCHEMA_VERSION),
        )

    def to_json(self, *, indent: int | None = 2) -> str:
        """Return strict JSON; NaN and arbitrary objects are rejected."""

        return json.dumps(
            self.to_dict(), ensure_ascii=False, allow_nan=False, indent=indent, sort_keys=True
        )

    @classmethod
    def from_json(cls, payload: str | bytes | bytearray) -> "ProcessingRecipe":
        if not isinstance(payload, (str, bytes, bytearray)):
            raise TypeError("配方 JSON 必须是字符串或字节")
        raw = payload.encode("utf-8") if isinstance(payload, str) else bytes(payload)
        if len(raw) > MAX_RECIPE_BYTES:
            raise ValueError(f"配方 JSON 不能超过 {MAX_RECIPE_BYTES} 字节")
        try:
            parsed = json.loads(raw)
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("配方 JSON 无效") from exc
        return cls.from_dict(parsed)


@dataclass(frozen=True)
class DataFrameFingerprint:
    """Non-reversible input identity containing only schema, shape and SHA-256."""

    row_count: int
    column_count: int
    schema: tuple[tuple[str, str], ...]
    sha256: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "row_count": self.row_count,
            "column_count": self.column_count,
            "schema": [
                {"name": name, "dtype": dtype} for name, dtype in self.schema
            ],
            "sha256": self.sha256,
        }


@dataclass(frozen=True)
class RecipeStepReport:
    """Shape delta and warnings for one recipe step."""

    step_index: int
    operation: str
    name: str | None
    status: Literal["applied", "skipped"]
    rows_before: int
    rows_after: int
    columns_before: int
    columns_after: int
    warnings: tuple[str, ...] = ()
    details: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "warnings", tuple(self.warnings))
        frozen = _freeze_json(self.details, path="step_report.details")
        assert isinstance(frozen, Mapping)
        object.__setattr__(self, "details", frozen)

    def to_dict(self) -> dict[str, Any]:
        return {
            "step_index": self.step_index,
            "operation": self.operation,
            "name": self.name,
            "status": self.status,
            "rows_before": self.rows_before,
            "rows_after": self.rows_after,
            "columns_before": self.columns_before,
            "columns_after": self.columns_after,
            "warnings": list(self.warnings),
            "details": _thaw_json(self.details),
        }


@dataclass(frozen=True)
class RecipeRunReport:
    """Complete deterministic report returned by :func:`run_recipe`."""

    recipe_name: str
    recipe_schema_version: int
    dry_run: bool
    input_fingerprint: DataFrameFingerprint
    output_fingerprint: DataFrameFingerprint
    steps: tuple[RecipeStepReport, ...]
    warnings: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "recipe_name": self.recipe_name,
            "recipe_schema_version": self.recipe_schema_version,
            "dry_run": self.dry_run,
            "input_fingerprint": self.input_fingerprint.to_dict(),
            "output_fingerprint": self.output_fingerprint.to_dict(),
            "steps": [step.to_dict() for step in self.steps],
            "warnings": list(self.warnings),
        }


def _fingerprint_value(value: Any) -> str:
    try:
        missing = pd.isna(value)
        if isinstance(missing, (bool, np.bool_)) and bool(missing):
            return "null:"
    except (TypeError, ValueError):
        pass
    if isinstance(value, bool):
        return f"bool:{int(value)}"
    if isinstance(value, (int, np.integer)):
        return f"int:{int(value)}"
    if isinstance(value, (float, np.floating)):
        number = float(value)
        return f"float:{number.hex()}" if math.isfinite(number) else f"float:{number}"
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return f"datetime:{pd.Timestamp(value).isoformat()}"
    if isinstance(value, bytes):
        return f"bytes:{value.hex()}"
    if isinstance(value, str):
        return f"str:{value}"
    if isinstance(value, Decimal):
        return f"decimal:{value}"
    try:
        return "json:" + json.dumps(value, ensure_ascii=False, allow_nan=False, sort_keys=True)
    except (TypeError, ValueError):
        return f"object:{type(value).__module__}.{type(value).__qualname__}:{str(value)}"


def fingerprint_dataframe(df: pd.DataFrame) -> DataFrameFingerprint:
    """Return schema, row count and a content SHA-256 without sample values."""

    if not isinstance(df, pd.DataFrame):
        raise TypeError("df 必须是 pandas DataFrame")
    schema = tuple((str(column), str(dtype)) for column, dtype in zip(df.columns, df.dtypes))
    digest = hashlib.sha256()
    digest.update(
        json.dumps(schema, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    )
    digest.update(f"rows:{len(df)};columns:{df.shape[1]}".encode("ascii"))
    for row in df.itertuples(index=False, name=None):
        for value in row:
            encoded = _fingerprint_value(value).encode("utf-8", errors="surrogatepass")
            digest.update(len(encoded).to_bytes(8, "big"))
            digest.update(encoded)
        digest.update(b"\xff")
    return DataFrameFingerprint(
        row_count=len(df),
        column_count=df.shape[1],
        schema=schema,
        sha256=digest.hexdigest(),
    )


def _validate_frame_for_recipe(df: pd.DataFrame) -> None:
    if not isinstance(df, pd.DataFrame):
        raise TypeError("df 必须是 pandas DataFrame")
    if any(not isinstance(column, str) or not column for column in df.columns):
        raise ValueError("配方要求所有列名都是非空字符串")
    if df.columns.duplicated().any():
        raise ValueError("配方不支持重复列名")


def _require_columns(df: pd.DataFrame, columns: Sequence[str], *, operation: str) -> None:
    missing = [column for column in columns if column not in df.columns]
    if missing:
        raise KeyError(f"{operation} 引用了不存在的列：{missing}")


def _apply_replace(df: pd.DataFrame, params: Mapping[str, Any]) -> tuple[pd.DataFrame, dict[str, Any], list[str]]:
    result = df.copy(deep=True)
    matched = 0
    touched_columns: set[str] = set()
    for item in params["replacements"]:
        column = item["column"]
        _require_columns(result, [column], operation="replace")
        old, new = item["old"], item["new"]
        series = result[column]
        if old is None:
            mask = series.isna()
        else:
            try:
                mask = series.eq(old).fillna(False)
            except (TypeError, ValueError) as exc:
                raise ValueError(f"replace 无法比较列 {column!r} 的值类型") from exc
        count = int(mask.sum())
        if count:
            try:
                result.loc[mask, column] = new
            except (TypeError, ValueError):
                result[column] = result[column].astype(object)
                result.loc[mask, column] = new
            matched += count
            touched_columns.add(column)
    warnings = [] if matched else ["替换规则没有匹配任何单元格"]
    return result, {"matched_cells": matched, "columns": sorted(touched_columns)}, warnings


def _apply_fill_missing(df: pd.DataFrame, params: Mapping[str, Any]) -> tuple[pd.DataFrame, dict[str, Any], list[str]]:
    result = df.copy(deep=True)
    values = dict(params.get("values") or {})
    target_columns = list(params.get("columns") or result.columns)
    _require_columns(result, [*values, *target_columns], operation="fill_missing")
    filled = 0
    for column, value in values.items():
        mask = result[column].isna()
        count = int(mask.sum())
        if count:
            try:
                result.loc[mask, column] = value
            except (TypeError, ValueError):
                result[column] = result[column].astype(object)
                result.loc[mask, column] = value
            filled += count
    if "default" in params:
        default = params["default"]
        for column in target_columns:
            if column in values:
                continue
            mask = result[column].isna()
            count = int(mask.sum())
            if count:
                try:
                    result.loc[mask, column] = default
                except (TypeError, ValueError):
                    result[column] = result[column].astype(object)
                    result.loc[mask, column] = default
                filled += count
    warnings = [] if filled else ["所选列没有缺失值"]
    return result, {"filled_cells": filled, "columns": sorted(set(values) | set(target_columns))}, warnings


def _condition_mask(series: pd.Series, operator: str, value: Any = None) -> pd.Series:
    missing = series.isna()
    try:
        if operator == "is_null":
            mask = missing
        elif operator == "not_null":
            mask = ~missing
        elif operator == "eq":
            mask = series.eq(value) & ~missing
        elif operator == "ne":
            mask = series.ne(value) & ~missing
        elif operator == "gt":
            mask = series.gt(value) & ~missing
        elif operator == "gte":
            mask = series.ge(value) & ~missing
        elif operator == "lt":
            mask = series.lt(value) & ~missing
        elif operator == "lte":
            mask = series.le(value) & ~missing
        elif operator == "in":
            mask = series.isin(list(value)) & ~missing
        elif operator == "not_in":
            mask = ~series.isin(list(value)) & ~missing
        elif operator == "between":
            mask = series.ge(value[0]) & series.le(value[1]) & ~missing
        else:
            text = series.astype("string")
            needle = str(value)
            if operator == "contains":
                mask = text.str.contains(needle, regex=False, na=False)
            elif operator == "starts_with":
                mask = text.str.startswith(needle, na=False)
            elif operator == "ends_with":
                mask = text.str.endswith(needle, na=False)
            else:  # pragma: no cover - construction validation prevents this
                raise ValueError(f"不支持的筛选比较符：{operator}")
    except (TypeError, ValueError) as exc:
        raise ValueError(f"筛选比较符 {operator} 与列数据类型不兼容") from exc
    return mask.fillna(False).astype(bool)


def _apply_filter(df: pd.DataFrame, params: Mapping[str, Any]) -> tuple[pd.DataFrame, dict[str, Any], list[str]]:
    combine = params.get("combine", "and")
    mask = pd.Series(combine == "and", index=df.index, dtype=bool)
    condition_meta: list[dict[str, str]] = []
    for condition in params["conditions"]:
        column, operator = condition["column"], condition["operator"]
        _require_columns(df, [column], operation="filter")
        current = _condition_mask(df[column], operator, condition.get("value"))
        mask = (mask & current) if combine == "and" else (mask | current)
        condition_meta.append({"column": column, "operator": operator})
    result = df.loc[mask].copy(deep=True)
    if params.get("reset_index", True):
        result = result.reset_index(drop=True)
    warnings = ["筛选后没有保留任何数据行"] if result.empty and not df.empty else []
    return result, {"kept_rows": len(result), "removed_rows": len(df) - len(result), "conditions": condition_meta}, warnings


def _apply_step(df: pd.DataFrame, step: RecipeStep) -> tuple[pd.DataFrame, dict[str, Any], list[str]]:
    params = step.params
    if step.operation == "clean":
        result, report = smart_clean(df, CleaningConfig(**_cleaning_config_kwargs(params)))
        warnings = []
        if report.coerced_to_missing:
            total = sum(report.coerced_to_missing.values())
            warnings.append(f"类型推断产生 {total} 个新的缺失值")
        return result, report.to_dict(), warnings

    if step.operation == "replace":
        return _apply_replace(df, params)

    if step.operation == "select_rename_sort":
        prepared = _thaw_json(params)
        result = select_rename_sort(df, **prepared)
        return result, {"selected_columns": list(result.columns)}, []

    if step.operation == "fill_missing":
        return _apply_fill_missing(df, params)

    if step.operation == "drop_duplicates":
        subset = list(params["subset"]) if params.get("subset") is not None else None
        if subset:
            _require_columns(df, subset, operation="drop_duplicates")
        before = len(df)
        result = df.copy(deep=True).drop_duplicates(
            subset=subset, keep=params.get("keep", "first")
        )
        if params.get("reset_index", True):
            result = result.reset_index(drop=True)
        removed = before - len(result)
        warnings = [] if removed else ["没有发现符合规则的重复行"]
        return result, {"removed_rows": removed, "subset": subset or []}, warnings

    if step.operation == "filter":
        return _apply_filter(df, params)

    if step.operation == "summary":
        prepared = _thaw_json(params)
        result = group_summary(df, **prepared)
        return result, {"group_count": len(result), "group_columns": [prepared["by"]] if isinstance(prepared["by"], str) else prepared["by"]}, []

    raise ValueError(f"不支持的配方操作：{step.operation}")  # pragma: no cover


def run_recipe(
    df: pd.DataFrame,
    recipe: ProcessingRecipe,
    *,
    dry_run: bool = False,
) -> tuple[pd.DataFrame, RecipeRunReport]:
    """Execute a recipe and return ``(new_dataframe, structured_report)``.

    ``dry_run=True`` performs the same deterministic transformations on an
    isolated copy and marks the report as a preview.  It never mutates ``df``;
    the caller decides whether to commit the returned preview to its session.
    """

    _validate_frame_for_recipe(df)
    if not isinstance(recipe, ProcessingRecipe):
        raise TypeError("recipe 必须是 ProcessingRecipe")
    _require_bool(dry_run, name="dry_run")

    input_fingerprint = fingerprint_dataframe(df)
    result = df.copy(deep=True)
    reports: list[RecipeStepReport] = []
    run_warnings: list[str] = []
    for index, step in enumerate(recipe.steps, start=1):
        rows_before, columns_before = result.shape
        if not step.enabled:
            reports.append(
                RecipeStepReport(
                    step_index=index,
                    operation=step.operation,
                    name=step.name,
                    status="skipped",
                    rows_before=rows_before,
                    rows_after=rows_before,
                    columns_before=columns_before,
                    columns_after=columns_before,
                    warnings=("步骤已禁用",),
                )
            )
            continue
        transformed, details, warnings = _apply_step(result, step)
        _validate_frame_for_recipe(transformed)
        result = transformed.copy(deep=True)
        reports.append(
            RecipeStepReport(
                step_index=index,
                operation=step.operation,
                name=step.name,
                status="applied",
                rows_before=rows_before,
                rows_after=len(result),
                columns_before=columns_before,
                columns_after=result.shape[1],
                warnings=tuple(warnings),
                details=details,
            )
        )
        run_warnings.extend(f"第 {index} 步：{warning}" for warning in warnings)

    output_fingerprint = fingerprint_dataframe(result)
    report = RecipeRunReport(
        recipe_name=recipe.name,
        recipe_schema_version=recipe.schema_version,
        dry_run=dry_run,
        input_fingerprint=input_fingerprint,
        output_fingerprint=output_fingerprint,
        steps=tuple(reports),
        warnings=tuple(run_warnings),
    )
    return result.copy(deep=True), report


__all__ = [
    "ALLOWED_OPERATIONS",
    "FILTER_OPERATORS",
    "DataFrameFingerprint",
    "ProcessingRecipe",
    "RECIPE_SCHEMA_VERSION",
    "RecipeRunReport",
    "RecipeStep",
    "RecipeStepReport",
    "fingerprint_dataframe",
    "run_recipe",
]
