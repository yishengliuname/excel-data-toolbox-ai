"""Data models shared by the spreadsheet processing engine.

The models in this module intentionally contain no file-system behaviour.  They
are small, serialisable value objects that a desktop or web UI can render
directly.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from copy import deepcopy
import json
from pathlib import Path
from types import MappingProxyType
from typing import Any, Literal, Mapping, Sequence


MissingStrategy = Literal["keep", "drop_rows", "fill"]
MaskStrategy = Literal["partial", "full", "hash", "phone", "email", "name", "id"]


def _json_default(value: Any) -> str:
    """Return a stable string representation for JSON log details."""

    if isinstance(value, Path):
        return str(value)
    if isinstance(value, datetime):
        return value.isoformat()
    return str(value)


def _json_value(value: Any) -> Any:
    """Convert common pandas/numpy/date values into JSON-friendly primitives."""

    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (Path, datetime)):
        return _json_default(value)
    if isinstance(value, Mapping):
        return {str(key): _json_value(item) for key, item in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_value(item) for item in value]
    if hasattr(value, "isoformat"):
        try:
            return value.isoformat()
        except (TypeError, ValueError):
            pass
    if hasattr(value, "item"):
        try:
            return _json_value(value.item())
        except (TypeError, ValueError):
            pass
    return str(value)


def _freeze_value(value: Any) -> Any:
    """Create a defensive, recursively read-only snapshot for audit models."""

    if isinstance(value, Mapping):
        return MappingProxyType(
            {str(key): _freeze_value(item) for key, item in value.items()}
        )
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_value(item) for item in value)
    if isinstance(value, set):
        return frozenset(_freeze_value(item) for item in value)
    try:
        return deepcopy(value)
    except Exception:  # pragma: no cover - defensive for unusual UI objects
        return value


@dataclass(frozen=True)
class ColumnProfile:
    """Summary statistics for one column.

    ``position`` disambiguates workbooks that contain duplicate column labels.
    ``sample_values`` contains at most the caller-requested number of distinct,
    non-null values.
    """

    name: str
    position: int
    dtype: str
    semantic_type: str
    non_null_count: int
    missing_count: int
    missing_percent: float
    unique_count: int
    sample_values: tuple[Any, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "position": self.position,
            "dtype": self.dtype,
            "semantic_type": self.semantic_type,
            "non_null_count": self.non_null_count,
            "missing_count": self.missing_count,
            "missing_percent": self.missing_percent,
            "unique_count": self.unique_count,
            "sample_values": [_json_value(value) for value in self.sample_values],
        }


@dataclass(frozen=True)
class DataFrameProfile:
    """Overview of a DataFrame returned by :func:`profile_dataframe`."""

    row_count: int
    column_count: int
    duplicate_row_count: int
    missing_cell_count: int
    memory_bytes: int
    columns: tuple[ColumnProfile, ...]

    def to_dict(self) -> dict[str, Any]:
        return {
            "row_count": self.row_count,
            "column_count": self.column_count,
            "duplicate_row_count": self.duplicate_row_count,
            "missing_cell_count": self.missing_cell_count,
            "memory_bytes": self.memory_bytes,
            "columns": [column.to_dict() for column in self.columns],
        }


@dataclass(frozen=True)
class CleaningConfig:
    """Configuration for the conservative smart-cleaning pipeline.

    Type inference converts a text column only when at least
    ``type_inference_threshold`` of its non-null values parse successfully.
    The default of ``1.0`` avoids silently discarding malformed values.
    """

    trim_whitespace: bool = True
    normalize_blank_strings: bool = True
    drop_empty_rows: bool = True
    drop_empty_columns: bool = True
    drop_duplicates: bool = True
    duplicate_subset: tuple[str, ...] | None = None
    keep_duplicate: Literal["first", "last", False] = "first"
    infer_types: bool = True
    type_inference_threshold: float = 1.0
    missing_strategy: MissingStrategy = "keep"
    missing_subset: tuple[str, ...] | None = None
    drop_missing_how: Literal["any", "all"] = "any"
    fill_values: Mapping[str, Any] = field(default_factory=dict)
    fill_numeric_with: int | float | None = 0
    fill_text_with: str | None = "未填写"
    fill_boolean_with: bool | None = False
    reset_index: bool = True

    def __post_init__(self) -> None:
        if not 0 < self.type_inference_threshold <= 1:
            raise ValueError("type_inference_threshold 必须在 (0, 1] 范围内")
        object.__setattr__(self, "fill_values", _freeze_value(self.fill_values))


@dataclass(frozen=True)
class CleaningReport:
    """Before/after metrics emitted by :func:`smart_clean`."""

    rows_before: int
    rows_after: int
    columns_before: int
    columns_after: int
    empty_rows_removed: int
    empty_columns_removed: tuple[str, ...]
    duplicate_rows_removed: int
    missing_cells_before: int
    missing_cells_after: int
    inferred_types: Mapping[str, str]
    coerced_to_missing: Mapping[str, int] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "inferred_types", _freeze_value(self.inferred_types))
        object.__setattr__(
            self, "coerced_to_missing", _freeze_value(self.coerced_to_missing)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "rows_before": self.rows_before,
            "rows_after": self.rows_after,
            "columns_before": self.columns_before,
            "columns_after": self.columns_after,
            "empty_rows_removed": self.empty_rows_removed,
            "empty_columns_removed": list(self.empty_columns_removed),
            "duplicate_rows_removed": self.duplicate_rows_removed,
            "missing_cells_before": self.missing_cells_before,
            "missing_cells_after": self.missing_cells_after,
            "inferred_types": dict(self.inferred_types),
            "coerced_to_missing": dict(self.coerced_to_missing),
        }


@dataclass(frozen=True)
class OperationRecord:
    """One auditable processing action."""

    timestamp: str
    action: str
    input_tables: tuple[str, ...] = ()
    output_tables: tuple[str, ...] = ()
    details: Mapping[str, Any] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "input_tables", tuple(self.input_tables))
        object.__setattr__(self, "output_tables", tuple(self.output_tables))
        object.__setattr__(self, "details", _freeze_value(self.details))

    def to_dict(self, *, flatten_details: bool = False) -> dict[str, Any]:
        result: dict[str, Any] = {
            "timestamp": self.timestamp,
            "action": self.action,
            "input_tables": list(self.input_tables),
            "output_tables": list(self.output_tables),
        }
        if flatten_details:
            result["details"] = json.dumps(
                dict(self.details), ensure_ascii=False, default=_json_default, sort_keys=True
            )
        else:
            result["details"] = _json_value(dict(self.details))
        return result


class OperationLog:
    """A small append-only operation log suitable for UI sessions.

    Data processing functions never mutate their DataFrame inputs.  If a log is
    supplied, appending an audit record is their only intentional state change.
    Use :attr:`entries` to obtain an immutable snapshot.
    """

    def __init__(self, entries: Sequence[OperationRecord] | None = None) -> None:
        self._entries = list(entries or ())

    @property
    def entries(self) -> tuple[OperationRecord, ...]:
        return tuple(self._entries)

    def record(
        self,
        action: str,
        *,
        input_tables: Sequence[str] = (),
        output_tables: Sequence[str] = (),
        details: Mapping[str, Any] | None = None,
    ) -> OperationRecord:
        record = OperationRecord(
            timestamp=datetime.now(timezone.utc).isoformat(timespec="seconds"),
            action=action,
            input_tables=tuple(str(name) for name in input_tables),
            output_tables=tuple(str(name) for name in output_tables),
            details=dict(details or {}),
        )
        self._entries.append(record)
        return record

    def to_dicts(self, *, flatten_details: bool = False) -> list[dict[str, Any]]:
        return [entry.to_dict(flatten_details=flatten_details) for entry in self._entries]

    def __len__(self) -> int:
        return len(self._entries)


@dataclass(frozen=True)
class ExportResult:
    """Files and table metrics produced by :func:`export_tables`."""

    output_path: Path
    format: Literal["xlsx", "csv", "zip"]
    files: tuple[Path, ...]
    rows_by_table: Mapping[str, int]
    operation_count: int

    def __post_init__(self) -> None:
        object.__setattr__(self, "output_path", Path(self.output_path))
        object.__setattr__(self, "files", tuple(Path(path) for path in self.files))
        object.__setattr__(self, "rows_by_table", _freeze_value(self.rows_by_table))

    def to_dict(self) -> dict[str, Any]:
        return {
            "output_path": str(self.output_path),
            "format": self.format,
            "files": [str(path) for path in self.files],
            "rows_by_table": dict(self.rows_by_table),
            "operation_count": self.operation_count,
        }
