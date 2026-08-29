"""Pure, UI-friendly tabular transformations for Excel and CSV data.

All transformation functions copy their inputs and return new objects.  Passing
an :class:`OperationLog` is optional; appending an audit event is then the only
intentional state change.  This makes the module suitable for Streamlit, Flask,
desktop GUI, or command-line front ends.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, Sequence
import hashlib
from pathlib import Path
import re
from typing import Any, Literal

import pandas as pd
from pandas.api import types as ptypes

from .io_utils import export_tables_to_path, load_tables_from_files, sanitise_sheet_name
from .models import (
    CleaningConfig,
    CleaningReport,
    ColumnProfile,
    DataFrameProfile,
    ExportResult,
    MaskStrategy,
    OperationLog,
    OperationRecord,
)


_TRUE_TEXT = {"true", "yes", "y", "是", "真", "对"}
_FALSE_TEXT = {"false", "no", "n", "否", "假", "错"}
_LEADING_ZERO_ID = re.compile(r"^[+-]?0\d+$")
_DATE_HINT = re.compile(r"[-/:年月日Tt]")
_IDENTIFIER_NAME = re.compile(
    r"(^|[_\s])(id|uuid|code)($|[_\s])|编号|编码|代码|单号|证件|身份证|手机号|电话|邮编",
    re.I,
)


def _hashable_profile_value(value: Any) -> Any:
    """Return a stable, hashable proxy for nested values used in profiles."""

    try:
        hash(value)
        return value
    except TypeError:
        return repr(value)


def _profile_series(series: pd.Series) -> pd.Series:
    return series.map(_hashable_profile_value, na_action="ignore")


def _missing_mask(series: pd.Series) -> pd.Series:
    """Return a mask that also treats empty and whitespace text as missing."""

    mask = series.isna()
    if ptypes.is_object_dtype(series.dtype) or ptypes.is_string_dtype(series.dtype):
        mask = mask | series.map(
            lambda value: isinstance(value, str) and not value.strip(),
            na_action="ignore",
        ).fillna(False)
    return mask.astype(bool)


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
        raise ValueError(f"{name} 含重复列名，处理前请先重命名：{duplicates}")


def _normalise_columns(columns: str | Sequence[str], *, argument: str) -> list[str]:
    result = [columns] if isinstance(columns, str) else list(columns)
    if not result:
        raise ValueError(f"{argument} 不能为空")
    return result


def _validate_columns(frame: pd.DataFrame, columns: Sequence[str], *, argument: str) -> None:
    missing = [column for column in columns if column not in frame.columns]
    if missing:
        raise KeyError(f"{argument} 包含不存在的列：{missing}")


def _log(
    log: OperationLog | None,
    action: str,
    *,
    input_tables: Sequence[str] = (),
    output_tables: Sequence[str] = (),
    details: Mapping[str, Any] | None = None,
) -> None:
    if log is not None:
        log.record(
            action,
            input_tables=input_tables,
            output_tables=output_tables,
            details=details,
        )


def load_tables(
    paths: str | Path | Iterable[str | Path],
    *,
    csv_encoding: str | None = None,
    csv_options: Mapping[str, Any] | None = None,
    excel_options: Mapping[str, Any] | None = None,
    log: OperationLog | None = None,
) -> dict[str, pd.DataFrame]:
    """Read every supplied Excel worksheet and CSV/TSV into a table mapping.

    Returns:
        A new ``dict[str, DataFrame]`` keyed as ``文件名::工作表``. CSV and TSV
        inputs use ``文件名::CSV``. Duplicate keys receive a numeric suffix.
    """

    tables = load_tables_from_files(
        paths,
        csv_encoding=csv_encoding,
        csv_options=csv_options,
        excel_options=excel_options,
    )
    _log(
        log,
        "读取文件",
        output_tables=list(tables),
        details={
            "table_count": len(tables),
            "rows": {name: len(frame) for name, frame in tables.items()},
        },
    )
    return tables


def _semantic_type(series: pd.Series, column_name: str) -> str:
    non_null = series.loc[~_missing_mask(series)]
    if non_null.empty:
        return "empty"
    if ptypes.is_bool_dtype(series.dtype):
        return "boolean"
    if ptypes.is_datetime64_any_dtype(series.dtype):
        return "datetime"
    if ptypes.is_integer_dtype(series.dtype):
        return "integer"
    if ptypes.is_numeric_dtype(series.dtype):
        return "number"
    unique_count = int(_profile_series(non_null).nunique(dropna=True))
    if _IDENTIFIER_NAME.search(str(column_name)) or (
        unique_count == len(non_null) and non_null.astype(str).str.len().median() >= 6
    ):
        return "identifier"
    if unique_count <= max(20, int(len(non_null) * 0.1)):
        return "category"
    return "text"


def profile_dataframe(df: pd.DataFrame, *, sample_values: int = 3) -> DataFrameProfile:
    """Create row/column, missing-value, duplicate, type, and sample statistics.

    Returns:
        :class:`DataFrameProfile`; call ``to_dict()`` for JSON/UI rendering.
    """

    _require_dataframe(df, allow_duplicate_columns=True)
    if sample_values < 0:
        raise ValueError("sample_values 不能为负数")
    columns: list[ColumnProfile] = []
    row_count = len(df)
    for position in range(df.shape[1]):
        series = df.iloc[:, position]
        missing_mask = _missing_mask(series)
        comparable = _profile_series(series.mask(missing_mask))
        label = str(df.columns[position])
        missing_count = int(missing_mask.sum())
        sample_positions = comparable.notna() & ~comparable.duplicated(keep="first")
        samples = tuple(series.loc[sample_positions].head(sample_values).tolist())
        columns.append(
            ColumnProfile(
                name=label,
                position=position,
                dtype=str(series.dtype),
                semantic_type=_semantic_type(series, label),
                non_null_count=int((~missing_mask).sum()),
                missing_count=missing_count,
                missing_percent=round((missing_count / row_count * 100) if row_count else 0.0, 2),
                unique_count=int(comparable.nunique(dropna=True)),
                sample_values=samples,
            )
        )
    comparable_frame = df.copy(deep=False)
    missing_cell_count = 0
    for position in range(comparable_frame.shape[1]):
        missing_mask = _missing_mask(comparable_frame.iloc[:, position])
        missing_cell_count += int(missing_mask.sum())
        comparable_frame.iloc[:, position] = comparable_frame.iloc[:, position].mask(
            missing_mask
        )
    comparable_frame = comparable_frame.apply(_profile_series)
    return DataFrameProfile(
        row_count=row_count,
        column_count=df.shape[1],
        duplicate_row_count=int(comparable_frame.duplicated().sum()),
        missing_cell_count=missing_cell_count,
        memory_bytes=int(df.memory_usage(index=True, deep=True).sum()),
        columns=tuple(columns),
    )


def _strip_text_values(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy(deep=True)
    for column in result.columns:
        series = result[column]
        if ptypes.is_object_dtype(series.dtype) or ptypes.is_string_dtype(series.dtype):
            result[column] = series.map(
                lambda value: value.strip() if isinstance(value, str) else value,
                na_action="ignore",
            )
    return result


def _normalise_blanks(frame: pd.DataFrame) -> pd.DataFrame:
    result = frame.copy(deep=True)
    for column in result.columns:
        series = result[column]
        if ptypes.is_object_dtype(series.dtype) or ptypes.is_string_dtype(series.dtype):
            blank_mask = series.map(
                lambda value: isinstance(value, str) and not value.strip(),
                na_action="ignore",
            ).fillna(False)
            if bool(blank_mask.any()):
                result.loc[blank_mask, column] = pd.NA
    return result


def _infer_text_series(
    series: pd.Series, threshold: float
) -> tuple[pd.Series, str | None, int]:
    """Infer bool, number, or date while protecting leading-zero identifiers."""

    non_null_mask = series.notna()
    if not bool(non_null_mask.any()):
        return series, None, 0
    raw = series.loc[non_null_mask]
    if not raw.map(lambda value: isinstance(value, str)).all():
        return series, None, 0
    text = raw.astype(str).str.strip()
    lowered = text.str.casefold()

    bool_tokens = _TRUE_TEXT | _FALSE_TEXT
    bool_success = lowered.isin(bool_tokens)
    if float(bool_success.mean()) >= threshold:
        mapping = {token: True for token in _TRUE_TEXT} | {
            token: False for token in _FALSE_TEXT
        }
        converted = lowered.map(mapping).astype("boolean")
        result = pd.Series(pd.NA, index=series.index, dtype="boolean")
        result.loc[non_null_mask] = converted.to_numpy()
        return result, "boolean", int((~bool_success).sum())

    has_leading_zero_id = text.map(lambda value: bool(_LEADING_ZERO_ID.match(value))).any()
    integer_text = text.str.fullmatch(r"[+-]?\d+")
    has_long_integer = bool(
        (integer_text & (text.str.lstrip("+-").str.len() > 15)).any()
    )
    if has_long_integer:
        return series, None, 0
    if not has_leading_zero_id:
        if float(integer_text.mean()) >= threshold:
            parsed_values = [int(value) for value in text.loc[integer_text]]
            int64_min, int64_max = -(2**63), 2**63 - 1
            if all(int64_min <= value <= int64_max for value in parsed_values):
                result = pd.Series(pd.NA, index=series.index, dtype="Int64")
                converted = pd.Series(pd.NA, index=text.index, dtype="Int64")
                converted.loc[integer_text] = pd.array(parsed_values, dtype="Int64")
                result.loc[non_null_mask] = converted.to_numpy()
                return result, "integer", int((~integer_text).sum())
        numeric = pd.to_numeric(text, errors="coerce")
        numeric_success = numeric.notna()
        if float(numeric_success.mean()) >= threshold:
            result = pd.Series(float("nan"), index=series.index, dtype="float64")
            result.loc[non_null_mask] = numeric.to_numpy()
            return result, "number", int((~numeric_success).sum())

    if bool(text.map(lambda value: bool(_DATE_HINT.search(value))).any()):
        try:
            parsed = pd.to_datetime(text, errors="coerce", format="mixed")
        except (TypeError, ValueError):  # pandas < 2.0 has no format='mixed'
            parsed = pd.to_datetime(text, errors="coerce")
        date_success = parsed.notna()
        if float(date_success.mean()) >= threshold:
            parsed_iterator = iter(parsed.array)
            values = [
                next(parsed_iterator) if present else pd.NaT
                for present in non_null_mask.to_numpy()
            ]
            result = pd.Series(
                pd.array(values, dtype=parsed.dtype), index=series.index
            )
            return result, "datetime", int((~date_success).sum())
    return series, None, 0


def _infer_types(
    frame: pd.DataFrame, threshold: float
) -> tuple[pd.DataFrame, dict[str, str], dict[str, int]]:
    result = frame.copy(deep=True)
    inferred: dict[str, str] = {}
    coerced: dict[str, int] = {}
    for column in result.columns:
        series = result[column]
        if not (ptypes.is_object_dtype(series.dtype) or ptypes.is_string_dtype(series.dtype)):
            continue
        if _IDENTIFIER_NAME.search(str(column)):
            continue
        converted, inferred_type, coerced_count = _infer_text_series(series, threshold)
        if inferred_type:
            result[column] = converted
            inferred[str(column)] = inferred_type
            if coerced_count:
                coerced[str(column)] = coerced_count
    return result, inferred, coerced


def _smart_fill(frame: pd.DataFrame, config: CleaningConfig) -> pd.DataFrame:
    result = frame.copy(deep=True)
    explicit = dict(config.fill_values)
    unknown_columns = [name for name in explicit if name not in result.columns]
    if unknown_columns:
        raise KeyError(f"fill_values 包含不存在的列：{unknown_columns}")
    for column in result.columns:
        if column in explicit:
            fill_value = explicit[column]
        elif ptypes.is_bool_dtype(result[column].dtype):
            fill_value = config.fill_boolean_with
        elif ptypes.is_numeric_dtype(result[column].dtype):
            fill_value = config.fill_numeric_with
        elif ptypes.is_datetime64_any_dtype(result[column].dtype):
            fill_value = None
        else:
            fill_value = config.fill_text_with
        if fill_value is not None:
            try:
                result[column] = result[column].fillna(fill_value)
            except (TypeError, ValueError):
                result[column] = result[column].astype(object).fillna(fill_value)
    return result


def smart_clean(
    df: pd.DataFrame,
    config: CleaningConfig | None = None,
    *,
    table_name: str = "数据表",
    log: OperationLog | None = None,
) -> tuple[pd.DataFrame, CleaningReport]:
    """Conservatively clean whitespace, blanks, duplicates, missing data, and types.

    Returns:
        ``(cleaned_dataframe, cleaning_report)``. The input is never modified.
    """

    _require_dataframe(df)
    config = config or CleaningConfig()
    result = df.copy(deep=True)
    rows_before, columns_before = result.shape
    missing_before = int(result.isna().sum().sum())

    if config.trim_whitespace:
        result = _strip_text_values(result)
    if config.normalize_blank_strings:
        result = _normalise_blanks(result)

    rows_before_empty_drop = len(result)
    if config.drop_empty_rows:
        result = result.dropna(axis=0, how="all")
    empty_rows_removed = rows_before_empty_drop - len(result)

    columns_before_empty_drop = list(result.columns)
    if config.drop_empty_columns:
        result = result.dropna(axis=1, how="all")
    empty_columns_removed = tuple(
        str(column) for column in columns_before_empty_drop if column not in result.columns
    )

    inferred: dict[str, str] = {}
    coerced: dict[str, int] = {}
    if config.infer_types:
        result, inferred, coerced = _infer_types(
            result, config.type_inference_threshold
        )

    subset = list(config.missing_subset) if config.missing_subset else None
    if subset:
        _validate_columns(result, subset, argument="missing_subset")
    if config.missing_strategy == "drop_rows":
        result = result.dropna(axis=0, how=config.drop_missing_how, subset=subset)
    elif config.missing_strategy == "fill":
        result = _smart_fill(result, config)
    elif config.missing_strategy != "keep":
        raise ValueError("missing_strategy 必须是 keep、drop_rows 或 fill")

    duplicate_rows_removed = 0
    if config.drop_duplicates:
        duplicate_subset = list(config.duplicate_subset) if config.duplicate_subset else None
        if duplicate_subset:
            _validate_columns(result, duplicate_subset, argument="duplicate_subset")
        before = len(result)
        result = result.drop_duplicates(
            subset=duplicate_subset, keep=config.keep_duplicate
        )
        duplicate_rows_removed = before - len(result)
    if config.reset_index:
        result = result.reset_index(drop=True)

    report = CleaningReport(
        rows_before=rows_before,
        rows_after=len(result),
        columns_before=columns_before,
        columns_after=result.shape[1],
        empty_rows_removed=empty_rows_removed,
        empty_columns_removed=empty_columns_removed,
        duplicate_rows_removed=duplicate_rows_removed,
        missing_cells_before=missing_before,
        missing_cells_after=int(result.isna().sum().sum()),
        inferred_types=inferred,
        coerced_to_missing=coerced,
    )
    _log(
        log,
        "智能清洗",
        input_tables=[table_name],
        output_tables=[table_name],
        details=report.to_dict(),
    )
    return result, report


def select_rename_sort(
    df: pd.DataFrame,
    *,
    columns: Sequence[str] | None = None,
    rename: Mapping[str, str] | None = None,
    sort_by: str | Sequence[str] | None = None,
    ascending: bool | Sequence[bool] = True,
    na_position: Literal["first", "last"] = "last",
    reset_index: bool = True,
    table_name: str = "数据表",
    output_name: str | None = None,
    log: OperationLog | None = None,
) -> pd.DataFrame:
    """Select/reorder columns, rename them, and optionally sort rows.

    Returns:
        A transformed copy of ``df``.
    """

    _require_dataframe(df)
    result = df.copy(deep=True)
    if columns is not None:
        selected = list(columns)
        _validate_columns(result, selected, argument="columns")
        result = result.loc[:, selected].copy()
    if rename:
        unknown = [column for column in rename if column not in result.columns]
        if unknown:
            raise KeyError(f"rename 包含不存在的列：{unknown}")
        result = result.rename(columns=dict(rename))
        if result.columns.duplicated().any():
            raise ValueError("重命名后产生了重复列名")
    if sort_by is not None:
        sort_columns = _normalise_columns(sort_by, argument="sort_by")
        _validate_columns(result, sort_columns, argument="sort_by")
        result = result.sort_values(
            by=sort_columns, ascending=ascending, na_position=na_position, kind="stable"
        )
    if reset_index:
        result = result.reset_index(drop=True)
    _log(
        log,
        "列选择/重命名/排序",
        input_tables=[table_name],
        output_tables=[output_name or table_name],
        details={
            "columns": list(result.columns),
            "rename": dict(rename or {}),
            "sort_by": [] if sort_by is None else _normalise_columns(sort_by, argument="sort_by"),
        },
    )
    return result


def concat_tables(
    tables: Mapping[str, pd.DataFrame] | Sequence[pd.DataFrame],
    *,
    join: Literal["outer", "inner"] = "outer",
    ignore_index: bool = True,
    source_column: str | None = "来源表",
    output_name: str = "纵向合并结果",
    log: OperationLog | None = None,
) -> pd.DataFrame:
    """Vertically append tables by column name.

    Returns:
        One new DataFrame. With mapping input, ``source_column`` records each
        source table; pass ``None`` to disable it.
    """

    if isinstance(tables, Mapping):
        items = list(tables.items())
    else:
        items = [(str(index + 1), frame) for index, frame in enumerate(tables)]
    if not items:
        raise ValueError("至少需要一个数据表进行纵向合并")
    prepared: list[pd.DataFrame] = []
    for name, frame in items:
        _require_dataframe(frame, f"tables[{name!r}]")
        copied = frame.copy(deep=True)
        if source_column is not None:
            if source_column in copied.columns:
                raise ValueError(f"来源列 {source_column!r} 已存在")
            copied.insert(0, source_column, str(name))
        prepared.append(copied)
    result = pd.concat(prepared, axis=0, join=join, ignore_index=ignore_index, sort=False)
    _log(
        log,
        "纵向合并",
        input_tables=[str(name) for name, _ in items],
        output_tables=[output_name],
        details={"join": join, "row_count": len(result), "source_column": source_column},
    )
    return result


def join_tables(
    left: pd.DataFrame,
    right: pd.DataFrame,
    *,
    on: str | Sequence[str] | None = None,
    left_on: str | Sequence[str] | None = None,
    right_on: str | Sequence[str] | None = None,
    how: Literal["left", "right", "inner", "outer", "cross"] = "left",
    suffixes: tuple[str, str] = ("_左", "_右"),
    validate: str | None = None,
    left_name: str = "左表",
    right_name: str = "右表",
    output_name: str = "键连接结果",
    log: OperationLog | None = None,
) -> pd.DataFrame:
    """Join two tables using pandas merge semantics and return a new DataFrame."""

    _require_dataframe(left, "left")
    _require_dataframe(right, "right")
    if how != "cross" and on is None and (left_on is None or right_on is None):
        raise ValueError("非 cross 连接必须提供 on，或同时提供 left_on 和 right_on")
    result = pd.merge(
        left.copy(deep=True),
        right.copy(deep=True),
        how=how,
        on=on,
        left_on=left_on,
        right_on=right_on,
        suffixes=suffixes,
        validate=validate,
        sort=False,
    )
    _log(
        log,
        "键连接",
        input_tables=[left_name, right_name],
        output_tables=[output_name],
        details={
            "how": how,
            "on": on,
            "left_on": left_on,
            "right_on": right_on,
            "row_count": len(result),
        },
    )
    return result


def lookup_match(
    source: pd.DataFrame,
    lookup: pd.DataFrame,
    *,
    source_key: str | Sequence[str],
    lookup_key: str | Sequence[str] | None = None,
    value_columns: Sequence[str] | None = None,
    keep_lookup_duplicate: Literal["first", "last", False] = "first",
    add_match_column: bool = True,
    match_column: str = "匹配状态",
    source_name: str = "源表",
    lookup_name: str = "查找表",
    output_name: str = "跨表匹配结果",
    log: OperationLog | None = None,
) -> pd.DataFrame:
    """Perform Excel-VLOOKUP-style matching without multiplying source rows.

    Duplicate keys in the lookup table are resolved with
    ``keep_lookup_duplicate``. Returns the source rows plus selected lookup
    columns and, by default, a ``匹配状态`` column.
    """

    _require_dataframe(source, "source")
    _require_dataframe(lookup, "lookup")
    source_keys = _normalise_columns(source_key, argument="source_key")
    lookup_keys = (
        source_keys
        if lookup_key is None
        else _normalise_columns(lookup_key, argument="lookup_key")
    )
    if len(source_keys) != len(lookup_keys):
        raise ValueError("source_key 和 lookup_key 的列数必须相同")
    _validate_columns(source, source_keys, argument="source_key")
    _validate_columns(lookup, lookup_keys, argument="lookup_key")
    values = (
        [column for column in lookup.columns if column not in lookup_keys]
        if value_columns is None
        else list(value_columns)
    )
    _validate_columns(lookup, values, argument="value_columns")
    selected_lookup = lookup.loc[:, [*lookup_keys, *values]].copy()
    duplicate_count = int(selected_lookup.duplicated(subset=lookup_keys, keep=False).sum())
    temporary_keys: list[str] = []
    occupied_names = {str(column) for column in [*source.columns, *selected_lookup.columns]}
    for position, lookup_column in enumerate(lookup_keys):
        candidate = f"__toolbox_lookup_key_{position}__"
        while candidate in occupied_names:
            candidate += "_"
        occupied_names.add(candidate)
        temporary_keys.append(candidate)
    selected_lookup = selected_lookup.rename(
        columns=dict(zip(lookup_keys, temporary_keys))
    )
    selected_lookup = selected_lookup.drop_duplicates(
        subset=temporary_keys, keep=keep_lookup_duplicate
    )
    indicator = "__toolbox_match__"
    while indicator in source.columns or indicator in selected_lookup.columns:
        indicator += "_"
    result = source.copy(deep=True).merge(
        selected_lookup,
        how="left",
        left_on=source_keys,
        right_on=temporary_keys,
        suffixes=("", "_查找"),
        indicator=indicator,
        sort=False,
        validate="many_to_one",
    )
    result = result.drop(columns=temporary_keys)
    matched_count = int((result[indicator] == "both").sum())
    if add_match_column:
        if match_column in result.columns:
            raise ValueError(f"匹配状态列 {match_column!r} 已存在")
        result[match_column] = result[indicator].map(
            {"both": "已匹配", "left_only": "未匹配", "right_only": "已匹配"}
        )
    result = result.drop(columns=indicator)
    _log(
        log,
        "跨表查找匹配",
        input_tables=[source_name, lookup_name],
        output_tables=[output_name],
        details={
            "source_key": source_keys,
            "lookup_key": lookup_keys,
            "value_columns": values,
            "matched_rows": matched_count,
            "unmatched_rows": len(result) - matched_count,
            "lookup_duplicate_rows": duplicate_count,
        },
    )
    return result.reset_index(drop=True)


def _flatten_columns(columns: pd.Index) -> list[str]:
    if not isinstance(columns, pd.MultiIndex):
        return [str(column) for column in columns]
    flattened: list[str] = []
    for parts in columns.to_flat_index():
        text = "_".join(str(part) for part in parts if str(part) not in {"", "None"})
        flattened.append(text)
    return flattened


def group_summary(
    df: pd.DataFrame,
    *,
    by: str | Sequence[str],
    aggregations: Mapping[str, str | Sequence[str]],
    dropna: bool = False,
    sort: bool = True,
    table_name: str = "数据表",
    output_name: str = "分组汇总",
    log: OperationLog | None = None,
) -> pd.DataFrame:
    """Group rows and calculate one or more named pandas aggregations.

    Example: ``aggregations={"销售额": ["sum", "mean"], "订单号": "count"}``.
    Returns a flat-column DataFrame.
    """

    _require_dataframe(df)
    group_columns = _normalise_columns(by, argument="by")
    _validate_columns(df, group_columns, argument="by")
    if not aggregations:
        raise ValueError("aggregations 不能为空")
    _validate_columns(df, list(aggregations), argument="aggregations")
    result = (
        df.copy(deep=True)
        .groupby(
            group_columns,
            dropna=dropna,
            sort=sort,
            as_index=False,
            observed=True,
        )
        .agg(dict(aggregations))
    )
    result.columns = _flatten_columns(result.columns)
    _log(
        log,
        "分组汇总",
        input_tables=[table_name],
        output_tables=[output_name],
        details={
            "by": group_columns,
            "aggregations": {key: value for key, value in aggregations.items()},
            "group_count": len(result),
        },
    )
    return result.reset_index(drop=True)


def _split_label(key: Any) -> str:
    parts = key if isinstance(key, tuple) else (key,)
    text = "_".join("缺失值" if pd.isna(part) else str(part) for part in parts)
    return sanitise_sheet_name(text, fallback="空白组")


def split_dataframe(
    df: pd.DataFrame,
    *,
    by: str | Sequence[str] | None = None,
    rows_per_table: int | None = None,
    drop_group_columns: bool = False,
    table_name: str = "数据表",
    log: OperationLog | None = None,
) -> dict[str, pd.DataFrame]:
    """Split a DataFrame by group values or into fixed-size row chunks.

    Exactly one of ``by`` and ``rows_per_table`` must be supplied. Returns a
    mapping whose keys are safe, unique Excel worksheet names.
    """

    _require_dataframe(df)
    if (by is None) == (rows_per_table is None):
        raise ValueError("必须且只能提供 by 或 rows_per_table 其中一个")
    result: dict[str, pd.DataFrame] = {}
    used: set[str] = set()

    def add(name: str, frame: pd.DataFrame) -> None:
        base = sanitise_sheet_name(name)
        candidate = base
        counter = 2
        while candidate.casefold() in used:
            suffix = f"_{counter}"
            candidate = f"{base[: 31 - len(suffix)]}{suffix}"
            counter += 1
        used.add(candidate.casefold())
        result[candidate] = frame.reset_index(drop=True).copy(deep=True)

    if by is not None:
        group_columns = _normalise_columns(by, argument="by")
        _validate_columns(df, group_columns, argument="by")
        grouper: str | list[str] = group_columns[0] if len(group_columns) == 1 else group_columns
        for key, group in df.copy(deep=True).groupby(
            grouper, dropna=False, sort=False, observed=True
        ):
            if drop_group_columns:
                group = group.drop(columns=group_columns)
            add(_split_label(key), group)
        if not result:
            add("空数据", df.drop(columns=group_columns) if drop_group_columns else df)
        details: dict[str, Any] = {"by": group_columns, "part_count": len(result)}
    else:
        if not isinstance(rows_per_table, int) or rows_per_table <= 0:
            raise ValueError("rows_per_table 必须是正整数")
        for index, start in enumerate(range(0, len(df), rows_per_table), start=1):
            add(f"第{index}部分", df.iloc[start : start + rows_per_table])
        # An empty input still yields one empty table, which is easier for UIs.
        if not result:
            add("第1部分", df)
        details = {"rows_per_table": rows_per_table, "part_count": len(result)}
    _log(
        log,
        "拆分数据表",
        input_tables=[table_name],
        output_tables=list(result),
        details=details,
    )
    return result


def _mask_value(
    value: Any,
    strategy: MaskStrategy,
    *,
    salt: str,
    mask_char: str,
    keep_start: int,
    keep_end: int,
) -> Any:
    if pd.isna(value):
        return value
    text = str(value)
    if strategy == "hash":
        return hashlib.sha256(f"{salt}{text}".encode("utf-8")).hexdigest()[:16]
    if strategy == "full":
        return mask_char * max(3, len(text))
    if strategy == "phone":
        digits = re.sub(r"\D", "", text)
        if len(digits) >= 7:
            return f"{digits[:3]}{mask_char * (len(digits) - 7)}{digits[-4:]}"
    if strategy == "email" and "@" in text:
        local, domain = text.rsplit("@", 1)
        return f"{local[:1]}{mask_char * max(3, len(local) - 1)}@{domain}"
    if strategy == "name":
        return text[:1] + mask_char * max(1, len(text) - 1)
    if strategy == "id" and len(text) > 8:
        return f"{text[:4]}{mask_char * (len(text) - 8)}{text[-4:]}"
    if len(text) <= keep_start + keep_end:
        return mask_char * max(1, len(text))
    tail = text[-keep_end:] if keep_end else ""
    return f"{text[:keep_start]}{mask_char * (len(text) - keep_start - keep_end)}{tail}"


def mask_columns(
    df: pd.DataFrame,
    columns: Mapping[str, MaskStrategy] | Sequence[str],
    *,
    strategy: MaskStrategy = "partial",
    salt: str = "",
    mask_char: str = "*",
    keep_start: int = 1,
    keep_end: int = 1,
    table_name: str = "数据表",
    output_name: str | None = None,
    log: OperationLog | None = None,
) -> pd.DataFrame:
    """Mask sensitive columns using partial, full, hash, phone, email, name or ID rules.

    Returns:
        A masked copy. Null values stay null, and the original frame is unchanged.
    """

    _require_dataframe(df)
    if len(mask_char) != 1:
        raise ValueError("mask_char 必须是单个字符")
    if keep_start < 0 or keep_end < 0:
        raise ValueError("keep_start 和 keep_end 不能为负数")
    if isinstance(columns, Mapping):
        rules = dict(columns)
    else:
        rules = {column: strategy for column in columns}
    if not rules:
        raise ValueError("columns 不能为空")
    _validate_columns(df, list(rules), argument="columns")
    allowed = {"partial", "full", "hash", "phone", "email", "name", "id"}
    invalid = {rule for rule in rules.values() if rule not in allowed}
    if invalid:
        raise ValueError(f"不支持的脱敏策略：{sorted(invalid)}")
    result = df.copy(deep=True)
    for column, rule in rules.items():
        result[column] = result[column].map(
            lambda value, selected=rule: _mask_value(
                value,
                selected,
                salt=salt,
                mask_char=mask_char,
                keep_start=keep_start,
                keep_end=keep_end,
            )
        )
    _log(
        log,
        "数据脱敏",
        input_tables=[table_name],
        output_tables=[output_name or table_name],
        details={"rules": rules, "row_count": len(result)},
    )
    return result


def export_tables(
    tables: Mapping[str, pd.DataFrame],
    output_path: str | Path,
    *,
    operation_log: OperationLog | Sequence[OperationRecord] | None = None,
    include_log: bool = True,
    index: bool = False,
    csv_encoding: str = "utf-8-sig",
    escape_formulas: bool = True,
    overwrite: bool = False,
) -> ExportResult:
    """Export tables to XLSX, one CSV, or a ZIP containing one CSV per table.

    XLSX/ZIP embeds the operation log. A single CSV gets an adjacent
    ``*_操作日志.csv``. Potential spreadsheet-formula text is escaped by
    default. Returns :class:`ExportResult` with all created paths.
    """

    return export_tables_to_path(
        tables,
        output_path,
        operation_log=operation_log,
        include_log=include_log,
        index=index,
        csv_encoding=csv_encoding,
        escape_formulas=escape_formulas,
        overwrite=overwrite,
    )


__all__ = [
    "CleaningConfig",
    "CleaningReport",
    "DataFrameProfile",
    "ExportResult",
    "OperationLog",
    "concat_tables",
    "export_tables",
    "group_summary",
    "join_tables",
    "load_tables",
    "lookup_match",
    "mask_columns",
    "profile_dataframe",
    "select_rename_sort",
    "smart_clean",
    "split_dataframe",
]
