"""Deterministic post-export verification for customer deliverables.

The verifier never trusts a successful writer return value.  It reopens the
artifact, compares its shape and stable aggregate fingerprints with the source
tables, and produces both machine-readable and customer-readable reports.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
import hashlib
import json
import math
from pathlib import Path
import zipfile
from typing import Any, Mapping

import pandas as pd


FORMULA_PREFIXES = ("=", "+", "-", "@")


@dataclass(frozen=True)
class TableAcceptance:
    table: str
    status: str
    expected_rows: int
    actual_rows: int
    expected_columns: int
    actual_columns: int
    missing_cells: int
    duplicate_rows: int
    numeric_totals_match: bool
    content_fingerprint_match: bool
    notes: tuple[str, ...]


@dataclass(frozen=True)
class DeliveryAcceptance:
    status: str
    artifact: str
    artifact_sha256: str
    checked_at: str
    checks_passed: int
    checks_total: int
    tables: tuple[TableAcceptance, ...]
    warnings: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _normalise_scalar(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return None
    if isinstance(value, str):
        if not value.strip():
            return None
        if value.strip().casefold() in {"nan", "inf", "+inf", "-inf", "infinity", "+infinity", "-infinity"}:
            return None
        if len(value) > 1 and value[0] == "'" and value[1] in FORMULA_PREFIXES:
            value = value[1:]
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        # XLSX stores IEEE-754 values through decimal text and may differ in
        # the final machine-precision digits after reopening. Aggregate totals
        # are checked separately with strict tolerance; 12 significant digits
        # keep the row fingerprint stable without hiding business differences.
        return format(value, ".12g")
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        try:
            return _normalise_scalar(value.item())
        except (TypeError, ValueError):
            pass
    return value if isinstance(value, (str, bool)) else str(value)


def dataframe_fingerprint(frame: pd.DataFrame) -> str:
    """Return a stable value fingerprint insensitive to pandas dtype drift."""

    digest = hashlib.sha256()
    digest.update(json.dumps([str(column) for column in frame.columns], ensure_ascii=False).encode("utf-8"))
    for row in frame.itertuples(index=False, name=None):
        values = [_normalise_scalar(value) for value in row]
        digest.update(json.dumps(values, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def _blank_mask(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.apply(
        lambda column: column.map(
            lambda value: _normalise_scalar(value) is None
        )
    )


def _trim_trailing_blank_rows(frame: pd.DataFrame) -> pd.DataFrame:
    """Excel cannot persist trailing rows made entirely of empty strings."""

    result = frame.reset_index(drop=True).copy()
    if result.empty:
        return result
    meaningful = ~_blank_mask(result).all(axis=1)
    if not bool(meaningful.any()):
        return result.iloc[0:0].copy()
    last = int(meaningful[meaningful].index[-1])
    return result.iloc[: last + 1].reset_index(drop=True)


def _numeric_totals(frame: pd.DataFrame, columns: list[str] | None = None) -> dict[str, float]:
    totals: dict[str, float] = {}
    selected = columns if columns is not None else [str(column) for column in frame.columns]
    for column in selected:
        if column not in frame.columns:
            continue
        raw = frame[column]
        if pd.api.types.is_datetime64_any_dtype(raw.dtype):
            continue
        non_null = raw.dropna()
        if not non_null.empty and float(non_null.map(lambda value: isinstance(value, (pd.Timestamp, datetime))).mean()) >= 0.8:
            continue
        series = pd.to_numeric(raw, errors="coerce")
        numeric_count = int(series.notna().sum())
        if numeric_count == 0:
            continue
        if columns is None and not pd.api.types.is_numeric_dtype(raw.dtype):
            denominator = max(1, int(raw.notna().sum()))
            if numeric_count / denominator < 0.9:
                continue
        totals[str(column)] = float(series.sum(skipna=True))
    return totals


def _totals_match(left: Mapping[str, float], right: Mapping[str, float]) -> bool:
    if set(left) != set(right):
        return False
    return all(math.isclose(left[key], right[key], rel_tol=1e-9, abs_tol=1e-7) for key in left)


def _read_xlsx(
    path: Path,
    expected_tables: Mapping[str, pd.DataFrame] | None = None,
) -> dict[str, pd.DataFrame]:
    """Read data regions from both plain and professional-layout workbooks.

    Professional reports reserve three rows for titles and thirty-nine rows on
    dashboard sheets.  Reading every sheet with ``header=0`` therefore treats
    the presentation title as a table header and produces false data-loss
    alarms.  The verifier locates the exact expected header, then reads only
    the declared data rectangle; charts and KPI cards remain outside it.
    """

    if not expected_tables:
        return {
            str(name): frame
            for name, frame in pd.read_excel(path, sheet_name=None, dtype=object).items()
        }
    workbook = pd.ExcelFile(path)
    available = set(map(str, workbook.sheet_names))
    tables: dict[str, pd.DataFrame] = {}
    for raw_name, expected in expected_tables.items():
        name = str(raw_name)
        if name not in available:
            continue
        raw = pd.read_excel(workbook, sheet_name=name, header=None, dtype=object)
        expected_columns = [str(column) for column in expected.columns]
        header_index: int | None = None
        maximum_scan = min(len(raw), 100)
        for position in range(maximum_scan):
            values = [str(value) if pd.notna(value) else "" for value in raw.iloc[position, : len(expected_columns)].tolist()]
            if values == expected_columns:
                header_index = position
                break
        if header_index is None:
            # Preserve a useful failed comparison rather than hiding a missing
            # or corrupted header behind an exception.
            tables[name] = pd.read_excel(workbook, sheet_name=name, dtype=object)
            continue
        start = header_index + 1
        data = raw.iloc[start : start + len(expected), : len(expected_columns)].copy()
        data.columns = expected_columns
        data = data.reset_index(drop=True)
        tables[name] = data
    return tables


def _read_zip(path: Path) -> dict[str, pd.DataFrame]:
    tables: dict[str, pd.DataFrame] = {}
    with zipfile.ZipFile(path) as archive:
        for member in archive.namelist():
            if member.lower().endswith(".csv") and not member.endswith("/"):
                with archive.open(member) as stream:
                    tables[Path(member).stem] = pd.read_csv(stream, dtype=object)
    return tables


def verify_delivery(
    artifact: str | Path,
    expected_tables: Mapping[str, pd.DataFrame],
    *,
    allow_extra_tables: bool = True,
) -> DeliveryAcceptance:
    """Reopen and verify an XLSX/CSV/ZIP deliverable.

    A formula escaped by a leading apostrophe is treated as the same inert text
    for fingerprint purposes.  This keeps security escaping from becoming a
    false data-loss report.
    """

    path = Path(artifact).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"交付文件不存在：{path}")
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        actual_tables = _read_xlsx(path, expected_tables)
    elif suffix == ".zip":
        actual_tables = _read_zip(path)
    elif suffix == ".csv":
        actual_tables = {path.stem: pd.read_csv(path, dtype=object)}
    else:
        raise ValueError("自动验收仅支持 .xlsx、.csv 或 .zip")

    warnings: list[str] = []
    results: list[TableAcceptance] = []
    passed = 0
    total = 0

    for expected_name, expected in expected_tables.items():
        total += 5
        actual = actual_tables.get(str(expected_name))
        if actual is None:
            results.append(
                TableAcceptance(
                    str(expected_name), "failed", len(expected), 0,
                    len(expected.columns), 0, 0, 0, False, False,
                    ("导出文件缺少该数据表",),
                )
            )
            continue
        expected_copy = _trim_trailing_blank_rows(expected)
        actual_copy = _trim_trailing_blank_rows(actual)
        expected_copy.columns = [str(column) for column in expected_copy.columns]
        actual_copy.columns = [str(column) for column in actual_copy.columns]
        row_ok = len(expected_copy) == len(actual_copy)
        column_ok = list(expected_copy.columns) == list(actual_copy.columns)
        expected_missing = int(_blank_mask(expected_copy).sum().sum())
        actual_missing = int(_blank_mask(actual_copy).sum().sum())
        null_ok = expected_missing == actual_missing
        expected_totals = _numeric_totals(expected_copy)
        totals_ok = column_ok and _totals_match(
            expected_totals,
            _numeric_totals(actual_copy, columns=list(expected_totals)),
        )
        fingerprint_ok = column_ok and dataframe_fingerprint(expected_copy) == dataframe_fingerprint(actual_copy)
        checks = (row_ok, column_ok, null_ok, totals_ok, fingerprint_ok)
        passed += sum(bool(item) for item in checks)
        notes: list[str] = []
        if not row_ok:
            notes.append("行数不一致")
        if not column_ok:
            notes.append("字段名称或顺序不一致")
        if not null_ok:
            notes.append("空值数量不一致")
        if not totals_ok:
            notes.append("数值汇总不一致")
        if not fingerprint_ok:
            notes.append("逐行内容指纹不一致（可能仅为日期或数字显示格式变化）")
        results.append(
            TableAcceptance(
                table=str(expected_name),
                status="passed" if all(checks) else "failed",
                expected_rows=len(expected_copy),
                actual_rows=len(actual_copy),
                expected_columns=len(expected_copy.columns),
                actual_columns=len(actual_copy.columns),
                missing_cells=actual_missing,
                duplicate_rows=int(actual_copy.duplicated().sum()),
                numeric_totals_match=totals_ok,
                content_fingerprint_match=fingerprint_ok,
                notes=tuple(notes),
            )
        )

    expected_names = {str(name) for name in expected_tables}
    extras = sorted(set(actual_tables) - expected_names)
    if extras and not allow_extra_tables:
        warnings.append(f"发现未声明的数据表：{'、'.join(extras)}")
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    status = "passed" if all(item.status == "passed" for item in results) else "failed"
    return DeliveryAcceptance(
        status=status,
        artifact=path.name,
        artifact_sha256=digest,
        checked_at=datetime.now().isoformat(timespec="seconds"),
        checks_passed=passed,
        checks_total=total,
        tables=tuple(results),
        warnings=tuple(warnings),
    )


def acceptance_frame(report: DeliveryAcceptance) -> pd.DataFrame:
    rows = []
    for item in report.tables:
        rows.append(
            {
                "数据表": item.table,
                "验收状态": "通过" if item.status == "passed" else "失败",
                "预期行数": item.expected_rows,
                "实际行数": item.actual_rows,
                "预期列数": item.expected_columns,
                "实际列数": item.actual_columns,
                "空值单元格": item.missing_cells,
                "重复行": item.duplicate_rows,
                "数值合计一致": "是" if item.numeric_totals_match else "否",
                "内容指纹一致": "是" if item.content_fingerprint_match else "否",
                "说明": "；".join(item.notes),
            }
        )
    return pd.DataFrame(rows)


def write_acceptance_json(report: DeliveryAcceptance, destination: str | Path) -> Path:
    path = Path(destination)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)
    return path


__all__ = [
    "DeliveryAcceptance",
    "TableAcceptance",
    "acceptance_frame",
    "dataframe_fingerprint",
    "verify_delivery",
    "write_acceptance_json",
]
