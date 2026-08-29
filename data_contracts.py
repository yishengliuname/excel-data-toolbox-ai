"""Versioned, local data contracts for repeatable spreadsheet work.

The contract layer deliberately stays deterministic.  AI may propose a
contract, but only a locally validated contract can be activated and used as a
delivery gate.  Contracts contain schema and aggregate metadata only; they do
not persist customer cell values.
"""

from __future__ import annotations

import hashlib
import json
import math
import re
import tempfile
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

CONTRACT_SCHEMA_VERSION = 1
_IDENTIFIER_HINT = re.compile(r"(^|[_\s])(id|code|no|number|key)($|[_\s])|编号|编码|单号|序号|工号|账号", re.I)
_DATE_HINT = re.compile(r"日期|时间|月份|年度|date|time|month|year", re.I)
_RATE_HINT = re.compile(r"率|比例|占比|百分比|满意度|评分|得分|rate|ratio|percent|score", re.I)
_MONEY_HINT = re.compile(r"金额|销售额|成本|利润|收入|支出|预算|单价|余额|应收|应付|amount|revenue|cost|profit|price", re.I)
_SAFE_FIELD = re.compile(r"^[^\x00-\x1f]{1,200}$")


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _json_scalar(value: Any) -> Any:
    if value is None or value is pd.NA:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.isoformat()
    return str(value)[:500]


def _semantic_type(series: pd.Series, name: str) -> str:
    non_null = series.dropna()
    if pd.api.types.is_bool_dtype(series):
        return "boolean"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "datetime"
    if pd.api.types.is_integer_dtype(series):
        return "integer"
    if pd.api.types.is_numeric_dtype(series):
        return "number"
    if _DATE_HINT.search(name) and len(non_null):
        parsed = pd.to_datetime(non_null, errors="coerce")
        if float(parsed.notna().mean()) >= 0.85:
            return "datetime"
    if len(non_null):
        numeric = pd.to_numeric(non_null.astype("string").str.replace(",", "", regex=False).str.replace("%", "", regex=False), errors="coerce")
        if float(numeric.notna().mean()) >= 0.9 and not _IDENTIFIER_HINT.search(name):
            return "number"
    return "string"


def _type_compatible(series: pd.Series, expected: str, name: str) -> tuple[bool, float]:
    non_null = series.dropna()
    if not len(non_null):
        return True, 1.0
    actual = _semantic_type(series, name)
    if actual == expected or {actual, expected} <= {"integer", "number"}:
        return True, 1.0
    if expected in {"integer", "number"}:
        parsed = pd.to_numeric(non_null.astype("string").str.replace(",", "", regex=False).str.replace("%", "", regex=False), errors="coerce")
        rate = float(parsed.notna().mean())
        return rate >= 0.95, rate
    if expected == "datetime":
        parsed = pd.to_datetime(non_null, errors="coerce")
        rate = float(parsed.notna().mean())
        return rate >= 0.95, rate
    return False, 0.0


@dataclass(frozen=True)
class ColumnContract:
    name: str
    semantic_type: str
    required: bool = True
    nullable: bool = True
    unique: bool = False
    allowed_values: tuple[Any, ...] = ()
    minimum: float | None = None
    maximum: float | None = None
    pattern: str | None = None
    role: str = ""
    description: str = ""

    def __post_init__(self) -> None:
        if not _SAFE_FIELD.match(str(self.name)):
            raise ValueError("合同字段名无效")
        if self.semantic_type not in {"string", "integer", "number", "boolean", "datetime"}:
            raise ValueError(f"字段 {self.name} 的类型无效")
        if self.minimum is not None and self.maximum is not None and self.minimum > self.maximum:
            raise ValueError(f"字段 {self.name} 的最小值不能大于最大值")
        if self.pattern:
            re.compile(self.pattern)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["allowed_values"] = list(self.allowed_values)
        return payload


@dataclass(frozen=True)
class DataContract:
    contract_id: str
    name: str
    version: int
    created_at: str
    columns: tuple[ColumnContract, ...]
    allow_extra_columns: bool = True
    minimum_rows: int = 1
    maximum_rows: int | None = None
    key_columns: tuple[str, ...] = ()
    cross_field_rules: tuple[Mapping[str, Any], ...] = ()
    source_fingerprint: str = ""
    schema_version: int = CONTRACT_SCHEMA_VERSION
    notes: str = ""

    def __post_init__(self) -> None:
        names = [column.name for column in self.columns]
        if not self.name.strip() or len(self.name) > 120:
            raise ValueError("合同名称无效")
        if len(names) != len(set(names)) or not names:
            raise ValueError("合同字段必须非空且不能重复")
        if not set(self.key_columns).issubset(names):
            raise ValueError("合同主键必须存在于字段定义中")
        if self.minimum_rows < 0 or (self.maximum_rows is not None and self.maximum_rows < self.minimum_rows):
            raise ValueError("合同数据行数范围无效")
        for rule in self.cross_field_rules:
            _validate_cross_rule(rule, set(names))

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "contract_id": self.contract_id,
            "name": self.name,
            "version": self.version,
            "created_at": self.created_at,
            "allow_extra_columns": self.allow_extra_columns,
            "minimum_rows": self.minimum_rows,
            "maximum_rows": self.maximum_rows,
            "key_columns": list(self.key_columns),
            "source_fingerprint": self.source_fingerprint,
            "notes": self.notes,
            "columns": [column.to_dict() for column in self.columns],
            "cross_field_rules": [dict(rule) for rule in self.cross_field_rules],
        }

    @classmethod
    def from_dict(cls, payload: Mapping[str, Any]) -> "DataContract":
        if int(payload.get("schema_version", 0)) != CONTRACT_SCHEMA_VERSION:
            raise ValueError("数据合同版本不受支持")
        raw_columns = payload.get("columns")
        if not isinstance(raw_columns, list):
            raise ValueError("数据合同 columns 必须是数组")
        columns = []
        for item in raw_columns:
            if not isinstance(item, Mapping):
                raise ValueError("数据合同字段定义无效")
            columns.append(ColumnContract(
                name=str(item.get("name") or ""),
                semantic_type=str(item.get("semantic_type") or ""),
                required=bool(item.get("required", True)),
                nullable=bool(item.get("nullable", True)),
                unique=bool(item.get("unique", False)),
                allowed_values=tuple(item.get("allowed_values") or ()),
                minimum=float(item["minimum"]) if item.get("minimum") is not None else None,
                maximum=float(item["maximum"]) if item.get("maximum") is not None else None,
                pattern=str(item["pattern"]) if item.get("pattern") else None,
                role=str(item.get("role") or "")[:80],
                description=str(item.get("description") or "")[:500],
            ))
        rules = payload.get("cross_field_rules") or []
        if not isinstance(rules, list) or any(not isinstance(item, Mapping) for item in rules):
            raise ValueError("数据合同跨字段规则无效")
        return cls(
            contract_id=str(payload.get("contract_id") or uuid.uuid4().hex[:16]),
            name=str(payload.get("name") or "未命名数据合同"),
            version=int(payload.get("version") or 1),
            created_at=str(payload.get("created_at") or _now()),
            columns=tuple(columns),
            allow_extra_columns=bool(payload.get("allow_extra_columns", True)),
            minimum_rows=int(payload.get("minimum_rows") or 0),
            maximum_rows=int(payload["maximum_rows"]) if payload.get("maximum_rows") is not None else None,
            key_columns=tuple(str(item) for item in (payload.get("key_columns") or ())),
            cross_field_rules=tuple(dict(item) for item in rules),
            source_fingerprint=str(payload.get("source_fingerprint") or ""),
            schema_version=int(payload.get("schema_version") or CONTRACT_SCHEMA_VERSION),
            notes=str(payload.get("notes") or "")[:2000],
        )


@dataclass(frozen=True)
class ContractIssue:
    severity: str
    code: str
    field: str
    message: str
    failed_count: int = 0
    sample_rows: tuple[int, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {**asdict(self), "sample_rows": list(self.sample_rows)}


@dataclass(frozen=True)
class ContractValidationResult:
    contract_id: str
    contract_version: int
    passed: bool
    checked_at: str
    row_count: int
    issues: tuple[ContractIssue, ...]
    schema_drift: Mapping[str, Any]
    metrics: Mapping[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "contract_id": self.contract_id,
            "contract_version": self.contract_version,
            "passed": self.passed,
            "checked_at": self.checked_at,
            "row_count": self.row_count,
            "issues": [item.to_dict() for item in self.issues],
            "schema_drift": dict(self.schema_drift),
            "metrics": dict(self.metrics),
        }


def dataframe_schema_fingerprint(frame: pd.DataFrame) -> str:
    payload = {
        "columns": [
            {"name": str(column), "type": _semantic_type(frame[column], str(column))}
            for column in frame.columns
        ],
        "rows": int(len(frame)),
    }
    raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def infer_data_contract(
    frame: pd.DataFrame,
    *,
    name: str,
    contract_id: str | None = None,
    version: int = 1,
    allow_extra_columns: bool = True,
    strict_nulls: bool = False,
) -> DataContract:
    if not isinstance(frame, pd.DataFrame) or not len(frame.columns):
        raise ValueError("只能从包含字段的数据表生成合同")
    columns: list[ColumnContract] = []
    keys: list[str] = []
    for raw_name in frame.columns:
        column_name = str(raw_name)
        series = frame[raw_name]
        non_null = series.dropna()
        semantic = _semantic_type(series, column_name)
        uniqueness = float(non_null.nunique(dropna=True) / len(non_null)) if len(non_null) else 0.0
        is_identifier = bool(_IDENTIFIER_HINT.search(column_name))
        unique = bool(len(non_null) and uniqueness == 1.0 and (is_identifier or len(non_null) >= 5))
        if unique and is_identifier:
            keys.append(column_name)
        allowed: tuple[Any, ...] = ()
        distinct = int(non_null.nunique(dropna=True))
        if semantic in {"string", "boolean"} and 0 < distinct <= min(30, max(8, int(len(non_null) * 0.05))):
            allowed = tuple(_json_scalar(item) for item in non_null.drop_duplicates().head(30).tolist())
        minimum = maximum = None
        if semantic in {"integer", "number"} and len(non_null):
            numeric = pd.to_numeric(non_null, errors="coerce").dropna()
            if len(numeric):
                minimum = float(numeric.min())
                maximum = float(numeric.max())
        role = (
            "identifier" if is_identifier else
            "date" if semantic == "datetime" else
            "rate" if _RATE_HINT.search(column_name) else
            "money" if _MONEY_HINT.search(column_name) else
            "measure" if semantic in {"integer", "number"} else
            "dimension"
        )
        columns.append(ColumnContract(
            name=column_name,
            semantic_type=semantic,
            required=True,
            nullable=not strict_nulls and bool(series.isna().any()),
            unique=unique,
            allowed_values=allowed,
            minimum=minimum,
            maximum=maximum,
            role=role,
        ))
    return DataContract(
        contract_id=contract_id or uuid.uuid4().hex[:16],
        name=name.strip()[:120] or "数据合同",
        version=max(1, int(version)),
        created_at=_now(),
        columns=tuple(columns),
        allow_extra_columns=allow_extra_columns,
        minimum_rows=1 if len(frame) else 0,
        key_columns=tuple(keys[:4]),
        source_fingerprint=dataframe_schema_fingerprint(frame),
        notes="由已确认样例的字段结构和数据分布自动生成；业务阈值需人工确认后补充。",
    )


def _validate_cross_rule(rule: Mapping[str, Any], columns: set[str]) -> None:
    kind = str(rule.get("type") or "")
    if kind not in {"compare", "formula", "balance"}:
        raise ValueError("跨字段规则类型只允许 compare、formula、balance")
    fields = [str(item) for item in (rule.get("fields") or ())]
    if not fields or not set(fields).issubset(columns) or len(fields) > 8:
        raise ValueError("跨字段规则引用了无效字段")
    if kind == "compare" and str(rule.get("operator") or "") not in {"eq", "ne", "gt", "ge", "lt", "le"}:
        raise ValueError("compare 规则运算符无效")
    if kind in {"formula", "balance"}:
        coefficients = rule.get("coefficients")
        if not isinstance(coefficients, Mapping) or set(map(str, coefficients)) != set(fields):
            raise ValueError("公式规则必须为每个字段提供系数")
        for value in coefficients.values():
            numeric = float(value)
            if not math.isfinite(numeric) or abs(numeric) > 1_000_000:
                raise ValueError("公式规则系数无效")


def _failed_rows(mask: pd.Series, maximum: int = 20) -> tuple[int, tuple[int, ...]]:
    failed = mask.fillna(False).astype(bool)
    indexes = [int(position) + 2 for position in range(len(failed)) if bool(failed.iloc[position])][:maximum]
    return int(failed.sum()), tuple(indexes)


def validate_contract(frame: pd.DataFrame, contract: DataContract) -> ContractValidationResult:
    if not isinstance(frame, pd.DataFrame):
        raise TypeError("contract validation requires a DataFrame")
    actual = [str(column) for column in frame.columns]
    expected = [column.name for column in contract.columns]
    missing = [name for name in expected if name not in actual]
    extra = [name for name in actual if name not in expected]
    issues: list[ContractIssue] = []
    if missing:
        for name in missing:
            column = next(item for item in contract.columns if item.name == name)
            if column.required:
                issues.append(ContractIssue("error", "missing_column", name, f"缺少必需字段：{name}"))
    if extra and not contract.allow_extra_columns:
        issues.append(ContractIssue("error", "extra_columns", "", f"出现合同外字段：{', '.join(extra[:20])}", len(extra)))
    if len(frame) < contract.minimum_rows:
        issues.append(ContractIssue("error", "row_count_low", "", f"数据行数 {len(frame)} 低于合同下限 {contract.minimum_rows}"))
    if contract.maximum_rows is not None and len(frame) > contract.maximum_rows:
        issues.append(ContractIssue("error", "row_count_high", "", f"数据行数 {len(frame)} 高于合同上限 {contract.maximum_rows}"))

    for column in contract.columns:
        if column.name not in frame.columns:
            continue
        series = frame[column.name]
        compatible, parse_rate = _type_compatible(series, column.semantic_type, column.name)
        if not compatible:
            issues.append(ContractIssue("error", "type_mismatch", column.name, f"字段类型应为 {column.semantic_type}，可解析率仅 {parse_rate:.1%}"))
        if not column.nullable:
            count, rows = _failed_rows(series.isna() | series.astype("string").str.strip().eq(""))
            if count:
                issues.append(ContractIssue("error", "missing_values", column.name, f"字段不允许为空，共 {count} 行", count, rows))
        if column.unique:
            duplicate_mask = series.notna() & series.duplicated(keep=False)
            count, rows = _failed_rows(duplicate_mask)
            if count:
                issues.append(ContractIssue("error", "duplicate_values", column.name, f"字段应唯一，共 {count} 行重复", count, rows))
        if column.allowed_values:
            allowed = {str(item) for item in column.allowed_values}
            invalid = series.notna() & ~series.astype("string").isin(allowed)
            count, rows = _failed_rows(invalid)
            if count:
                issues.append(ContractIssue("error", "invalid_value", column.name, f"出现合同允许范围外的值，共 {count} 行", count, rows))
        if column.minimum is not None or column.maximum is not None:
            numeric = pd.to_numeric(series, errors="coerce")
            invalid = pd.Series(False, index=series.index)
            if column.minimum is not None:
                invalid |= numeric.notna() & numeric.lt(column.minimum)
            if column.maximum is not None:
                invalid |= numeric.notna() & numeric.gt(column.maximum)
            count, rows = _failed_rows(invalid)
            if count:
                issues.append(ContractIssue("warning", "range_drift", column.name, f"数值超出样例范围，共 {count} 行；需确认是正常增长还是异常", count, rows))
        if column.pattern:
            invalid = series.notna() & ~series.astype("string").str.fullmatch(column.pattern, na=False)
            count, rows = _failed_rows(invalid)
            if count:
                issues.append(ContractIssue("error", "pattern_mismatch", column.name, f"格式不符合合同规则，共 {count} 行", count, rows))

    for rule_index, rule in enumerate(contract.cross_field_rules, start=1):
        fields = [str(item) for item in rule.get("fields") or ()]
        if any(field not in frame.columns for field in fields):
            continue
        kind = str(rule.get("type"))
        tolerance = float(rule.get("tolerance") or 0.0)
        label = str(rule.get("name") or f"跨字段规则{rule_index}")[:120]
        if kind == "compare":
            left, right = frame[fields[0]], frame[fields[1]]
            operator = str(rule.get("operator"))
            comparisons = {"eq": left.eq(right), "ne": left.ne(right), "gt": left.gt(right), "ge": left.ge(right), "lt": left.lt(right), "le": left.le(right)}
            invalid = ~(comparisons[operator] | left.isna() | right.isna())
        else:
            total = pd.Series(0.0, index=frame.index)
            present = pd.Series(False, index=frame.index)
            for field_name, coefficient in (rule.get("coefficients") or {}).items():
                numeric = pd.to_numeric(frame[str(field_name)], errors="coerce")
                total = total.add(numeric.fillna(0) * float(coefficient), fill_value=0)
                present |= numeric.notna()
            expected_value = float(rule.get("expected") or 0.0)
            invalid = present & total.sub(expected_value).abs().gt(tolerance)
        count, rows = _failed_rows(invalid)
        if count:
            issues.append(ContractIssue("error", "cross_field_rule", ",".join(fields), f"{label}未通过，共 {count} 行", count, rows))

    errors = sum(1 for item in issues if item.severity == "error")
    warnings = sum(1 for item in issues if item.severity == "warning")
    return ContractValidationResult(
        contract_id=contract.contract_id,
        contract_version=contract.version,
        passed=errors == 0,
        checked_at=_now(),
        row_count=len(frame),
        issues=tuple(issues),
        schema_drift={"missing_columns": missing, "extra_columns": extra, "reordered": [name for name in actual if name in expected] != [name for name in expected if name in actual]},
        metrics={"error_count": errors, "warning_count": warnings, "column_count": len(actual), "schema_fingerprint": dataframe_schema_fingerprint(frame)},
    )


class DataContractStore:
    """Crash-safe versioned JSON store with one active version per contract."""

    def __init__(self, root: str | Path) -> None:
        self.root = Path(root)
        self.root.mkdir(parents=True, exist_ok=True)

    def _directory(self, contract_id: str) -> Path:
        if not re.fullmatch(r"[A-Za-z0-9_-]{8,64}", contract_id):
            raise ValueError("合同编号无效")
        return self.root / contract_id

    def save(self, contract: DataContract, *, activate: bool = True) -> Path:
        directory = self._directory(contract.contract_id)
        directory.mkdir(parents=True, exist_ok=True)
        destination = directory / f"v{contract.version}.json"
        raw = json.dumps(contract.to_dict(), ensure_ascii=False, indent=2, allow_nan=False).encode("utf-8")
        with tempfile.NamedTemporaryFile(dir=directory, prefix=".contract_", suffix=".tmp", delete=False) as stream:
            stream.write(raw)
            temporary = Path(stream.name)
        temporary.replace(destination)
        if activate:
            (directory / "active.txt").write_text(str(contract.version), encoding="ascii")
        return destination

    def load(self, contract_id: str, version: int | None = None) -> DataContract:
        directory = self._directory(contract_id)
        if version is None:
            try:
                version = int((directory / "active.txt").read_text(encoding="ascii").strip())
            except (OSError, ValueError) as exc:
                raise FileNotFoundError("数据合同没有启用版本") from exc
        path = directory / f"v{int(version)}.json"
        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, Mapping):
            raise ValueError("数据合同文件无效")
        return DataContract.from_dict(payload)

    def next_version(self, contract_id: str) -> int:
        directory = self._directory(contract_id)
        versions = []
        if directory.exists():
            for path in directory.glob("v*.json"):
                try:
                    versions.append(int(path.stem[1:]))
                except ValueError:
                    continue
        return max(versions, default=0) + 1

    def list(self) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        for directory in self.root.iterdir() if self.root.exists() else ():
            if not directory.is_dir():
                continue
            try:
                contract = self.load(directory.name)
            except (OSError, ValueError, json.JSONDecodeError):
                continue
            versions = sorted(int(path.stem[1:]) for path in directory.glob("v*.json") if path.stem[1:].isdigit())
            results.append({"contract_id": contract.contract_id, "name": contract.name, "active_version": contract.version, "versions": versions, "columns": len(contract.columns), "created_at": contract.created_at})
        return sorted(results, key=lambda item: (str(item["name"]), str(item["contract_id"])))


def issues_frame(result: ContractValidationResult) -> pd.DataFrame:
    rows = []
    for issue in result.issues:
        rows.append({
            "严重程度": issue.severity,
            "问题代码": issue.code,
            "字段": issue.field,
            "说明": issue.message,
            "失败行数": issue.failed_count,
            "样例Excel行号": ", ".join(map(str, issue.sample_rows)),
        })
    return pd.DataFrame(rows, columns=["严重程度", "问题代码", "字段", "说明", "失败行数", "样例Excel行号"])


__all__ = [
    "CONTRACT_SCHEMA_VERSION", "ColumnContract", "ContractIssue",
    "ContractValidationResult", "DataContract", "DataContractStore",
    "dataframe_schema_fingerprint", "infer_data_contract", "issues_frame",
    "validate_contract",
]
