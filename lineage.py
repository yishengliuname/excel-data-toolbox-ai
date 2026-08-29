"""Local operation lineage and delivery evidence.

The store records metadata, fingerprints and aggregate facts only.  It avoids
persisting customer cell values or prompts so an evidence package can be safely
handed to a customer or support engineer.
"""

from __future__ import annotations

import hashlib
import json
import math
import sqlite3
import uuid
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping

import pandas as pd

LINEAGE_SCHEMA_VERSION = 1


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="milliseconds")


def _safe_json(value: Any, *, depth: int = 0) -> Any:
    if depth > 6:
        return "[depth-limit]"
    if value is None or value is pd.NA:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, str):
        return value[:1000]
    if isinstance(value, Mapping):
        return {str(key)[:120]: _safe_json(item, depth=depth + 1) for key, item in list(value.items())[:100]}
    if isinstance(value, (list, tuple, set)):
        return [_safe_json(item, depth=depth + 1) for item in list(value)[:100]]
    return str(value)[:1000]


def dataframe_fingerprint(frame: pd.DataFrame) -> str:
    """Stable content fingerprint without storing the content itself."""

    digest = hashlib.sha256()
    digest.update(str(tuple(map(str, frame.columns))).encode("utf-8"))
    digest.update(str(tuple(map(str, frame.dtypes))).encode("utf-8"))
    digest.update(str(frame.shape).encode("ascii"))
    if len(frame):
        try:
            hashed = pd.util.hash_pandas_object(frame, index=True, categorize=True)
            digest.update(hashed.to_numpy().tobytes())
        except (TypeError, ValueError):
            safe = frame.astype("string").fillna("<NA>")
            digest.update(pd.util.hash_pandas_object(safe, index=True).to_numpy().tobytes())
    return digest.hexdigest()


def dataset_metadata(name: str, frame: pd.DataFrame, *, source: str = "") -> dict[str, Any]:
    numeric_totals: dict[str, float] = {}
    for column in frame.columns:
        if pd.api.types.is_numeric_dtype(frame[column]):
            value = pd.to_numeric(frame[column], errors="coerce").sum(min_count=1)
            if pd.notna(value) and math.isfinite(float(value)):
                numeric_totals[str(column)[:120]] = round(float(value), 8)
        if len(numeric_totals) >= 30:
            break
    return {
        "name": str(name)[:200],
        "source": str(source)[:300],
        "rows": int(len(frame)),
        "columns": [str(column)[:200] for column in frame.columns[:300]],
        "dtypes": {str(column)[:200]: str(frame[column].dtype) for column in frame.columns[:300]},
        "missing_cells": int(frame.isna().sum().sum()),
        "duplicate_rows": int(frame.duplicated().sum()) if len(frame.columns) else 0,
        "numeric_totals": numeric_totals,
        "fingerprint": dataframe_fingerprint(frame),
    }


@dataclass(frozen=True)
class LineageRun:
    run_id: str
    task_id: str
    job_name: str
    status: str
    started_at: str
    completed_at: str | None
    rule_version: str
    model: str
    prompt_version: str
    input_datasets: tuple[Mapping[str, Any], ...]
    output_datasets: tuple[Mapping[str, Any], ...]
    parameters: Mapping[str, Any]
    approvals: tuple[Mapping[str, Any], ...]
    error: str = ""

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["input_datasets"] = [dict(item) for item in self.input_datasets]
        payload["output_datasets"] = [dict(item) for item in self.output_datasets]
        payload["approvals"] = [dict(item) for item in self.approvals]
        return payload


class LineageStore:
    def __init__(self, database: str | Path) -> None:
        self.database = Path(database)
        self.database.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _connect(self):
        connection = sqlite3.connect(self.database, timeout=15)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA foreign_keys=ON")
        try:
            yield connection
            connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            connection.close()

    def _init_db(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS lineage_runs (
                    run_id TEXT PRIMARY KEY,
                    task_id TEXT NOT NULL,
                    job_name TEXT NOT NULL,
                    status TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    completed_at TEXT,
                    rule_version TEXT NOT NULL,
                    model TEXT NOT NULL,
                    prompt_version TEXT NOT NULL,
                    input_json TEXT NOT NULL,
                    output_json TEXT NOT NULL,
                    parameters_json TEXT NOT NULL,
                    approvals_json TEXT NOT NULL,
                    error TEXT NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_lineage_task_started
                    ON lineage_runs(task_id, started_at DESC);
                """
            )

    def start(
        self,
        task_id: str,
        job_name: str,
        *,
        inputs: Iterable[Mapping[str, Any]] = (),
        parameters: Mapping[str, Any] | None = None,
        rule_version: str = "local-v1",
        model: str = "",
        prompt_version: str = "",
    ) -> str:
        run_id = uuid.uuid4().hex
        with self._connect() as connection:
            connection.execute(
                "INSERT INTO lineage_runs VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    run_id, str(task_id)[:80], str(job_name)[:160], "running", _now(), None,
                    str(rule_version)[:80], str(model)[:80], str(prompt_version)[:80],
                    json.dumps(_safe_json(list(inputs)), ensure_ascii=False, allow_nan=False),
                    "[]", json.dumps(_safe_json(parameters or {}), ensure_ascii=False, allow_nan=False),
                    "[]", "",
                ),
            )
        return run_id

    def complete(
        self,
        run_id: str,
        *,
        outputs: Iterable[Mapping[str, Any]] = (),
        approvals: Iterable[Mapping[str, Any]] = (),
    ) -> None:
        with self._connect() as connection:
            changed = connection.execute(
                "UPDATE lineage_runs SET status='completed', completed_at=?, output_json=?, approvals_json=? WHERE run_id=? AND status='running'",
                (
                    _now(),
                    json.dumps(_safe_json(list(outputs)), ensure_ascii=False, allow_nan=False),
                    json.dumps(_safe_json(list(approvals)), ensure_ascii=False, allow_nan=False),
                    run_id,
                ),
            ).rowcount
        if not changed:
            raise KeyError("血缘运行不存在或已经结束")

    def fail(self, run_id: str, error: BaseException | str) -> None:
        message = str(error)[:1000]
        with self._connect() as connection:
            changed = connection.execute(
                "UPDATE lineage_runs SET status='failed', completed_at=?, error=? WHERE run_id=? AND status='running'",
                (_now(), message, run_id),
            ).rowcount
        if not changed:
            raise KeyError("血缘运行不存在或已经结束")

    def append_completed(
        self,
        task_id: str,
        job_name: str,
        *,
        inputs: Iterable[Mapping[str, Any]] = (),
        outputs: Iterable[Mapping[str, Any]] = (),
        parameters: Mapping[str, Any] | None = None,
        approvals: Iterable[Mapping[str, Any]] = (),
        rule_version: str = "local-v1",
        model: str = "",
        prompt_version: str = "",
    ) -> str:
        run_id = self.start(
            task_id, job_name, inputs=inputs, parameters=parameters,
            rule_version=rule_version, model=model, prompt_version=prompt_version,
        )
        self.complete(run_id, outputs=outputs, approvals=approvals)
        return run_id

    def _row(self, row: sqlite3.Row) -> LineageRun:
        def load(key: str, fallback: Any) -> Any:
            try:
                return json.loads(row[key])
            except (json.JSONDecodeError, TypeError):
                return fallback
        return LineageRun(
            run_id=row["run_id"], task_id=row["task_id"], job_name=row["job_name"],
            status=row["status"], started_at=row["started_at"], completed_at=row["completed_at"],
            rule_version=row["rule_version"], model=row["model"], prompt_version=row["prompt_version"],
            input_datasets=tuple(load("input_json", [])), output_datasets=tuple(load("output_json", [])),
            parameters=load("parameters_json", {}), approvals=tuple(load("approvals_json", [])),
            error=row["error"],
        )

    def list(self, task_id: str, *, limit: int = 200) -> list[LineageRun]:
        with self._connect() as connection:
            rows = connection.execute(
                "SELECT * FROM lineage_runs WHERE task_id=? ORDER BY started_at DESC LIMIT ?",
                (str(task_id), max(1, min(int(limit), 2000))),
            ).fetchall()
        return [self._row(row) for row in rows]

    def evidence(self, task_id: str) -> dict[str, Any]:
        runs = list(reversed(self.list(task_id, limit=2000)))
        return {
            "schema_version": LINEAGE_SCHEMA_VERSION,
            "task_id": str(task_id),
            "generated_at": _now(),
            "run_count": len(runs),
            "completed": sum(item.status == "completed" for item in runs),
            "failed": sum(item.status == "failed" for item in runs),
            "runs": [item.to_dict() for item in runs],
            "privacy_note": "证据包只包含结构、规则、汇总和哈希指纹，不包含客户单元格原值。",
        }

    def export_evidence(self, task_id: str, destination: str | Path) -> Path:
        path = Path(destination)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(self.evidence(task_id), ensure_ascii=False, indent=2, allow_nan=False), encoding="utf-8")
        return path


__all__ = [
    "LINEAGE_SCHEMA_VERSION", "LineageRun", "LineageStore",
    "dataframe_fingerprint", "dataset_metadata",
]
