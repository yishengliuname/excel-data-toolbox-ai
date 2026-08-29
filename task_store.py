"""Crash-safe, one-task-one-folder persistence and retention controls."""

from __future__ import annotations

from datetime import datetime, timedelta
import gzip
import json
import os
from pathlib import Path
import shutil
import sqlite3
import tempfile
from typing import Any, Mapping

import pandas as pd


TASK_ID_CHARS = frozenset("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-_")


def _safe_task_id(value: str) -> str:
    task_id = str(value).strip().upper()
    if not 6 <= len(task_id) <= 64 or any(character not in TASK_ID_CHARS for character in task_id):
        raise ValueError("任务编号无效")
    return task_id


class TaskRepository:
    """Persist task manifests and data frames without executable pickle data."""

    def __init__(self, root: str | Path, *, retention_days: int = 30) -> None:
        if not 1 <= int(retention_days) <= 3650:
            raise ValueError("retention_days 必须在 1~3650 之间")
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)
        self.retention_days = int(retention_days)
        self.db_path = self.root / "tasks.sqlite3"
        self._initialise()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path, timeout=10)
        connection.execute("PRAGMA journal_mode=WAL")
        connection.execute("PRAGMA foreign_keys=ON")
        return connection

    def _initialise(self) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS tasks (
                    task_id TEXT PRIMARY KEY,
                    task_name TEXT NOT NULL,
                    status TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    table_count INTEGER NOT NULL DEFAULT 0,
                    operation_count INTEGER NOT NULL DEFAULT 0,
                    expires_at TEXT NOT NULL
                )
                """
            )

    def task_dir(self, task_id: str) -> Path:
        safe = _safe_task_id(task_id)
        candidate = (self.root / safe).resolve()
        if self.root not in candidate.parents:
            raise ValueError("任务目录越界")
        return candidate

    def create(self, task_id: str, task_name: str) -> Path:
        safe = _safe_task_id(task_id)
        now = datetime.now()
        expires = now + timedelta(days=self.retention_days)
        task_dir = self.task_dir(safe)
        (task_dir / "source_files").mkdir(parents=True, exist_ok=True)
        (task_dir / "deliverables").mkdir(parents=True, exist_ok=True)
        (task_dir / "tables").mkdir(parents=True, exist_ok=True)
        with self._connect() as connection:
            connection.execute(
                """INSERT OR REPLACE INTO tasks
                (task_id,task_name,status,created_at,updated_at,table_count,operation_count,expires_at)
                VALUES (?,?,?,?,?,?,?,?)""",
                (safe, str(task_name)[:120], "active", now.isoformat(), now.isoformat(), 0, 0, expires.isoformat()),
            )
        return task_dir

    def save(
        self,
        task_id: str,
        *,
        task_name: str,
        tables: Mapping[str, tuple[str, pd.DataFrame, str, bool]],
        active_table: str | None,
        operations: list[dict[str, Any]],
        file_names: list[str],
        import_warnings: list[str],
    ) -> Path:
        task_dir = self.task_dir(task_id)
        task_dir.mkdir(parents=True, exist_ok=True)
        table_dir = task_dir / "tables"
        table_dir.mkdir(exist_ok=True)
        table_manifest: list[dict[str, Any]] = []
        live_files: set[str] = set()
        for table_id, (name, frame, source, original) in tables.items():
            filename = f"{table_id}.json.gz"
            live_files.add(filename)
            target = table_dir / filename
            handle, temp_name = tempfile.mkstemp(prefix=f".{table_id}_", suffix=".json.gz", dir=table_dir)
            os.close(handle)
            temp = Path(temp_name)
            try:
                payload = frame.reset_index(drop=True).to_json(
                    orient="table", date_format="iso", date_unit="ms", force_ascii=False
                ).encode("utf-8")
                with gzip.open(temp, "wb", compresslevel=6) as stream:
                    stream.write(payload)
                os.replace(temp, target)
            finally:
                temp.unlink(missing_ok=True)
            table_manifest.append(
                {
                    "id": table_id,
                    "name": name,
                    "source": source,
                    "original": bool(original),
                    "file": filename,
                    "rows": len(frame),
                    "columns": len(frame.columns),
                    # pandas' table JSON reader normalises datetime precision
                    # (for example us -> ns).  Keep the declared dtypes so a
                    # recovered order behaves exactly like the live task.
                    "dtypes": [str(dtype) for dtype in frame.dtypes],
                }
            )
        for stale in table_dir.glob("*.json.gz"):
            if stale.name not in live_files:
                stale.unlink(missing_ok=True)
        manifest = {
            "schema_version": 1,
            "task_id": _safe_task_id(task_id),
            "task_name": str(task_name)[:120],
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "active_table": active_table,
            "file_names": list(file_names)[:1000],
            "import_warnings": list(import_warnings)[:2000],
            "operations": operations[-2000:],
            "tables": table_manifest,
        }
        manifest_path = task_dir / "manifest.json"
        temp_manifest = task_dir / ".manifest.tmp"
        temp_manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        os.replace(temp_manifest, manifest_path)
        now = datetime.now()
        expires = now + timedelta(days=self.retention_days)
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO tasks
                (task_id,task_name,status,created_at,updated_at,table_count,operation_count,expires_at)
                VALUES (?,?,?,?,?,?,?,?)
                ON CONFLICT(task_id) DO UPDATE SET task_name=excluded.task_name,
                status=excluded.status, updated_at=excluded.updated_at,
                table_count=excluded.table_count, operation_count=excluded.operation_count,
                expires_at=excluded.expires_at""",
                (
                    _safe_task_id(task_id), str(task_name)[:120], "active", now.isoformat(), now.isoformat(),
                    len(table_manifest), len(operations), expires.isoformat(),
                ),
            )
        return manifest_path

    def load(self, task_id: str) -> dict[str, Any]:
        task_dir = self.task_dir(task_id)
        manifest_path = task_dir / "manifest.json"
        if not manifest_path.is_file():
            raise FileNotFoundError("找不到可恢复的任务")
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        if not isinstance(manifest, dict) or manifest.get("schema_version") != 1:
            raise ValueError("任务清单版本无效")
        tables: dict[str, tuple[str, pd.DataFrame, str, bool]] = {}
        for item in manifest.get("tables", []):
            if not isinstance(item, dict):
                raise ValueError("任务数据表清单无效")
            table_id = str(item.get("id") or "")
            filename = str(item.get("file") or "")
            if filename != f"{table_id}.json.gz":
                raise ValueError("任务数据文件名无效")
            path = (task_dir / "tables" / filename).resolve()
            if task_dir not in path.parents or not path.is_file():
                raise ValueError("任务数据文件缺失或越界")
            with gzip.open(path, "rt", encoding="utf-8") as stream:
                frame = pd.read_json(stream, orient="table")
            declared_dtypes = item.get("dtypes", [])
            if declared_dtypes:
                if not isinstance(declared_dtypes, list) or len(declared_dtypes) != len(frame.columns):
                    raise ValueError("任务数据类型清单无效")
                for column_index, dtype in enumerate(declared_dtypes):
                    if not isinstance(dtype, str) or len(dtype) > 100:
                        raise ValueError("任务字段类型无效")
                    try:
                        frame.isetitem(column_index, frame.iloc[:, column_index].astype(dtype))
                    except (TypeError, ValueError):
                        # Old/third-party pandas extension dtypes may not be
                        # installed on recovery; data remains available and the
                        # manifest still records the intended type.
                        pass
            tables[table_id] = (
                str(item.get("name") or "数据表"), frame,
                str(item.get("source") or "恢复任务"), bool(item.get("original")),
            )
        manifest["loaded_tables"] = tables
        return manifest

    def list_tasks(self, *, limit: int = 50) -> list[dict[str, Any]]:
        safe_limit = max(1, min(int(limit), 500))
        with self._connect() as connection:
            connection.row_factory = sqlite3.Row
            rows = connection.execute(
                "SELECT * FROM tasks ORDER BY updated_at DESC LIMIT ?", (safe_limit,)
            ).fetchall()
        return [dict(row) for row in rows]

    def delete(self, task_id: str) -> None:
        directory = self.task_dir(task_id)
        if directory.exists():
            shutil.rmtree(directory)
        with self._connect() as connection:
            connection.execute("DELETE FROM tasks WHERE task_id=?", (_safe_task_id(task_id),))

    def purge_expired(self, *, now: datetime | None = None) -> list[str]:
        current = (now or datetime.now()).isoformat()
        with self._connect() as connection:
            rows = connection.execute("SELECT task_id FROM tasks WHERE expires_at < ?", (current,)).fetchall()
        removed: list[str] = []
        for (task_id,) in rows:
            self.delete(task_id)
            removed.append(task_id)
        return removed


__all__ = ["TaskRepository"]
