"""Encrypted local database profiles and AST-backed read-only querying."""

from __future__ import annotations

import importlib.util
import json
import re
import sqlite3
import uuid
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Mapping

import pandas as pd

try:
    from .large_data import validate_read_only_sql
    from .secure_secrets import SecureSecretStore
except ImportError:  # pragma: no cover - direct script imports
    from large_data import validate_read_only_sql
    from secure_secrets import SecureSecretStore


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def _safe_sql(sql: str, *, dialect: str = "") -> str:
    query = validate_read_only_sql(sql)
    if importlib.util.find_spec("sqlglot") is None:
        return query
    import sqlglot  # type: ignore
    from sqlglot import expressions as exp  # type: ignore
    try:
        statements = sqlglot.parse(query, read=dialect or None)
    except Exception as exc:
        raise ValueError(f"SQL 语法无效：{str(exc)[:300]}") from exc
    if len(statements) != 1 or statements[0] is None:
        raise ValueError("只允许一条 SQL 查询")
    tree = statements[0]
    forbidden = tuple(
        kind for kind in (
            getattr(exp, "Insert", None), getattr(exp, "Update", None),
            getattr(exp, "Delete", None), getattr(exp, "Create", None),
            getattr(exp, "Drop", None), getattr(exp, "Alter", None),
            getattr(exp, "Command", None), getattr(exp, "Merge", None),
        )
        if kind is not None
    )
    if any(tree.find(kind) is not None for kind in forbidden):
        raise ValueError("SQL AST 检测到写入或管理语句")
    if tree.find(exp.Select) is None:
        raise ValueError("只允许 SELECT/WITH 查询")
    return query


@dataclass(frozen=True)
class ConnectionProfile:
    profile_id: str
    name: str
    kind: str
    dialect: str
    description: str
    created_at: str
    updated_at: str
    configured: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class ConnectionProfileStore:
    """Metadata in SQLite, credentials/paths in the Windows DPAPI vault."""

    def __init__(self, database: str | Path, secret_store: SecureSecretStore) -> None:
        self.database = Path(database)
        self.database.parent.mkdir(parents=True, exist_ok=True)
        self.secret_store = secret_store
        self._init_db()

    @contextmanager
    def _connect(self):
        connection = sqlite3.connect(self.database, timeout=15)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA journal_mode=WAL")
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
                CREATE TABLE IF NOT EXISTS connection_profiles (
                    profile_id TEXT PRIMARY KEY, name TEXT NOT NULL, kind TEXT NOT NULL,
                    dialect TEXT NOT NULL, description TEXT NOT NULL,
                    created_at TEXT NOT NULL, updated_at TEXT NOT NULL);
                CREATE TABLE IF NOT EXISTS query_bookmarks (
                    bookmark_id TEXT PRIMARY KEY, profile_id TEXT NOT NULL,
                    name TEXT NOT NULL, sql_text TEXT NOT NULL,
                    watermark_column TEXT NOT NULL, watermark_value TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    FOREIGN KEY(profile_id) REFERENCES connection_profiles(profile_id) ON DELETE CASCADE);
                """
            )

    def _secret_name(self, profile_id: str) -> str:
        return f"DB_PROFILE_{profile_id}"

    def save(
        self,
        *,
        name: str,
        kind: str,
        secret: Mapping[str, Any],
        dialect: str = "",
        description: str = "",
        profile_id: str | None = None,
    ) -> ConnectionProfile:
        if kind not in {"sqlite", "odbc"}:
            raise ValueError("连接类型只支持 sqlite 或 odbc")
        profile_id = profile_id or uuid.uuid4().hex[:16]
        if not re.fullmatch(r"[A-Za-z0-9_-]{8,64}", profile_id):
            raise ValueError("数据库连接编号无效")
        display = str(name).strip()[:120]
        if not display:
            raise ValueError("数据库连接名称不能为空")
        safe_secret: dict[str, str] = {}
        if kind == "sqlite":
            raw_path = str(secret.get("path") or "").strip()
            path = Path(raw_path).expanduser().resolve()
            if not path.is_file() or path.suffix.lower() not in {".db", ".sqlite", ".sqlite3"}:
                raise ValueError("SQLite 文件不存在或格式不支持")
            safe_secret = {"path": str(path)}
            dialect = "sqlite"
        else:
            connection_string = str(secret.get("connection_string") or "").strip()
            if not connection_string or len(connection_string) > 4096 or any(char in connection_string for char in "\r\n\x00"):
                raise ValueError("ODBC 连接字符串无效")
            safe_secret = {"connection_string": connection_string}
            dialect = str(dialect or "").strip().lower()[:40]
        encoded = json.dumps(safe_secret, ensure_ascii=False, separators=(",", ":"))
        self.secret_store.set(self._secret_name(profile_id), encoded)
        now = _now()
        with self._connect() as connection:
            existing = connection.execute("SELECT created_at FROM connection_profiles WHERE profile_id=?", (profile_id,)).fetchone()
            created = existing["created_at"] if existing else now
            connection.execute(
                """INSERT INTO connection_profiles VALUES (?,?,?,?,?,?,?)
                ON CONFLICT(profile_id) DO UPDATE SET name=excluded.name,kind=excluded.kind,
                dialect=excluded.dialect,description=excluded.description,updated_at=excluded.updated_at""",
                (profile_id, display, kind, dialect, str(description)[:500], created, now),
            )
        return self.get(profile_id)

    def _row(self, row: sqlite3.Row) -> ConnectionProfile:
        configured = bool(self.secret_store.get(self._secret_name(row["profile_id"])))
        return ConnectionProfile(
            profile_id=row["profile_id"], name=row["name"], kind=row["kind"], dialect=row["dialect"],
            description=row["description"], created_at=row["created_at"], updated_at=row["updated_at"], configured=configured,
        )

    def get(self, profile_id: str) -> ConnectionProfile:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM connection_profiles WHERE profile_id=?", (str(profile_id),)).fetchone()
        if not row:
            raise KeyError("数据库连接不存在")
        return self._row(row)

    def list(self) -> list[ConnectionProfile]:
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM connection_profiles ORDER BY name").fetchall()
        return [self._row(row) for row in rows]

    def delete(self, profile_id: str) -> bool:
        with self._connect() as connection:
            changed = connection.execute("DELETE FROM connection_profiles WHERE profile_id=?", (str(profile_id),)).rowcount
        self.secret_store.delete(self._secret_name(str(profile_id)))
        return bool(changed)

    def _secret(self, profile_id: str) -> tuple[ConnectionProfile, dict[str, str]]:
        profile = self.get(profile_id)
        raw = self.secret_store.get(self._secret_name(profile_id))
        if not raw:
            raise RuntimeError("数据库连接凭据缺失")
        payload = json.loads(raw)
        if not isinstance(payload, dict):
            raise RuntimeError("数据库连接凭据损坏")
        return profile, {str(key): str(value) for key, value in payload.items()}

    def test(self, profile_id: str) -> dict[str, Any]:
        started = datetime.now()
        profile, _ = self._secret(profile_id)
        if profile.kind == "sqlite":
            result = self.query(profile_id, "SELECT sqlite_version() AS version", max_rows=5)
        else:
            result = self.query(profile_id, "SELECT 1 AS connected", max_rows=5)
        return {"connected": True, "profile": profile.to_dict(), "duration_ms": int((datetime.now() - started).total_seconds() * 1000), "returned_rows": len(result)}

    def schema(self, profile_id: str, *, maximum_tables: int = 200) -> list[dict[str, Any]]:
        profile, secret = self._secret(profile_id)
        maximum_tables = min(max(1, int(maximum_tables)), 1000)
        results: list[dict[str, Any]] = []
        if profile.kind == "sqlite":
            uri = Path(secret["path"]).resolve().as_uri() + "?mode=ro"
            connection = sqlite3.connect(uri, uri=True, timeout=15)
            try:
                names = [row[0] for row in connection.execute("SELECT name FROM sqlite_master WHERE type IN ('table','view') AND name NOT LIKE 'sqlite_%' ORDER BY name LIMIT ?", (maximum_tables,)).fetchall()]
                for name in names:
                    escaped = name.replace('"', '""')
                    columns = connection.execute(f'PRAGMA table_info("{escaped}")').fetchall()
                    results.append({"schema": "main", "table": name, "columns": [{"name": row[1], "type": row[2], "nullable": not bool(row[3]), "primary_key": bool(row[5])} for row in columns]})
            finally:
                connection.close()
            return results
        try:
            import pyodbc  # type: ignore
        except ImportError as exc:
            raise RuntimeError("缺少 pyodbc，无法使用 ODBC 连接") from exc
        connection = pyodbc.connect(secret["connection_string"], timeout=15)
        try:
            cursor = connection.cursor()
            tables = []
            for row in cursor.tables(tableType="TABLE"):
                tables.append((str(row.table_schem or ""), str(row.table_name)))
                if len(tables) >= maximum_tables:
                    break
            for schema_name, table_name in tables:
                columns = []
                for row in cursor.columns(table=table_name, schema=schema_name or None):
                    columns.append({"name": str(row.column_name), "type": str(row.type_name), "nullable": bool(row.nullable), "primary_key": False})
                results.append({"schema": schema_name, "table": table_name, "columns": columns})
        finally:
            connection.close()
        return results

    def query(self, profile_id: str, sql: str, *, max_rows: int = 300_000, timeout: int = 30) -> pd.DataFrame:
        profile, secret = self._secret(profile_id)
        query = _safe_sql(sql, dialect=profile.dialect)
        max_rows = min(max(1, int(max_rows)), 1_000_000)
        if profile.kind == "sqlite":
            uri = Path(secret["path"]).resolve().as_uri() + "?mode=ro"
            connection = sqlite3.connect(uri, uri=True, timeout=min(max(1, timeout), 120))
            try:
                result = pd.read_sql_query(f"SELECT * FROM ({query}) AS result LIMIT {max_rows + 1}", connection)
            finally:
                connection.close()
        else:
            try:
                import pyodbc  # type: ignore
            except ImportError as exc:
                raise RuntimeError("缺少 pyodbc，无法使用 ODBC 连接") from exc
            connection = pyodbc.connect(secret["connection_string"], timeout=min(max(1, timeout), 120))
            try:
                result = pd.read_sql_query(query, connection)
            finally:
                connection.close()
        if len(result) > max_rows:
            raise ValueError(f"查询结果超过 {max_rows:,} 行安全上限")
        return result

    def save_bookmark(
        self,
        profile_id: str,
        *,
        name: str,
        sql: str,
        watermark_column: str = "",
        watermark_value: str = "",
        bookmark_id: str | None = None,
    ) -> str:
        profile = self.get(profile_id)
        del profile
        query = _safe_sql(sql)
        bookmark_id = bookmark_id or uuid.uuid4().hex[:16]
        with self._connect() as connection:
            connection.execute(
                """INSERT INTO query_bookmarks VALUES (?,?,?,?,?,?,?)
                ON CONFLICT(bookmark_id) DO UPDATE SET name=excluded.name,sql_text=excluded.sql_text,
                watermark_column=excluded.watermark_column,watermark_value=excluded.watermark_value,updated_at=excluded.updated_at""",
                (bookmark_id, profile_id, str(name)[:120], query, str(watermark_column)[:200], str(watermark_value)[:500], _now()),
            )
        return bookmark_id

    def list_bookmarks(self, profile_id: str) -> list[dict[str, Any]]:
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM query_bookmarks WHERE profile_id=? ORDER BY name", (profile_id,)).fetchall()
        return [dict(row) for row in rows]


__all__ = ["ConnectionProfile", "ConnectionProfileStore"]
