"""Optional DuckDB acceleration for large local CSV/Parquet orders."""

from __future__ import annotations

from pathlib import Path
import re
from typing import Any, Iterable

import pandas as pd


READ_ONLY_SQL = re.compile(r"^\s*(?:with\b[\s\S]+?\bselect\b|select\b)", re.IGNORECASE)
FORBIDDEN_SQL = re.compile(
    r"\b(insert|update|delete|drop|alter|create|truncate|merge|grant|revoke|attach|detach|copy|call|execute|exec|install|load|pragma|vacuum|read_csv|read_csv_auto|read_parquet|glob|sqlite_scan|postgres_scan|mysql_scan|httpfs)\b",
    re.IGNORECASE,
)


class LargeDataUnavailable(RuntimeError):
    pass


def duckdb_available() -> bool:
    try:
        import duckdb  # noqa: F401
    except ImportError:
        return False
    return True


def validate_read_only_sql(sql: str) -> str:
    query = str(sql).strip()
    if not query or len(query) > 100_000:
        raise ValueError("SQL 为空或过长")
    without_comments = re.sub(r"--[^\r\n]*|/\*[\s\S]*?\*/", " ", query)
    if not READ_ONLY_SQL.match(without_comments) or FORBIDDEN_SQL.search(without_comments):
        raise ValueError("只允许单条 SELECT/WITH 只读查询")
    statements = [item for item in without_comments.split(";") if item.strip()]
    if len(statements) != 1:
        raise ValueError("只允许执行一条只读查询")
    return query.rstrip(";")


def query_files(
    paths: Iterable[str | Path],
    sql: str,
    *,
    max_rows: int = 1_000_000,
) -> pd.DataFrame:
    """Query local CSV/Parquet files as ``input_1``, ``input_2`` ... views."""

    if not duckdb_available():
        raise LargeDataUnavailable("未安装 DuckDB；请运行健康检查并安装大数据可选组件")
    import duckdb

    safe_query = validate_read_only_sql(sql)
    resolved = [Path(path).resolve() for path in paths]
    if not resolved or len(resolved) > 50:
        raise ValueError("需要 1~50 个本地数据文件")
    connection = duckdb.connect(database=":memory:", read_only=False)
    try:
        for index, path in enumerate(resolved, start=1):
            if not path.is_file() or path.suffix.lower() not in {".csv", ".parquet"}:
                raise ValueError(f"大数据引擎不支持文件：{path.name}")
            view = f"input_{index}"
            # DuckDB does not allow prepared parameters in a CREATE VIEW
            # statement.  Paths are resolved by the application and quoted as
            # SQL string literals; doubling apostrophes prevents injection.
            path_literal = str(path).replace("'", "''")
            if path.suffix.lower() == ".csv":
                connection.execute(
                    f"CREATE VIEW {view} AS SELECT * FROM read_csv_auto('{path_literal}', header=true, all_varchar=false)",
                )
            else:
                connection.execute(
                    f"CREATE VIEW {view} AS SELECT * FROM read_parquet('{path_literal}')"
                )
        limited_query = f"SELECT * FROM ({safe_query}) AS result LIMIT {int(max_rows) + 1}"
        result = connection.execute(limited_query).fetch_df()
        if len(result) > max_rows:
            raise ValueError(f"查询结果超过 {max_rows:,} 行安全上限；请增加汇总或过滤条件")
        return result
    finally:
        connection.close()


def dataframe_query(tables: dict[str, pd.DataFrame], sql: str, *, max_rows: int = 1_000_000) -> pd.DataFrame:
    if not duckdb_available():
        raise LargeDataUnavailable("未安装 DuckDB；请安装大数据可选组件")
    import duckdb

    safe_query = validate_read_only_sql(sql)
    if not tables or len(tables) > 50:
        raise ValueError("需要 1~50 张数据表")
    connection = duckdb.connect(database=":memory:")
    try:
        for name, frame in tables.items():
            if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,62}", name):
                raise ValueError(f"DuckDB 表别名无效：{name}")
            connection.register(name, frame)
        result = connection.execute(
            f"SELECT * FROM ({safe_query}) AS result LIMIT {int(max_rows) + 1}"
        ).fetch_df()
        if len(result) > max_rows:
            raise ValueError(f"查询结果超过 {max_rows:,} 行安全上限")
        return result
    finally:
        connection.close()


def profile_file(path: str | Path) -> dict[str, Any]:
    source = Path(path).resolve()
    return {
        "filename": source.name,
        "size_bytes": source.stat().st_size,
        "recommended_engine": "duckdb" if source.stat().st_size >= 20 * 1024 * 1024 else "pandas",
        "duckdb_available": duckdb_available(),
    }


__all__ = [
    "LargeDataUnavailable", "dataframe_query", "duckdb_available",
    "profile_file", "query_files", "validate_read_only_sql",
]
