"""Deterministic selection of pandas, DuckDB or Polars execution engines."""

from __future__ import annotations

import importlib.util
from dataclasses import asdict, dataclass
from typing import Any, Mapping, Sequence

import pandas as pd


@dataclass(frozen=True)
class EngineDecision:
    engine: str
    reason: str
    estimated_rows: int
    estimated_bytes: int
    operation: str
    fallback_chain: tuple[str, ...]

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["fallback_chain"] = list(self.fallback_chain)
        return payload


def available_engines() -> dict[str, bool]:
    return {
        "pandas": True,
        "duckdb": importlib.util.find_spec("duckdb") is not None,
        "polars": importlib.util.find_spec("polars") is not None,
    }


def estimate_frames(frames: Sequence[pd.DataFrame]) -> tuple[int, int]:
    rows = sum(len(frame) for frame in frames)
    memory = 0
    for frame in frames:
        try:
            memory += int(frame.memory_usage(index=True, deep=True).sum())
        except (TypeError, ValueError):
            memory += int(frame.memory_usage(index=True).sum())
    return rows, memory


def choose_engine(
    frames: Sequence[pd.DataFrame],
    *,
    operation: str,
    prefer: str = "auto",
    result_rows_limit: int = 300_000,
) -> EngineDecision:
    if not frames:
        raise ValueError("引擎选择至少需要一张数据表")
    if prefer not in {"auto", "pandas", "duckdb", "polars"}:
        raise ValueError("计算引擎偏好无效")
    availability = available_engines()
    rows, memory = estimate_frames(frames)
    operation = str(operation or "transform")
    if prefer != "auto":
        if availability.get(prefer):
            return EngineDecision(prefer, "用户或已确认方案指定", rows, memory, operation, (prefer, "pandas") if prefer != "pandas" else ("pandas",))
        return EngineDecision("pandas", f"指定的 {prefer} 未安装，安全回退 pandas", rows, memory, operation, ("pandas",))
    relational = operation in {"concat", "join", "summary", "pivot", "trend", "query", "compare"}
    transform = operation in {"clean", "select", "replace", "mask", "filter"}
    if availability["duckdb"] and relational and (rows >= 150_000 or memory >= 160 * 1024 * 1024):
        return EngineDecision("duckdb", "大规模关联/聚合使用向量化 SQL 与查询下推", rows, memory, operation, ("duckdb", "polars", "pandas"))
    if availability["polars"] and transform and (rows >= 250_000 or memory >= 256 * 1024 * 1024):
        return EngineDecision("polars", "大规模转换使用 Lazy/Streaming 执行", rows, memory, operation, ("polars", "duckdb", "pandas"))
    if availability["duckdb"] and relational and rows >= max(50_000, result_rows_limit // 2):
        return EngineDecision("duckdb", "中大型聚合优先减少 Python 内存复制", rows, memory, operation, ("duckdb", "pandas"))
    return EngineDecision("pandas", "数据规模适合内存处理，优先保持 Excel 类型与兼容性", rows, memory, operation, ("pandas",))


def group_summary_auto(
    frame: pd.DataFrame,
    *,
    by: Sequence[str],
    aggregations: Mapping[str, str],
    prefer: str = "auto",
) -> tuple[pd.DataFrame, EngineDecision]:
    if not by or not aggregations:
        raise ValueError("分组汇总需要维度和指标")
    missing = [name for name in [*by, *aggregations] if name not in frame.columns]
    if missing:
        raise KeyError(f"分组汇总字段不存在：{', '.join(missing)}")
    allowed = {"sum", "mean", "min", "max", "count", "nunique"}
    if any(operation not in allowed for operation in aggregations.values()):
        raise ValueError("聚合函数不在白名单")
    decision = choose_engine([frame], operation="summary", prefer=prefer)
    if decision.engine == "duckdb":
        try:
            import duckdb  # type: ignore
            def quote(name: str) -> str:
                return '"' + name.replace('"', '""') + '"'
            expressions = []
            function_map = {"mean": "AVG", "nunique": "COUNT(DISTINCT {column})"}
            for column, operation in aggregations.items():
                alias = str(column)
                if operation == "nunique":
                    expression = function_map[operation].format(column=quote(column))
                else:
                    function = function_map.get(operation, operation.upper())
                    expression = f"{function}({quote(column)})"
                expressions.append(f"{expression} AS {quote(alias)}")
            group = ", ".join(quote(name) for name in by)
            sql = f"SELECT {group}, {', '.join(expressions)} FROM input_frame GROUP BY {group}"
            connection = duckdb.connect(":memory:")
            try:
                connection.register("input_frame", frame)
                result = connection.execute(sql).fetchdf()
            finally:
                connection.close()
            return result, decision
        except Exception:
            decision = EngineDecision("pandas", "DuckDB 执行失败，已回退 pandas", decision.estimated_rows, decision.estimated_bytes, decision.operation, ("pandas",))
    result = frame.groupby(list(by), dropna=False, observed=True).agg(dict(aggregations)).reset_index()
    flattened = []
    for column in result.columns:
        if isinstance(column, tuple):
            flattened.append("_".join(str(item) for item in column if str(item)))
        else:
            flattened.append(str(column))
    result.columns = flattened
    return result, decision


def concat_auto(
    frames: Sequence[pd.DataFrame],
    *,
    join: str = "outer",
    prefer: str = "auto",
) -> tuple[pd.DataFrame, EngineDecision]:
    if len(frames) < 2:
        raise ValueError("合并至少需要两张表")
    if join not in {"outer", "inner"}:
        raise ValueError("字段对齐方式无效")
    decision = choose_engine(frames, operation="concat", prefer=prefer)
    if decision.engine == "polars":
        try:
            import polars as pl  # type: ignore
            how = "diagonal" if join == "outer" else "vertical_relaxed"
            result = pl.concat([pl.from_pandas(frame) for frame in frames], how=how).to_pandas()
            return result, decision
        except Exception:
            decision = EngineDecision("pandas", "Polars 执行失败，已回退 pandas", decision.estimated_rows, decision.estimated_bytes, decision.operation, ("pandas",))
    return pd.concat(list(frames), ignore_index=True, join=join, sort=False), decision


__all__ = ["EngineDecision", "available_engines", "choose_engine", "concat_auto", "estimate_frames", "group_summary_auto"]
