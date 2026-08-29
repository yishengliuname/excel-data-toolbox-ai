"""Persistent local schedules for approved, deterministic batch jobs."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
import json
from pathlib import Path
import sqlite3
import threading
import time
import uuid
from typing import Any, Callable, Mapping


@dataclass(frozen=True)
class Schedule:
    id: str
    name: str
    kind: str
    expression: str
    job_type: str
    payload: Mapping[str, Any]
    enabled: bool
    next_run: str
    last_run: str | None = None
    last_status: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def _next_run(kind: str, expression: str, now: datetime | None = None) -> datetime:
    current = now or datetime.now()
    if kind == "interval_minutes":
        minutes = int(expression)
        if not 1 <= minutes <= 10080:
            raise ValueError("间隔分钟必须在 1~10080 之间")
        return current + timedelta(minutes=minutes)
    if kind == "daily":
        try:
            hour, minute = (int(item) for item in expression.split(":"))
        except (ValueError, TypeError):
            raise ValueError("每日时间格式必须为 HH:MM") from None
        if not 0 <= hour <= 23 or not 0 <= minute <= 59:
            raise ValueError("每日时间无效")
        candidate = current.replace(hour=hour, minute=minute, second=0, microsecond=0)
        return candidate if candidate > current else candidate + timedelta(days=1)
    if kind == "weekly":
        try:
            weekday_text, clock = expression.split("@", 1)
            weekday = int(weekday_text)
            hour, minute = (int(item) for item in clock.split(":"))
        except (ValueError, TypeError):
            raise ValueError("每周格式必须为 0@HH:MM（0代表周一）") from None
        if not 0 <= weekday <= 6 or not 0 <= hour <= 23 or not 0 <= minute <= 59:
            raise ValueError("每周时间无效")
        days = (weekday - current.weekday()) % 7
        candidate = (current + timedelta(days=days)).replace(hour=hour, minute=minute, second=0, microsecond=0)
        return candidate if candidate > current else candidate + timedelta(days=7)
    raise ValueError("计划类型必须是 interval_minutes、daily 或 weekly")


class LocalScheduler:
    """Small scheduler that only invokes explicitly registered local job types."""

    def __init__(self, database: str | Path) -> None:
        self.database = Path(database)
        self.database.parent.mkdir(parents=True, exist_ok=True)
        self.callbacks: dict[str, Callable[[Mapping[str, Any]], None]] = {}
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.database, timeout=10)
        connection.row_factory = sqlite3.Row
        return connection

    def _init_db(self) -> None:
        with self._connect() as connection:
            connection.execute(
                """CREATE TABLE IF NOT EXISTS schedules (
                id TEXT PRIMARY KEY, name TEXT NOT NULL, kind TEXT NOT NULL,
                expression TEXT NOT NULL, job_type TEXT NOT NULL, payload TEXT NOT NULL,
                enabled INTEGER NOT NULL, next_run TEXT NOT NULL,
                last_run TEXT, last_status TEXT)"""
            )

    def register_job(self, job_type: str, callback: Callable[[Mapping[str, Any]], None]) -> None:
        if not job_type or not callable(callback):
            raise ValueError("任务类型和回调无效")
        self.callbacks[job_type] = callback

    def add(self, name: str, kind: str, expression: str, job_type: str, payload: Mapping[str, Any]) -> Schedule:
        if job_type not in self.callbacks:
            raise ValueError("计划任务类型尚未注册，不能创建不可执行的计划")
        encoded = json.dumps(dict(payload), ensure_ascii=False, allow_nan=False)
        if len(encoded.encode("utf-8")) > 100_000:
            raise ValueError("计划任务参数过大")
        identifier = uuid.uuid4().hex[:12]
        next_run = _next_run(kind, expression).isoformat(timespec="seconds")
        with self._connect() as connection:
            connection.execute(
                "INSERT INTO schedules VALUES (?,?,?,?,?,?,?,?,?,?)",
                (identifier, str(name)[:120], kind, expression, job_type, encoded, 1, next_run, None, None),
            )
        return Schedule(identifier, str(name)[:120], kind, expression, job_type, dict(payload), True, next_run)

    def list(self) -> list[Schedule]:
        with self._connect() as connection:
            rows = connection.execute("SELECT * FROM schedules ORDER BY next_run").fetchall()
        return [
            Schedule(
                id=row["id"], name=row["name"], kind=row["kind"], expression=row["expression"],
                job_type=row["job_type"], payload=json.loads(row["payload"]), enabled=bool(row["enabled"]),
                next_run=row["next_run"], last_run=row["last_run"], last_status=row["last_status"],
            )
            for row in rows
        ]

    def set_enabled(self, schedule_id: str, enabled: bool) -> None:
        with self._connect() as connection:
            changed = connection.execute(
                "UPDATE schedules SET enabled=? WHERE id=?", (int(bool(enabled)), schedule_id)
            ).rowcount
        if not changed:
            raise KeyError("计划任务不存在")

    def delete(self, schedule_id: str) -> None:
        with self._connect() as connection:
            changed = connection.execute("DELETE FROM schedules WHERE id=?", (schedule_id,)).rowcount
        if not changed:
            raise KeyError("计划任务不存在")

    def run_due(self, now: datetime | None = None) -> int:
        current = now or datetime.now()
        ran = 0
        for schedule in self.list():
            if not schedule.enabled or datetime.fromisoformat(schedule.next_run) > current:
                continue
            callback = self.callbacks.get(schedule.job_type)
            status = "failed: job type not registered"
            if callback is not None:
                try:
                    callback(schedule.payload)
                    status = "success"
                except Exception as exc:  # boundary: status must survive callback failure
                    status = f"failed: {type(exc).__name__}: {str(exc)[:300]}"
            next_run = _next_run(schedule.kind, schedule.expression, current).isoformat(timespec="seconds")
            with self._connect() as connection:
                connection.execute(
                    "UPDATE schedules SET last_run=?,last_status=?,next_run=? WHERE id=?",
                    (current.isoformat(timespec="seconds"), status, next_run, schedule.id),
                )
            ran += 1
        return ran

    def start(self, *, poll_seconds: int = 30) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        def loop() -> None:
            while not self._stop.wait(max(5, int(poll_seconds))):
                self.run_due()
        self._thread = threading.Thread(target=loop, name="biaoge-scheduler", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()
        if self._thread:
            self._thread.join(timeout=2)


__all__ = ["LocalScheduler", "Schedule"]
