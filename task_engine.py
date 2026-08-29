"""Persistent bounded background jobs with progress, cancellation and retry."""

from __future__ import annotations

import json
import queue
import sqlite3
import threading
import time
import uuid
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Mapping


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="milliseconds")


@dataclass(frozen=True)
class JobRecord:
    job_id: str
    task_id: str
    job_type: str
    status: str
    progress: int
    message: str
    created_at: str
    started_at: str | None
    completed_at: str | None
    attempts: int
    max_attempts: int
    result: Mapping[str, Any]
    error: str
    cancel_requested: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


class JobCancelled(RuntimeError):
    pass


class JobContext:
    def __init__(self, engine: "PersistentJobEngine", job_id: str, task_id: str) -> None:
        self.engine = engine
        self.job_id = job_id
        self.task_id = task_id

    def update(self, progress: int, message: str = "") -> None:
        self.engine._update_progress(self.job_id, progress, message)
        self.raise_if_cancelled()

    def cancelled(self) -> bool:
        return self.engine._cancel_requested(self.job_id)

    def raise_if_cancelled(self) -> None:
        if self.cancelled():
            raise JobCancelled("任务已由用户取消")


JobHandler = Callable[[JobContext, Mapping[str, Any]], Mapping[str, Any] | None]


class PersistentJobEngine:
    """A small local worker pool; only explicitly registered job types run."""

    def __init__(self, database: str | Path, *, workers: int = 2, max_queue: int = 100) -> None:
        self.database = Path(database)
        self.database.parent.mkdir(parents=True, exist_ok=True)
        self.workers = max(1, min(int(workers), 8))
        self.max_queue = max(1, min(int(max_queue), 1000))
        self.handlers: dict[str, JobHandler] = {}
        self._queue: queue.Queue[str] = queue.Queue(maxsize=self.max_queue)
        self._stop = threading.Event()
        self._threads: list[threading.Thread] = []
        self._init_db()
        self._recover()

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
            connection.execute(
                """CREATE TABLE IF NOT EXISTS jobs (
                job_id TEXT PRIMARY KEY, task_id TEXT NOT NULL, job_type TEXT NOT NULL,
                status TEXT NOT NULL, progress INTEGER NOT NULL, message TEXT NOT NULL,
                payload_json TEXT NOT NULL, result_json TEXT NOT NULL, error TEXT NOT NULL,
                created_at TEXT NOT NULL, started_at TEXT, completed_at TEXT,
                attempts INTEGER NOT NULL, max_attempts INTEGER NOT NULL,
                cancel_requested INTEGER NOT NULL)"""
            )
            connection.execute("CREATE INDEX IF NOT EXISTS idx_jobs_task_created ON jobs(task_id, created_at DESC)")

    def _recover(self) -> None:
        with self._connect() as connection:
            connection.execute(
                "UPDATE jobs SET status='queued', progress=0, message='程序重启后等待恢复', started_at=NULL WHERE status IN ('running','retrying')"
            )

    def register(self, job_type: str, handler: JobHandler) -> None:
        name = str(job_type).strip()
        if not name or not name.replace("_", "").isalnum() or not callable(handler):
            raise ValueError("后台任务类型无效")
        self.handlers[name] = handler

    def start(self) -> None:
        if self._threads:
            return
        self._stop.clear()
        with self._connect() as connection:
            queued = [row[0] for row in connection.execute("SELECT job_id FROM jobs WHERE status='queued' AND cancel_requested=0 ORDER BY created_at").fetchall()]
        for job_id in queued[: self.max_queue]:
            try:
                self._queue.put_nowait(job_id)
            except queue.Full:
                break
        for index in range(self.workers):
            thread = threading.Thread(target=self._worker, name=f"biaoge-job-{index + 1}", daemon=True)
            thread.start()
            self._threads.append(thread)

    def stop(self, *, timeout: float = 3.0) -> None:
        self._stop.set()
        for _ in self._threads:
            try:
                self._queue.put_nowait("")
            except queue.Full:
                break
        deadline = time.monotonic() + max(0.0, timeout)
        for thread in self._threads:
            thread.join(max(0.0, deadline - time.monotonic()))
        self._threads.clear()

    def submit(
        self,
        task_id: str,
        job_type: str,
        payload: Mapping[str, Any],
        *,
        max_attempts: int = 2,
    ) -> JobRecord:
        if job_type not in self.handlers:
            raise ValueError("后台任务类型尚未注册")
        encoded = json.dumps(dict(payload), ensure_ascii=False, allow_nan=False)
        if len(encoded.encode("utf-8")) > 500_000:
            raise ValueError("后台任务参数过大")
        active = self.count_active(task_id)
        if active >= 10:
            raise RuntimeError("同一任务已有过多后台作业，请等待完成或取消")
        if self._queue.qsize() >= self.max_queue:
            raise RuntimeError("后台任务队列已满")
        job_id = uuid.uuid4().hex[:20]
        created = _now()
        with self._connect() as connection:
            connection.execute(
                "INSERT INTO jobs VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (job_id, str(task_id)[:80], job_type, "queued", 0, "等待执行", encoded, "{}", "", created, None, None, 0, max(1, min(int(max_attempts), 5)), 0),
            )
        self._queue.put_nowait(job_id)
        return self.get(job_id)

    def count_active(self, task_id: str) -> int:
        with self._connect() as connection:
            row = connection.execute("SELECT COUNT(*) FROM jobs WHERE task_id=? AND status IN ('queued','running','retrying')", (task_id,)).fetchone()
        return int(row[0])

    def _claim(self, job_id: str) -> tuple[JobRecord, dict[str, Any]] | None:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM jobs WHERE job_id=?", (job_id,)).fetchone()
            if not row or row["status"] != "queued" or row["cancel_requested"]:
                return None
            changed = connection.execute(
                "UPDATE jobs SET status='running', started_at=?, progress=1, message='开始执行', attempts=attempts+1 WHERE job_id=? AND status='queued'",
                (_now(), job_id),
            ).rowcount
            if not changed:
                return None
            payload = json.loads(row["payload_json"])
        return self.get(job_id), payload

    def _worker(self) -> None:
        while not self._stop.is_set():
            try:
                job_id = self._queue.get(timeout=0.5)
            except queue.Empty:
                continue
            if not job_id:
                self._queue.task_done()
                continue
            try:
                claimed = self._claim(job_id)
                if not claimed:
                    continue
                record, payload = claimed
                handler = self.handlers.get(record.job_type)
                if handler is None:
                    raise RuntimeError("任务执行器未注册")
                context = JobContext(self, record.job_id, record.task_id)
                context.update(5, "正在准备数据")
                result = handler(context, payload) or {}
                context.raise_if_cancelled()
                encoded = json.dumps(dict(result), ensure_ascii=False, allow_nan=False)
                if len(encoded.encode("utf-8")) > 1_000_000:
                    result = {"message": "任务完成；详细结果已写入交付文件", "result_truncated": True}
                    encoded = json.dumps(result, ensure_ascii=False)
                with self._connect() as connection:
                    connection.execute(
                        "UPDATE jobs SET status='completed',progress=100,message='执行完成',result_json=?,completed_at=? WHERE job_id=?",
                        (encoded, _now(), job_id),
                    )
            except JobCancelled as exc:
                with self._connect() as connection:
                    connection.execute("UPDATE jobs SET status='cancelled',message=?,error=?,completed_at=? WHERE job_id=?", (str(exc), str(exc), _now(), job_id))
            except Exception as exc:  # worker boundary persists the failure
                with self._connect() as connection:
                    row = connection.execute("SELECT attempts,max_attempts,cancel_requested FROM jobs WHERE job_id=?", (job_id,)).fetchone()
                    if row and not row["cancel_requested"] and row["attempts"] < row["max_attempts"]:
                        connection.execute("UPDATE jobs SET status='queued',progress=0,message='失败后等待重试',error=? WHERE job_id=?", (f"{type(exc).__name__}: {str(exc)[:800]}", job_id))
                        try:
                            self._queue.put_nowait(job_id)
                        except queue.Full:
                            connection.execute("UPDATE jobs SET status='failed',message='重试队列已满',completed_at=? WHERE job_id=?", (_now(), job_id))
                    else:
                        connection.execute("UPDATE jobs SET status='failed',message='执行失败',error=?,completed_at=? WHERE job_id=?", (f"{type(exc).__name__}: {str(exc)[:800]}", _now(), job_id))
            finally:
                self._queue.task_done()

    def _update_progress(self, job_id: str, progress: int, message: str) -> None:
        with self._connect() as connection:
            connection.execute(
                "UPDATE jobs SET progress=?,message=? WHERE job_id=? AND status='running'",
                (max(1, min(int(progress), 99)), str(message)[:300], job_id),
            )

    def _cancel_requested(self, job_id: str) -> bool:
        with self._connect() as connection:
            row = connection.execute("SELECT cancel_requested FROM jobs WHERE job_id=?", (job_id,)).fetchone()
        return bool(row and row[0])

    def cancel(self, job_id: str) -> JobRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT status FROM jobs WHERE job_id=?", (job_id,)).fetchone()
            if not row:
                raise KeyError("后台任务不存在")
            if row["status"] in {"completed", "failed", "cancelled"}:
                return self.get(job_id)
            connection.execute("UPDATE jobs SET cancel_requested=1,message='正在取消' WHERE job_id=?", (job_id,))
            if row["status"] == "queued":
                connection.execute("UPDATE jobs SET status='cancelled',completed_at=?,message='已取消' WHERE job_id=?", (_now(), job_id))
        return self.get(job_id)

    def retry(self, job_id: str) -> JobRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT status FROM jobs WHERE job_id=?", (job_id,)).fetchone()
            if not row:
                raise KeyError("后台任务不存在")
            if row["status"] not in {"failed", "cancelled"}:
                raise ValueError("只有失败或已取消的任务可以重试")
            connection.execute("UPDATE jobs SET status='queued',progress=0,message='等待重试',error='',completed_at=NULL,cancel_requested=0 WHERE job_id=?", (job_id,))
        self._queue.put_nowait(job_id)
        return self.get(job_id)

    def _record(self, row: sqlite3.Row) -> JobRecord:
        try:
            result = json.loads(row["result_json"] or "{}")
        except json.JSONDecodeError:
            result = {}
        return JobRecord(
            job_id=row["job_id"], task_id=row["task_id"], job_type=row["job_type"], status=row["status"],
            progress=int(row["progress"]), message=row["message"], created_at=row["created_at"],
            started_at=row["started_at"], completed_at=row["completed_at"], attempts=int(row["attempts"]),
            max_attempts=int(row["max_attempts"]), result=result, error=row["error"], cancel_requested=bool(row["cancel_requested"]),
        )

    def get(self, job_id: str) -> JobRecord:
        with self._connect() as connection:
            row = connection.execute("SELECT * FROM jobs WHERE job_id=?", (job_id,)).fetchone()
        if not row:
            raise KeyError("后台任务不存在")
        return self._record(row)

    def list(self, task_id: str | None = None, *, limit: int = 100) -> list[JobRecord]:
        with self._connect() as connection:
            if task_id:
                rows = connection.execute("SELECT * FROM jobs WHERE task_id=? ORDER BY created_at DESC LIMIT ?", (task_id, min(max(int(limit), 1), 1000))).fetchall()
            else:
                rows = connection.execute("SELECT * FROM jobs ORDER BY created_at DESC LIMIT ?", (min(max(int(limit), 1), 1000),)).fetchall()
        return [self._record(row) for row in rows]

    def counts(self, task_id: str | None = None) -> dict[str, int]:
        """Return compact persistent queue statistics."""

        with self._connect() as connection:
            if task_id:
                rows = connection.execute(
                    "SELECT status, COUNT(*) AS amount FROM jobs WHERE task_id=? GROUP BY status",
                    (str(task_id),),
                ).fetchall()
            else:
                rows = connection.execute(
                    "SELECT status, COUNT(*) AS amount FROM jobs GROUP BY status"
                ).fetchall()
        result = {name: 0 for name in ("queued", "running", "completed", "failed", "cancelled")}
        for row in rows:
            result[str(row["status"])] = int(row["amount"])
        result["total"] = sum(result.values())
        return result


__all__ = ["JobCancelled", "JobContext", "JobRecord", "PersistentJobEngine"]
