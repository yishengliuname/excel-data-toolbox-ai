"""Request-scoped session registry for concurrent local tasks."""

from __future__ import annotations

import threading
from collections import OrderedDict
from contextvars import ContextVar
from typing import Any, Callable


class SessionRegistry:
    def __init__(
        self,
        factory: Callable[[str | None], Any],
        initial: Any,
        *,
        maximum_loaded: int = 12,
    ) -> None:
        self.factory = factory
        self.maximum_loaded = max(2, min(int(maximum_loaded), 100))
        self.lock = threading.RLock()
        self.sessions: OrderedDict[str, Any] = OrderedDict()
        self.sessions[str(initial.task_id)] = initial
        self.default_task_id = str(initial.task_id)
        self._current: ContextVar[str | None] = ContextVar("biaoge_task_session", default=None)

    def _touch(self, task_id: str) -> Any:
        session = self.sessions.pop(task_id)
        self.sessions[task_id] = session
        return session

    def _trim(self) -> None:
        while len(self.sessions) > self.maximum_loaded:
            task_id, session = next(iter(self.sessions.items()))
            if task_id == self.default_task_id or task_id == self._current.get():
                self.sessions.move_to_end(task_id)
                continue
            self.sessions.pop(task_id, None)
            try:
                session.close()
            except Exception:
                pass

    def bind(self, task_id: str | None, *, load: bool = True) -> Any:
        identifier = str(task_id or "").strip()
        with self.lock:
            if not identifier:
                identifier = self.default_task_id
            if identifier not in self.sessions:
                if not load:
                    identifier = self.default_task_id
                else:
                    self.sessions[identifier] = self.factory(identifier)
            session = self._touch(identifier)
            if task_id:
                self.default_task_id = identifier
            self._current.set(identifier)
            self._trim()
            return session

    def current(self) -> Any:
        with self.lock:
            identifier = self._current.get() or self.default_task_id
            if identifier not in self.sessions:
                identifier = self.default_task_id
            return self._touch(identifier)

    def new(self) -> Any:
        with self.lock:
            session = self.factory(None)
            identifier = str(session.task_id)
            self.sessions[identifier] = session
            self.default_task_id = identifier
            self._current.set(identifier)
            self._trim()
            return session

    def restore(self, task_id: str) -> Any:
        return self.bind(task_id, load=True)

    def reindex(self, old_task_id: str, session: Any) -> None:
        with self.lock:
            self.sessions.pop(str(old_task_id), None)
            identifier = str(session.task_id)
            self.sessions[identifier] = session
            self.default_task_id = identifier
            self._current.set(identifier)

    def close(self) -> None:
        with self.lock:
            sessions = list(self.sessions.values())
            self.sessions.clear()
        for session in sessions:
            try:
                session.close()
            except Exception:
                pass

    def loaded(self) -> list[str]:
        with self.lock:
            return list(self.sessions)


class SessionProxy:
    """Attribute-compatible proxy so existing handlers remain unchanged."""

    _internal = {"registry"}

    def __init__(self, registry: SessionRegistry) -> None:
        object.__setattr__(self, "registry", registry)

    def __getattr__(self, name: str) -> Any:
        if name == "reset":
            return self.reset
        if name == "restore":
            return self.restore
        if name == "close":
            return self.registry.close
        return getattr(self.registry.current(), name)

    def __setattr__(self, name: str, value: Any) -> None:
        if name in self._internal:
            object.__setattr__(self, name, value)
        else:
            setattr(self.registry.current(), name, value)

    def bind(self, task_id: str | None) -> Any:
        return self.registry.bind(task_id, load=bool(task_id))

    def reset(self, *, initial: bool = False) -> None:
        del initial
        self.registry.new()

    def restore(self, task_id: str) -> None:
        self.registry.restore(task_id)


__all__ = ["SessionProxy", "SessionRegistry"]
