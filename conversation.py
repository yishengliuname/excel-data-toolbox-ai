"""Task-scoped follow-up context for iterative spreadsheet analysis."""

from __future__ import annotations

import json
import re
import tempfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping, Sequence

CONVERSATION_SCHEMA_VERSION = 1
_FOLLOWUP_HINT = re.compile(r"刚才|上次|继续|再|改成|换成|只看|增加|删除|这个|那个|同样|沿用|还要|另外|把它|前面", re.I)
_SECRET = re.compile(r"\bsk-[A-Za-z0-9_-]{6,}\b", re.I)


def _safe_text(value: Any, maximum: int) -> str:
    return _SECRET.sub("[API Key 已隐藏]", str(value or "").strip())[:maximum]


def _safe_json(value: Any, *, depth: int = 0) -> Any:
    if depth > 5:
        return "[depth-limit]"
    if value is None or isinstance(value, (bool, int, float)):
        return value
    if isinstance(value, str):
        return _safe_text(value, 2000)
    if isinstance(value, Mapping):
        return {str(key)[:100]: _safe_json(item, depth=depth + 1) for key, item in list(value.items())[:100]}
    if isinstance(value, (list, tuple)):
        return [_safe_json(item, depth=depth + 1) for item in list(value)[:100]]
    if hasattr(value, "to_dict"):
        return _safe_json(value.to_dict(), depth=depth + 1)
    return _safe_text(value, 2000)


@dataclass(frozen=True)
class ConversationTurn:
    created_at: str
    user_request: str
    normalized_request: str
    route: str
    status: str
    plan_summary: str
    output_names: tuple[str, ...] = ()
    chart_spec: Mapping[str, Any] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["output_names"] = list(self.output_names)
        payload["chart_spec"] = dict(self.chart_spec)
        return payload


class ConversationStore:
    def __init__(self, path: str | Path, *, max_turns: int = 30) -> None:
        self.path = Path(path)
        self.max_turns = max(1, min(int(max_turns), 200))

    def list(self) -> list[ConversationTurn]:
        if not self.path.exists():
            return []
        try:
            payload = json.loads(self.path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return []
        if not isinstance(payload, Mapping) or int(payload.get("schema_version", 0)) != CONVERSATION_SCHEMA_VERSION:
            return []
        turns = []
        for raw in payload.get("turns") or []:
            if not isinstance(raw, Mapping):
                continue
            turns.append(ConversationTurn(
                created_at=_safe_text(raw.get("created_at"), 80),
                user_request=_safe_text(raw.get("user_request"), 8000),
                normalized_request=_safe_text(raw.get("normalized_request"), 8000),
                route=_safe_text(raw.get("route"), 80),
                status=_safe_text(raw.get("status"), 40),
                plan_summary=_safe_text(raw.get("plan_summary"), 2000),
                output_names=tuple(_safe_text(item, 200) for item in (raw.get("output_names") or ())[:80]),
                chart_spec=_safe_json(raw.get("chart_spec") or {}),
            ))
        return turns[-self.max_turns:]

    def append(
        self,
        *,
        user_request: str,
        normalized_request: str = "",
        route: str = "",
        status: str = "",
        plan_summary: str = "",
        output_names: Sequence[str] = (),
        chart_spec: Mapping[str, Any] | None = None,
    ) -> ConversationTurn:
        turn = ConversationTurn(
            created_at=datetime.now().isoformat(timespec="seconds"),
            user_request=_safe_text(user_request, 8000),
            normalized_request=_safe_text(normalized_request or user_request, 8000),
            route=_safe_text(route, 80), status=_safe_text(status, 40),
            plan_summary=_safe_text(plan_summary, 2000),
            output_names=tuple(_safe_text(item, 200) for item in output_names[:80]),
            chart_spec=_safe_json(chart_spec or {}),
        )
        turns = [*self.list(), turn][-self.max_turns:]
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"schema_version": CONVERSATION_SCHEMA_VERSION, "updated_at": turn.created_at, "turns": [item.to_dict() for item in turns]}
        with tempfile.NamedTemporaryFile(dir=self.path.parent, prefix=".conversation_", suffix=".tmp", delete=False, mode="w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2, allow_nan=False)
            temporary = Path(stream.name)
        temporary.replace(self.path)
        return turn

    def clear(self) -> None:
        self.path.unlink(missing_ok=True)

    def context(self) -> dict[str, Any]:
        turns = self.list()
        latest = turns[-1] if turns else None
        return {
            "turn_count": len(turns),
            "last_request": latest.normalized_request if latest else "",
            "last_route": latest.route if latest else "",
            "last_status": latest.status if latest else "",
            "last_plan_summary": latest.plan_summary if latest else "",
            "last_output_names": list(latest.output_names) if latest else [],
            "last_chart_spec": dict(latest.chart_spec) if latest else {},
        }

    def resolve(self, request: str) -> tuple[str, bool]:
        prompt = _safe_text(request, 8000)
        if not prompt:
            raise ValueError("需求不能为空")
        context = self.context()
        previous = str(context.get("last_request") or "")
        is_followup = bool(previous and (_FOLLOWUP_HINT.search(prompt) or len(prompt) <= 32))
        if not is_followup:
            return prompt, False
        route = str(context.get("last_route") or "")
        outputs = "、".join(map(str, context.get("last_output_names") or []))
        combined = (
            f"上一轮已确认任务：{previous}\n"
            f"上一轮任务类型：{route or '未标注'}；上一轮输出：{outputs or '未记录'}。\n"
            f"本轮是后续修改要求：{prompt}\n"
            "请沿用上一轮已确认的数据范围和业务口径，只修改本轮明确提出的内容；如果本轮与上一轮冲突，以本轮为准。"
        )
        return combined[:8000], True


__all__ = ["CONVERSATION_SCHEMA_VERSION", "ConversationStore", "ConversationTurn"]
