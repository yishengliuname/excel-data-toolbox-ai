"""Typed capability registry shared by AI planning, UI and execution gates."""

from __future__ import annotations

import json
import re
import threading
from dataclasses import asdict, dataclass
from typing import Any, Callable, Mapping, Sequence

Validator = Callable[[Mapping[str, Any]], None]
Executor = Callable[..., Any]


@dataclass(frozen=True)
class ToolDefinition:
    name: str
    title: str
    category: str
    description: str
    parameter_schema: Mapping[str, Any]
    output_schema: Mapping[str, Any]
    risk_level: str = "low"
    requires_confirmation: bool = False
    supports_preview: bool = True
    supports_background: bool = True
    version: str = "1.0.0"
    enabled: bool = True

    def __post_init__(self) -> None:
        if not re.fullmatch(r"[a-z][a-z0-9_]{1,79}", self.name):
            raise ValueError("工具名称无效")
        if self.risk_level not in {"low", "medium", "high"}:
            raise ValueError("工具风险等级无效")
        json.dumps(dict(self.parameter_schema), ensure_ascii=False, allow_nan=False)
        json.dumps(dict(self.output_schema), ensure_ascii=False, allow_nan=False)

    def to_dict(self) -> dict[str, Any]:
        payload = asdict(self)
        payload["parameter_schema"] = dict(self.parameter_schema)
        payload["output_schema"] = dict(self.output_schema)
        return payload


@dataclass
class RegisteredTool:
    definition: ToolDefinition
    validator: Validator | None = None
    executor: Executor | None = None


class ToolRegistry:
    def __init__(self) -> None:
        self._lock = threading.RLock()
        self._tools: dict[str, RegisteredTool] = {}

    def register(
        self,
        definition: ToolDefinition,
        *,
        validator: Validator | None = None,
        executor: Executor | None = None,
        replace: bool = False,
    ) -> None:
        with self._lock:
            if definition.name in self._tools and not replace:
                raise ValueError(f"工具已注册：{definition.name}")
            self._tools[definition.name] = RegisteredTool(definition, validator, executor)

    def get(self, name: str) -> RegisteredTool:
        with self._lock:
            try:
                item = self._tools[str(name)]
            except KeyError as exc:
                raise KeyError(f"工具未注册：{name}") from exc
            if not item.definition.enabled:
                raise RuntimeError(f"工具已禁用：{name}")
            return item

    def validate(self, name: str, params: Mapping[str, Any]) -> None:
        item = self.get(name)
        if not isinstance(params, Mapping):
            raise TypeError("工具参数必须是对象")
        schema = item.definition.parameter_schema
        allowed = set(map(str, schema.get("properties", {}))) if isinstance(schema, Mapping) else set()
        required = set(map(str, schema.get("required", ()))) if isinstance(schema, Mapping) else set()
        if allowed:
            unknown = set(map(str, params)) - allowed
            if unknown:
                raise ValueError(f"工具 {name} 包含未知参数：{', '.join(sorted(unknown))}")
        missing = [key for key in required if key not in params]
        if missing:
            raise ValueError(f"工具 {name} 缺少参数：{', '.join(sorted(missing))}")
        if item.validator:
            item.validator(params)

    def execute(self, name: str, params: Mapping[str, Any], *args: Any, **kwargs: Any) -> Any:
        item = self.get(name)
        self.validate(name, params)
        if item.executor is None:
            raise RuntimeError(f"工具 {name} 仅登记能力，尚未绑定执行器")
        return item.executor(params, *args, **kwargs)

    def manifest(self, *, enabled_only: bool = True) -> list[dict[str, Any]]:
        with self._lock:
            definitions = [item.definition for item in self._tools.values()]
        if enabled_only:
            definitions = [item for item in definitions if item.enabled]
        return [item.to_dict() for item in sorted(definitions, key=lambda value: (value.category, value.name))]

    def names(self) -> frozenset[str]:
        with self._lock:
            return frozenset(name for name, item in self._tools.items() if item.definition.enabled)


def _object_schema(properties: Sequence[str] = (), required: Sequence[str] = ()) -> dict[str, Any]:
    return {
        "type": "object",
        "additionalProperties": False,
        "properties": {name: {} for name in properties},
        "required": list(required),
    }


def build_builtin_registry(operation_names: Sequence[str]) -> ToolRegistry:
    registry = ToolRegistry()
    high_risk = {"vba", "database", "business_decision", "power_bi"}
    medium_risk = {"join", "lookup", "fuzzy_lookup", "reconcile", "compare", "mask", "split"}
    categories = {
        "clean": "数据处理", "replace": "数据处理", "select_rename_sort": "数据处理",
        "concat": "数据处理", "join": "数据关联", "lookup": "数据关联", "summary": "统计分析",
        "quality": "质量治理", "validate": "质量治理", "reconcile": "财务对账",
        "chart": "可视化", "finance": "财务分析", "adaptive_analysis_report": "经营报告",
        "sales_management_report": "经营报告", "quarterly_sales_report": "经营报告",
        "inventory_management_report": "经营报告", "hr_management_report": "经营报告",
        "selection_recommendation_report": "经营报告", "enterprise_diagnosis_report": "经营报告",
    }
    common_parameters = (
        "columns", "rename", "sort_by", "by", "aggregations", "rules", "task", "user_request",
        "include_charts", "top_n", "date_column", "value_columns", "category_columns", "value_column",
        "on", "left_on", "right_on", "how", "column", "source_key", "lookup_key", "value_columns",
        "name", "steps", "columns", "threshold", "tolerance", "date_tolerance_days", "output_name",
    )
    for raw in sorted(set(map(str, operation_names))):
        risk = "high" if raw in high_risk else "medium" if raw in medium_risk else "low"
        registry.register(ToolDefinition(
            name=raw,
            title=raw.replace("_", " "),
            category=categories.get(raw, "AI 白名单能力"),
            description=f"表格快处本地白名单操作：{raw}",
            parameter_schema=_object_schema(tuple(dict.fromkeys(common_parameters))),
            output_schema={"type": "object", "description": "本地结构化结果或交付表"},
            risk_level=risk,
            requires_confirmation=risk in {"medium", "high"},
            supports_preview=raw not in high_risk,
            supports_background=True,
        ))
    return registry


__all__ = ["RegisteredTool", "ToolDefinition", "ToolRegistry", "build_builtin_registry"]
