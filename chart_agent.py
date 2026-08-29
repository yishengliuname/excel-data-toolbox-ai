"""Constrained natural-language chart specifications.

The model may choose and edit presentation parameters, but it can never return
or execute Python, JavaScript, SQL, HTML, formulas, URLs, or arbitrary code.
Every accepted field is validated against the local table catalogue before the
existing deterministic chart engine receives it.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
import math
import re
from typing import Any


CHART_SPEC_VERSION = 1
CHART_TYPES = frozenset(
    {
        "bar", "horizontal_bar", "grouped_bar", "stacked_bar", "line", "area",
        "pie", "radar", "funnel", "waterfall", "treemap", "histogram",
        "scatter", "box", "heatmap", "gantt",
    }
)
AGGREGATIONS = frozenset({"sum", "count", "mean", "nunique", "max", "min"})
DATE_GRAINS = frozenset({"auto", "day", "week", "month", "quarter", "year"})
THEMES = frozenset({"default", "business_dark", "economist", "swiss", "finance", "warm", "minimal"})
NUMBER_FORMATS = frozenset({"auto", "number", "currency", "percent", "wan", "yi"})
SORT_MODES = frozenset({"auto", "asc", "desc", "source"})
STATUS_VALUES = frozenset({"ready", "clarification", "unsupported"})
HEX_COLOR = re.compile(r"^#[0-9A-Fa-f]{6}$")

_TOP_KEYS = {
    "schema_version", "status", "normalized_request", "message",
    "clarification_questions", "warnings", "chart",
}
_CHART_KEYS = {
    "chart_type", "dimension", "measure", "measures", "series", "aggregation", "top_n",
    "date_grain", "start", "end", "progress", "style_3d", "title", "theme",
    "number_format", "sort", "reference_lines", "highlight", "show_labels",
    "show_legend", "x_axis_label", "y_axis_label", "series_colors",
    "background_color", "text_color", "font_size", "legend_position",
    "label_rotation", "show_grid", "opacity", "bar_gap", "chart_height",
    "y_min", "y_max",
}

_VERSION_ALIASES = frozenset({"1", "1.0", "v1", "v1.0"})
_CHART_DEFAULTS: dict[str, Any] = {
    "chart_type": None,
    "dimension": None,
    "measure": None,
    "measures": [],
    "series": None,
    "aggregation": "sum",
    "top_n": 20,
    "date_grain": "auto",
    "start": None,
    "end": None,
    "progress": None,
    "style_3d": False,
    "title": None,
    "theme": "default",
    "number_format": "auto",
    "sort": "auto",
    "reference_lines": [],
    "highlight": None,
    "show_labels": True,
    "show_legend": True,
    "x_axis_label": None,
    "y_axis_label": None,
    "series_colors": [],
    "background_color": "#FFFFFF",
    "text_color": "#243831",
    "font_size": 12,
    "legend_position": "bottom",
    "label_rotation": 0,
    "show_grid": True,
    "opacity": 0.92,
    "bar_gap": 0.22,
    "chart_height": 340,
    "y_min": None,
    "y_max": None,
}

_COLOR_ALIASES = {
    "red": "#E53935", "红": "#E53935", "红色": "#E53935",
    "green": "#2E7D32", "绿": "#2E7D32", "绿色": "#2E7D32",
    "yellow": "#FBC02D", "黄": "#FBC02D", "黄色": "#FBC02D",
    "blue": "#1976D2", "蓝": "#1976D2", "蓝色": "#1976D2",
    "purple": "#7B1FA2", "紫": "#7B1FA2", "紫色": "#7B1FA2",
    "orange": "#F57C00", "橙": "#F57C00", "橙色": "#F57C00",
    "black": "#212121", "黑": "#212121", "黑色": "#212121",
}


class ChartSpecValidationError(ValueError):
    """Raised when a model-created chart spec is unsafe or inconsistent."""


def normalize_chart_spec_envelope(payload: Mapping[str, Any]) -> dict[str, Any]:
    """Repair harmless model drift around versions and display defaults.

    DeepSeek occasionally serializes version 1 as ``"1.0"``/``"v1"``, uses
    the older ``version`` key, or omits the version while returning every other
    required field.  Those forms have identical semantics, so they can be
    canonicalized locally without accepting unknown executable fields or a
    genuinely newer schema.  Missing presentation-only fields are filled with
    deterministic local defaults; data fields still have to pass the normal
    table-catalogue checks.
    """

    if not isinstance(payload, Mapping):
        raise ChartSpecValidationError("AI 图表计划必须是 JSON 对象")
    normalized = dict(payload)
    if not isinstance(normalized.get("normalized_request"), str) or not normalized.get("normalized_request", "").strip():
        normalized["normalized_request"] = "已根据用户需求生成图表方案"
    if not isinstance(normalized.get("message"), str) or not normalized.get("message", "").strip():
        normalized["message"] = "AI 已完成图表参数规划"
    if normalized.get("clarification_questions") is None:
        normalized["clarification_questions"] = []
    if normalized.get("warnings") is None:
        normalized["warnings"] = []
    if "schema_version" not in normalized and "version" in normalized:
        normalized["schema_version"] = normalized.pop("version")
    elif "version" in normalized:
        raise ChartSpecValidationError("AI 图表计划不能同时包含 version 和 schema_version")

    if "schema_version" not in normalized:
        if set(normalized) == _TOP_KEYS - {"schema_version"}:
            normalized["schema_version"] = CHART_SPEC_VERSION
        else:
            return normalized

    version = normalized.get("schema_version")
    if isinstance(version, str) and version.strip().lower() in _VERSION_ALIASES:
        normalized["schema_version"] = CHART_SPEC_VERSION
    elif isinstance(version, float) and version == float(CHART_SPEC_VERSION):
        normalized["schema_version"] = CHART_SPEC_VERSION

    raw_chart = normalized.get("chart")
    if normalized.get("status") == "ready" and isinstance(raw_chart, Mapping):
        chart = dict(raw_chart)
        raw_series = chart.get("series")
        raw_measure = chart.get("measure")
        if isinstance(raw_series, Sequence) and not isinstance(raw_series, (str, bytes)):
            if not chart.get("measures"):
                chart["measures"] = list(raw_series)
            chart["series"] = None
        if isinstance(raw_measure, Sequence) and not isinstance(raw_measure, (str, bytes)):
            if not chart.get("measures"):
                chart["measures"] = list(raw_measure)
            chart["measure"] = None
        raw_colors = chart.get("series_colors")
        measure_order = chart.get("measures")
        if isinstance(raw_colors, Mapping) and isinstance(measure_order, Sequence) and not isinstance(measure_order, (str, bytes)):
            chart["series_colors"] = [raw_colors.get(str(measure)) for measure in measure_order]
        if not set(chart).issubset(_CHART_KEYS):
            return normalized
        for key, value in _CHART_DEFAULTS.items():
            if key not in chart:
                chart[key] = list(value) if isinstance(value, list) else value
        for key in (
            "background_color", "text_color", "font_size", "legend_position",
            "label_rotation", "show_grid", "opacity", "bar_gap", "chart_height",
        ):
            if chart.get(key) is None or chart.get(key) == "":
                chart[key] = _CHART_DEFAULTS[key]
        top_n_value = chart.get("top_n")
        if top_n_value is None or top_n_value == "" or top_n_value == 0:
            chart["top_n"] = _CHART_DEFAULTS["top_n"]
        elif (
            isinstance(top_n_value, float)
            and not isinstance(top_n_value, bool)
            and top_n_value.is_integer()
        ):
            chart["top_n"] = int(top_n_value)
        if isinstance(chart.get("title"), str) and not chart["title"].strip():
            chart["title"] = None
        normalized["chart"] = chart
    return normalized


def _text(value: Any, name: str, *, maximum: int, allow_empty: bool = False) -> str:
    if not isinstance(value, str):
        raise ChartSpecValidationError(f"{name} 必须是文本")
    result = value.strip()
    if not allow_empty and not result:
        raise ChartSpecValidationError(f"{name} 不能为空")
    if len(result) > maximum:
        raise ChartSpecValidationError(f"{name} 最长 {maximum} 个字符")
    return result


def _string_list(value: Any, name: str, *, maximum_items: int = 8) -> list[str]:
    if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
        raise ChartSpecValidationError(f"{name} 必须是文本列表")
    if len(value) > maximum_items:
        raise ChartSpecValidationError(f"{name} 最多 {maximum_items} 项")
    return [_text(item, f"{name}[{index}]", maximum=300) for index, item in enumerate(value)]


def _catalog_columns(catalog: Mapping[str, Any]) -> set[str]:
    if not isinstance(catalog, Mapping) or set(catalog) != {"catalog_version", "tables"}:
        raise ChartSpecValidationError("数据目录结构无效")
    tables = catalog.get("tables")
    if not isinstance(tables, list) or len(tables) != 1 or not isinstance(tables[0], Mapping):
        raise ChartSpecValidationError("AI 可视化一次只能使用一张数据表")
    columns = tables[0].get("columns")
    if not isinstance(columns, list):
        raise ChartSpecValidationError("数据目录缺少字段")
    result: set[str] = set()
    for item in columns:
        if not isinstance(item, Mapping) or not isinstance(item.get("name"), str):
            raise ChartSpecValidationError("数据目录字段结构无效")
        result.add(item["name"])
    return result


def _optional_column(value: Any, name: str, columns: set[str]) -> str | None:
    if value is None or value == "":
        return None
    result = _text(value, name, maximum=200)
    if result not in columns:
        raise ChartSpecValidationError(f"{name} 引用了不存在的字段“{result}”")
    return result


def _column_list(value: Any, name: str, columns: set[str], *, maximum_items: int = 8) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
        raise ChartSpecValidationError(f"{name} 必须是字段名列表")
    if len(value) > maximum_items:
        raise ChartSpecValidationError(f"{name} 最多包含 {maximum_items} 个字段")
    result = [_text(item, f"{name}[{index}]", maximum=200) for index, item in enumerate(value)]
    if len(result) != len(set(result)):
        raise ChartSpecValidationError(f"{name} 不能包含重复字段")
    missing = [item for item in result if item not in columns]
    if missing:
        raise ChartSpecValidationError(f"{name} 引用了不存在的字段“{missing[0]}”")
    return result


def _optional_label(value: Any, name: str) -> str | None:
    if value is None or value == "":
        return None
    return _text(value, name, maximum=40)


def _color_list(value: Any, name: str, *, maximum_items: int = 8) -> list[str]:
    if value is None:
        return []
    if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
        raise ChartSpecValidationError(f"{name} 必须是颜色列表")
    if len(value) > maximum_items:
        raise ChartSpecValidationError(f"{name} 最多包含 {maximum_items} 个颜色")
    result: list[str] = []
    for index, item in enumerate(value):
        color = _text(item, f"{name}[{index}]", maximum=30)
        canonical = _COLOR_ALIASES.get(color.lower(), _COLOR_ALIASES.get(color, color))
        if not HEX_COLOR.fullmatch(canonical):
            raise ChartSpecValidationError(f"{name}[{index}] 必须是颜色名称或 #RRGGBB")
        result.append(canonical.upper())
    return result


def _color_value(value: Any, name: str) -> str:
    color = _text(value, name, maximum=30)
    canonical = _COLOR_ALIASES.get(color.lower(), _COLOR_ALIASES.get(color, color))
    if not HEX_COLOR.fullmatch(canonical):
        raise ChartSpecValidationError(f"{name} 必须是颜色名称或 #RRGGBB")
    return canonical.upper()


def _bounded_number(value: Any, name: str, minimum: float, maximum: float) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        raise ChartSpecValidationError(f"{name} 必须是数字")
    result = float(value)
    if not minimum <= result <= maximum:
        raise ChartSpecValidationError(f"{name} 必须在 {minimum}~{maximum} 之间")
    return result


def _bounded_int(value: Any, name: str, minimum: int, maximum: int) -> int:
    if isinstance(value, bool) or not isinstance(value, (int, float)) or int(value) != value:
        raise ChartSpecValidationError(f"{name} 必须是整数")
    result = int(value)
    if not minimum <= result <= maximum:
        raise ChartSpecValidationError(f"{name} 必须在 {minimum}~{maximum} 之间")
    return result


def _optional_finite(value: Any, name: str) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        raise ChartSpecValidationError(f"{name} 必须是有限数字或 null")
    return float(value)


def validate_chart_spec(payload: Mapping[str, Any], catalog: Mapping[str, Any]) -> dict[str, Any]:
    """Return a normalized, JSON-safe chart specification."""

    payload = normalize_chart_spec_envelope(payload)
    if not isinstance(payload, Mapping) or set(payload) != _TOP_KEYS:
        raise ChartSpecValidationError("AI 图表计划顶层字段不完整或包含未知字段")
    if payload.get("schema_version") != CHART_SPEC_VERSION:
        raise ChartSpecValidationError("AI 图表计划版本不受支持")
    status = payload.get("status")
    if status not in STATUS_VALUES:
        raise ChartSpecValidationError("status 无效")
    result: dict[str, Any] = {
        "schema_version": CHART_SPEC_VERSION,
        "status": status,
        "normalized_request": _text(payload.get("normalized_request"), "normalized_request", maximum=1000),
        "message": _text(payload.get("message"), "message", maximum=1000),
        "clarification_questions": _string_list(payload.get("clarification_questions"), "clarification_questions"),
        "warnings": _string_list(payload.get("warnings"), "warnings"),
        "chart": None,
    }
    if status != "ready":
        if payload.get("chart") is not None:
            raise ChartSpecValidationError("非 ready 状态不能包含图表规格")
        if status == "clarification" and not result["clarification_questions"]:
            raise ChartSpecValidationError("需要澄清时必须提出具体问题")
        return result

    raw = payload.get("chart")
    if not isinstance(raw, Mapping) or set(raw) != _CHART_KEYS:
        raise ChartSpecValidationError("chart 字段不完整或包含未知字段")
    columns = _catalog_columns(catalog)
    chart_type = raw.get("chart_type")
    if chart_type not in CHART_TYPES:
        raise ChartSpecValidationError("chart_type 不受支持")
    aggregation = raw.get("aggregation")
    date_grain = raw.get("date_grain")
    theme = raw.get("theme")
    number_format = raw.get("number_format")
    sort = raw.get("sort")
    if aggregation not in AGGREGATIONS or date_grain not in DATE_GRAINS:
        raise ChartSpecValidationError("统计方式或日期粒度无效")
    if theme not in THEMES or number_format not in NUMBER_FORMATS or sort not in SORT_MODES:
        raise ChartSpecValidationError("主题、数字格式或排序方式无效")
    top_n = raw.get("top_n")
    if isinstance(top_n, bool) or not isinstance(top_n, int) or not 1 <= top_n <= 50:
        raise ChartSpecValidationError("top_n 必须是 1~50 的整数")
    for key in ("style_3d", "show_labels", "show_legend"):
        if not isinstance(raw.get(key), bool):
            raise ChartSpecValidationError(f"{key} 必须是布尔值")

    dimension = _optional_column(raw.get("dimension"), "dimension", columns)
    measures = _column_list(raw.get("measures"), "measures", columns)
    measure = measures[0] if measures else _optional_column(raw.get("measure"), "measure", columns)
    series = _optional_column(raw.get("series"), "series", columns)
    start = _optional_column(raw.get("start"), "start", columns)
    end = _optional_column(raw.get("end"), "end", columns)
    progress = _optional_column(raw.get("progress"), "progress", columns)
    if chart_type == "gantt":
        if not dimension or not start or not end:
            raise ChartSpecValidationError("甘特图必须指定任务、开始日期和结束日期字段")
    else:
        if not measure:
            raise ChartSpecValidationError("当前图表必须指定指标字段")
        if chart_type not in {"histogram", "box"} and not dimension:
            raise ChartSpecValidationError("当前图表必须指定维度字段")
    if measures:
        if chart_type == "bar":
            chart_type = "grouped_bar"
        if chart_type not in {"grouped_bar", "stacked_bar", "radar", "heatmap"}:
            raise ChartSpecValidationError("多指标对比仅支持分组柱状图、堆叠柱状图、雷达图或热力图")
        if series:
            raise ChartSpecValidationError("多指标字段与分类系列字段不能同时使用")
        if dimension in measures:
            raise ChartSpecValidationError("横轴字段不能同时作为指标字段")
    if chart_type in {"grouped_bar", "stacked_bar", "radar", "heatmap"} and not series and not measures:
        raise ChartSpecValidationError("当前图表必须指定系列字段")
    if chart_type == "scatter" and dimension == measure:
        raise ChartSpecValidationError("散点图横纵轴不能相同")

    refs = raw.get("reference_lines")
    if not isinstance(refs, list) or len(refs) > 5:
        raise ChartSpecValidationError("reference_lines 必须是不超过 5 项的列表")
    reference_lines: list[dict[str, Any]] = []
    for index, item in enumerate(refs):
        if not isinstance(item, Mapping) or set(item) != {"value", "label", "color"}:
            raise ChartSpecValidationError(f"reference_lines[{index}] 结构无效")
        value = item.get("value")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)):
            raise ChartSpecValidationError(f"reference_lines[{index}].value 必须是有限数字")
        color = item.get("color")
        if not isinstance(color, str) or not HEX_COLOR.fullmatch(color):
            raise ChartSpecValidationError(f"reference_lines[{index}].color 必须是 #RRGGBB")
        reference_lines.append({"value": float(value), "label": _text(item.get("label"), f"reference_lines[{index}].label", maximum=80), "color": color.upper()})

    highlight = raw.get("highlight")
    safe_highlight = None
    if highlight is not None:
        if not isinstance(highlight, Mapping) or set(highlight) != {"field", "value", "color"}:
            raise ChartSpecValidationError("highlight 结构无效")
        field = _optional_column(highlight.get("field"), "highlight.field", columns)
        color = highlight.get("color")
        if not field or not isinstance(color, str) or not HEX_COLOR.fullmatch(color):
            raise ChartSpecValidationError("highlight 字段或颜色无效")
        safe_highlight = {"field": field, "value": _text(highlight.get("value"), "highlight.value", maximum=200), "color": color.upper()}

    title = raw.get("title")
    if title is not None:
        title = _text(title, "title", maximum=80)
    x_axis_label = _optional_label(raw.get("x_axis_label"), "x_axis_label") or dimension
    default_y_axis = "分数" if measures and all(re.search(r"得分|分数|成绩", item) for item in measures) else (measure or "数值")
    y_axis_label = _optional_label(raw.get("y_axis_label"), "y_axis_label") or default_y_axis
    series_colors = _color_list(raw.get("series_colors"), "series_colors")
    if measures and series_colors and len(series_colors) != len(measures):
        raise ChartSpecValidationError("多指标颜色数量必须与指标字段数量一致")
    background_color = _color_value(raw.get("background_color"), "background_color")
    text_color = _color_value(raw.get("text_color"), "text_color")
    font_size = _bounded_int(raw.get("font_size"), "font_size", 10, 24)
    legend_position = raw.get("legend_position")
    if legend_position not in {"top", "bottom", "left", "right"}:
        raise ChartSpecValidationError("legend_position 必须是 top、bottom、left 或 right")
    label_rotation = _bounded_int(raw.get("label_rotation"), "label_rotation", -90, 90)
    if not isinstance(raw.get("show_grid"), bool):
        raise ChartSpecValidationError("show_grid 必须是布尔值")
    opacity = _bounded_number(raw.get("opacity"), "opacity", 0.2, 1.0)
    bar_gap = _bounded_number(raw.get("bar_gap"), "bar_gap", 0.0, 0.8)
    chart_height = _bounded_int(raw.get("chart_height"), "chart_height", 240, 600)
    y_min = _optional_finite(raw.get("y_min"), "y_min")
    y_max = _optional_finite(raw.get("y_max"), "y_max")
    if y_min is not None and y_max is not None and y_min >= y_max:
        raise ChartSpecValidationError("y_min 必须小于 y_max")
    result["chart"] = {
        "chart_type": chart_type, "dimension": dimension, "measure": measure,
        "measures": measures, "series": series, "aggregation": aggregation, "top_n": top_n,
        "date_grain": date_grain, "start": start, "end": end, "progress": progress,
        "style_3d": raw["style_3d"], "title": title, "theme": theme,
        "number_format": number_format, "sort": sort,
        "reference_lines": reference_lines, "highlight": safe_highlight,
        "show_labels": raw["show_labels"], "show_legend": raw["show_legend"],
        "x_axis_label": x_axis_label, "y_axis_label": y_axis_label,
        "series_colors": series_colors,
        "background_color": background_color, "text_color": text_color,
        "font_size": font_size, "legend_position": legend_position,
        "label_rotation": label_rotation, "show_grid": raw["show_grid"],
        "opacity": opacity, "bar_gap": bar_gap, "chart_height": chart_height,
        "y_min": y_min, "y_max": y_max,
    }
    return result


CHART_SYSTEM_PROMPT = """你是表格快处的 AI 可视化规划器。你只能返回严格 JSON，绝不能返回代码、公式、HTML、SQL、URL 或解释性围栏。
目标：把用户自然语言转换为安全图表规格；若传入 current_spec，则基于它连续修改，未要求改变的字段保持不变。
只能使用目录中真实字段，只能选择提供的枚举。无法判断关键字段时 status=clarification 并提出最多3个短问题；能力之外则 status=unsupported。
顶层必须且只能包含 schema_version,status,normalized_request,message,clarification_questions,warnings,chart。schema_version 必须使用 JSON 数字 1（不能写成字符串、1.0、v1 或其他版本）。
ready 时 chart 必须且只能包含：chart_type,dimension,measure,measures,series,aggregation,top_n,date_grain,start,end,progress,style_3d,title,theme,number_format,sort,reference_lines,highlight,show_labels,show_legend,x_axis_label,y_axis_label,series_colors,background_color,text_color,font_size,legend_position,label_rotation,show_grid,opacity,bar_gap,chart_height,y_min,y_max。
单指标图：measure 是一个字段名，measures=[]。多轮得分、多年份指标等宽表对比：measures 放 2~8 个数值字段、measure 可填首个指标、series=null，并优先使用 grouped_bar；series 只用于“地区/渠道”等单个分类字段，绝不能把字段列表放入 series。
x_axis_label、y_axis_label 必须忠实执行用户指定的坐标轴名称。series_colors 是与 measures 严格一一对应的颜色数组；用户指定“第一红、第二绿、第三黄、第四蓝、第五紫”时必须返回 ["#E53935","#2E7D32","#FBC02D","#1976D2","#7B1FA2"]，不能改用主题默认色。
其他可控样式：background_color/text_color 使用颜色；font_size 10~24；legend_position=top/bottom/left/right；label_rotation=-90~90；show_grid 布尔值；opacity=0.2~1；bar_gap=0~0.8；chart_height=240~600；y_min/y_max 可为数字或 null。必须优先执行用户的明确样式要求，未指定时使用默认值。
reference_lines 每项必须为 {value,label,color}，color 为 #RRGGBB；highlight 为 null 或 {field,value,color}。当用户要求高亮最高/最大/峰值项时，highlight.value 必须写为字符串 "__MAX__"，由本地程序根据实际汇总结果求最大值，绝不能猜测具体月份或类别。
支持图表：bar,horizontal_bar,grouped_bar,stacked_bar,line,area,pie,radar,funnel,waterfall,treemap,histogram,scatter,box,heatmap,gantt。
统计：sum,count,mean,nunique,max,min。日期粒度：auto,day,week,month,quarter,year。主题：default,business_dark,economist,swiss,finance,warm,minimal。数字格式：auto,number,currency,percent,wan,yi。排序：auto,asc,desc,source。
clarification/unsupported 时 chart 必须为 null。所有布尔字段必须明确给出。"""
