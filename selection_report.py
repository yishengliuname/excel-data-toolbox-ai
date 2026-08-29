"""Deterministic, auditable candidate selection for unfamiliar workbooks.

The natural-language model may identify that a user wants to choose a number
of candidates, but it never decides who is selected.  This module infers the
identifier, repeated score and narrative-review columns from the uploaded
table, calculates a transparent ranking, and exposes every penalty and tie
breaker for human review.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
import math
import re
from typing import Any

import pandas as pd


_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:/.]+")
_IDENTIFIER_EXACT = {
    "序号", "编号", "候选编号", "候选人编号", "作品编号", "项目编号", "队伍编号",
    "姓名", "候选人", "选手", "作品", "项目", "团队", "队伍", "id", "no", "code",
}
_IDENTIFIER_HINT = re.compile(r"序号|编号|编码|候选|选手|姓名|作品|项目|团队|队伍|工号|^id$|^no$|code", re.I)
_SCORE_HINT = re.compile(r"得分|分数|成绩|评分|绩效|评审分|综合分|总分|score|grade|rating|points?", re.I)
_COMMENT_HINT = re.compile(r"问题|评语|评价|意见|备注|风险|缺陷|说明|反馈|comment|remark|issue|risk|note", re.I)
_ROUND_HINT = re.compile(r"第?\s*([0-9一二三四五六七八九十百]+)\s*(?:轮|次|阶段|期|回合)", re.I)

_SEVERE_RULES: tuple[tuple[str, re.Pattern[str], float], ...] = (
    ("重复/疑似重复", re.compile(r"(?:与|和)\s*\d+\s*(?:组|号)?.{0,3}重复|重复稿|稿件重复|内容重复|雷同|抄袭|剽窃", re.I), 12.0),
    ("诚信或有效性风险", re.compile(r"作弊|造假|无效|取消资格|不合格", re.I), 12.0),
    ("核心结果错误", re.compile(r"(?:核心|主要|关键)(?:模型|结论|结果|程序).{0,6}(?:错误|有误|不可信|偏离|缺陷)|(?:结论|结果|模型|程序).{0,4}(?:错误|有误|不可信|偏离|重大缺陷)", re.I), 8.0),
    ("态度/完成质量严重风险", re.compile(r"态度.{0,6}(?:问题|不端正)|完成很差|极度不端正|来搞笑", re.I), 8.0),
    ("直接判定为纯AI", re.compile(r"纯\s*ai", re.I), 6.0),
)
_MODERATE_RULES: tuple[tuple[str, re.Pattern[str], float], ...] = (
    ("明显AI痕迹", re.compile(r"ai.{0,6}(?:痕迹|嫌疑|撰写)|全文ai|怀疑全文ai", re.I), 1.5),
    ("内容或论证不足", re.compile(r"内容.{0,5}(?:缺失|偏少|不足)|分析.{0,6}(?:浅|少|不足)|论证.{0,5}不足", re.I), 1.0),
    ("严重格式/可读性问题", re.compile(r"格式问题很大|排版差|看不清|字号偏小", re.I), 0.75),
)
_POSITIVE_RULES: tuple[tuple[str, re.Pattern[str], float], ...] = (
    ("主要结果合理", re.compile(r"主要结果.{0,8}合理|结果总体合理|关键结果.{0,8}合理", re.I), 2.5),
    ("完成度较高", re.compile(r"完成度较高|完成情况很好|结构完整|质量较好", re.I), 1.5),
)


def _normalise(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


def _text(value: Any, *, limit: int = 800) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()[:limit]


def _numeric(series: pd.Series) -> tuple[pd.Series, float]:
    if pd.api.types.is_numeric_dtype(series.dtype) and not pd.api.types.is_bool_dtype(series.dtype):
        values = pd.to_numeric(series, errors="coerce")
        return values, float(values.notna().mean())
    raw = series.astype("string").str.strip()
    cleaned = raw.str.replace(r"[，,分%％\s]", "", regex=True)
    values = pd.to_numeric(cleaned, errors="coerce")
    denominator = int(raw.ne("").sum())
    return values, float(values.notna().sum() / denominator) if denominator else 0.0


def _chinese_integer(text: str) -> int | None:
    text = str(text or "").strip()
    if text.isdigit():
        return int(text)
    digits = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
              "六": 6, "七": 7, "八": 8, "九": 9}
    if len(text) == 1 and text in digits:
        return digits[text]
    if not text or any(character not in digits and character not in "十百千" for character in text):
        return None
    total = 0
    pending = 0
    for character in text:
        if character in digits:
            pending = digits[character]
            continue
        unit = {"十": 10, "百": 100, "千": 1000}[character]
        total += (pending or 1) * unit
        pending = 0
    return total + pending


_SELECTION_COUNT_PATTERNS = (
    # Natural descriptions are allowed between the action and the requested
    # count, e.g. “选取最优秀的八个组” or “挑出综合表现最好的 10 人”.
    r"(?:选(?:择|出|取)?|挑(?:选|出|取)?|推荐|确定|筛选|遴选).{0,20}?([0-9一二三四五六七八九十百千两]+)(?:个|名|人|组|支|队|项|份)",
    r"前([0-9一二三四五六七八九十百千两]+)(?:个|名|人|组|支|队|项|份)?",
    r"top\s*([0-9]+)",
)


def explicit_selection_count(prompt: str, *, maximum: int = 1000) -> int | None:
    """Return the explicitly requested count, or ``None`` when absent."""

    folded = re.sub(r"\s+", "", str(prompt or ""))
    for pattern in _SELECTION_COUNT_PATTERNS:
        match = re.search(pattern, folded, flags=re.I)
        if match:
            parsed = _chinese_integer(match.group(1))
            if parsed is not None and parsed > 0:
                return min(parsed, maximum)
    return None


def parse_selection_count(prompt: str, *, default: int = 8, maximum: int = 1000) -> int:
    """Extract phrases such as ``选8个``、``前十名`` or ``选取最优秀的八组``."""

    parsed = explicit_selection_count(prompt, maximum=maximum)
    if parsed is not None:
        return parsed
    return min(max(int(default), 1), maximum)


def _round_number(name: str, fallback: int) -> int:
    match = _ROUND_HINT.search(str(name))
    if match:
        parsed = _chinese_integer(match.group(1))
        if parsed is not None:
            return parsed
    leading = re.search(r"(?:^|\D)(\d{1,3})(?:\D|$)", str(name))
    return int(leading.group(1)) if leading else fallback


def infer_selection_columns(frame: pd.DataFrame) -> dict[str, Any]:
    """Infer candidate identifier, score columns and narrative review fields."""

    if not isinstance(frame, pd.DataFrame) or frame.empty:
        raise ValueError("候选数据表不能为空")
    columns = [str(column).strip() for column in frame.columns]
    if not columns:
        raise ValueError("候选数据表没有字段")

    numeric: dict[str, tuple[pd.Series, float]] = {column: _numeric(frame[column]) for column in columns}
    score_columns = [
        column for column in columns
        if _SCORE_HINT.search(column) and numeric[column][1] >= 0.45
    ]
    if not score_columns:
        raise ValueError("未识别到可计算的得分/成绩/评分字段")
    score_columns.sort(key=lambda column: (_round_number(column, columns.index(column) + 1), columns.index(column)))

    identifier_candidates: list[tuple[float, str]] = []
    rows = max(len(frame), 1)
    for index, column in enumerate(columns):
        if column in score_columns or _COMMENT_HINT.search(column):
            continue
        normalized = _normalise(column)
        nonblank = frame[column].dropna().astype("string").str.strip()
        nonblank = nonblank[nonblank.ne("")]
        uniqueness = float(nonblank.nunique() / max(len(nonblank), 1))
        name_score = 12.0 if normalized in {_normalise(item) for item in _IDENTIFIER_EXACT} else 0.0
        if _IDENTIFIER_HINT.search(column):
            name_score += 7.0
        name_score += uniqueness * 4.0 + (1.0 - index / max(len(columns), 1))
        if len(nonblank) >= max(2, math.ceil(rows * 0.5)):
            identifier_candidates.append((name_score, column))
    if not identifier_candidates:
        raise ValueError("未识别到序号、编号、姓名、作品或项目等候选标识字段")
    identifier_column = max(identifier_candidates, key=lambda item: item[0])[1]
    comment_columns = [column for column in columns if _COMMENT_HINT.search(column)]
    return {
        "identifier_column": identifier_column,
        "score_columns": score_columns,
        "comment_columns": comment_columns,
    }


def can_build_selection_report(frames: Sequence[pd.DataFrame]) -> bool:
    candidates = 0
    for frame in frames:
        try:
            infer_selection_columns(frame)
        except (TypeError, ValueError):
            continue
        candidates += 1
    return candidates == 1


def validate_selection_report_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("候选评选参数必须是对象")
    names = params.get("source_names")
    if not isinstance(names, (list, tuple)) or not names or not all(isinstance(item, str) and item.strip() for item in names):
        raise TypeError("source_names 必须是非空字符串数组")
    request = params.get("user_request", "")
    if not isinstance(request, str) or len(request) > 8_000:
        raise TypeError("user_request 必须是不超过8000字符的文本")
    top_n = params.get("top_n", 8)
    if isinstance(top_n, bool) or not isinstance(top_n, int) or not 1 <= top_n <= 1000:
        raise ValueError("top_n 必须是1到1000之间的整数")
    if not isinstance(params.get("include_charts", True), bool):
        raise TypeError("include_charts 必须是布尔值")


@dataclass(frozen=True)
class SelectionReportResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def _normalise_score(values: pd.Series) -> tuple[pd.Series, str]:
    valid = values.dropna()
    if valid.empty:
        return values.astype(float), "无有效分数"
    maximum = float(valid.max())
    minimum = float(valid.min())
    if 0 <= minimum and maximum <= 1.05:
        return values * 100.0, "0~1按百分制×100"
    if 0 <= minimum and maximum <= 5.05:
        return values * 20.0, "5分制按百分制×20"
    if 0 <= minimum and maximum <= 10.05:
        return values * 10.0, "10分制按百分制×10"
    return values.astype(float), "按原始分值（默认同量纲）"


def _risk_evaluation(text: str) -> tuple[float, float, list[str], list[str]]:
    penalty = 0.0
    bonus = 0.0
    risks: list[str] = []
    positives: list[str] = []
    for label, pattern, weight in _SEVERE_RULES:
        if pattern.search(text):
            penalty += weight
            risks.append(label)
    for label, pattern, weight in _MODERATE_RULES:
        if pattern.search(text):
            penalty += weight
            risks.append(label)
    for label, pattern, weight in _POSITIVE_RULES:
        if pattern.search(text):
            bonus += weight
            positives.append(label)
    return min(penalty, 30.0), min(bonus, 5.0), list(dict.fromkeys(risks)), positives


def _selection_reason(row: pd.Series) -> str:
    reasons = [
        f"有效均分{row['有效平均分']:.2f}",
        f"最新得分{row['最新得分']:.2f}",
    ]
    if row["得分趋势"] >= 5:
        reasons.append("近期明显提升")
    elif row["得分趋势"] <= -10:
        reasons.append("近期下降需复核")
    if row["正向依据"]:
        reasons.append(str(row["正向依据"]))
    if row["风险提示"]:
        reasons.append("风险：" + str(row["风险提示"]))
    return "；".join(reasons)


def build_selection_recommendation_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    user_request: str = "",
    top_n: int = 8,
    include_charts: bool = True,
) -> SelectionReportResult:
    """Rank candidates and return an audit-ready selection package."""

    validate_selection_report_params({
        "source_names": source_names,
        "user_request": user_request,
        "top_n": top_n,
        "include_charts": include_charts,
    })
    if len(frames) != len(source_names):
        raise ValueError("source_names 数量必须与输入表数量一致")
    candidates: list[tuple[pd.DataFrame, str, dict[str, Any]]] = []
    for frame, name in zip(frames, source_names):
        try:
            inferred = infer_selection_columns(frame)
        except (TypeError, ValueError):
            continue
        candidates.append((frame.copy(deep=True), name, inferred))
    if len(candidates) != 1:
        raise ValueError("必须且只能识别出一张候选评分主表")

    frame, source_name, inferred = candidates[0]
    identifier_column = inferred["identifier_column"]
    score_columns: list[str] = inferred["score_columns"]
    comment_columns: list[str] = inferred["comment_columns"]
    top_n = min(top_n, len(frame))

    normalized_scores: dict[str, pd.Series] = {}
    normalization_notes: dict[str, str] = {}
    for column in score_columns:
        raw, _ = _numeric(frame[column])
        normalized_scores[column], normalization_notes[column] = _normalise_score(raw)
    score_frame = pd.DataFrame(normalized_scores, index=frame.index)

    rows: list[dict[str, Any]] = []
    for source_index, source_row in frame.iterrows():
        candidate_id = _text(source_row.get(identifier_column), limit=120) or f"第{int(source_index) + 2}行"
        valid = score_frame.loc[source_index].dropna()
        comments = "；".join(
            _text(source_row.get(column), limit=1000)
            for column in comment_columns
            if _text(source_row.get(column), limit=1000)
        )
        penalty, bonus, risks, positives = _risk_evaluation(comments)
        if valid.empty:
            average = latest = first = float("nan")
            coverage = 0.0
            base = float("nan")
            composite = float("-inf")
            risks = [*risks, "无有效得分"]
        else:
            average = float(valid.mean())
            first = float(valid.iloc[0])
            latest = float(valid.iloc[-1])
            coverage = float(len(valid) / len(score_columns))
            base = 0.70 * average + 0.30 * latest
            missing_penalty = (1.0 - coverage) * 5.0
            composite = base - penalty - missing_penalty + bonus
        risk_level = "高" if penalty >= 12 or not valid.size else ("中" if penalty >= 4 else "低")
        rows.append({
            "源行号": int(source_index) + 2,
            identifier_column: candidate_id,
            "综合推荐分": composite,
            "基础表现分": base,
            "有效平均分": average,
            "最新得分": latest,
            "得分趋势": latest - first if valid.size else float("nan"),
            "有效得分轮数": int(len(valid)),
            "得分完整率": coverage,
            "风险扣分": penalty,
            "正向加分": bonus,
            "风险等级": risk_level,
            "风险提示": "、".join(risks),
            "正向依据": "、".join(positives),
            "评语摘要": comments[:600],
        })

    ranking = pd.DataFrame(rows)
    ranking["__有效__"] = ranking["综合推荐分"].map(math.isfinite)
    ranking = ranking.sort_values(
        ["__有效__", "综合推荐分", "最新得分", "有效平均分", "有效得分轮数", "源行号"],
        ascending=[False, False, False, False, False, True],
        kind="stable",
        na_position="last",
    ).reset_index(drop=True)
    ranking.insert(0, "综合排名", range(1, len(ranking) + 1))
    selectable = int(ranking["__有效__"].sum())
    selected_count = min(top_n, selectable)
    ranking["入选状态"] = ["建议入选" if index < selected_count else "候补/未入选" for index in range(len(ranking))]
    ranking["推荐理由"] = ranking.apply(_selection_reason, axis=1)
    ranking = ranking.drop(columns="__有效__")

    selected_columns = [
        "综合排名", identifier_column, "综合推荐分", "有效平均分", "最新得分", "得分趋势",
        "有效得分轮数", "得分完整率", "风险等级", "风险提示", "推荐理由", "源行号",
    ]
    selected = ranking.head(selected_count).loc[:, selected_columns].copy()
    selected.insert(0, "入选顺序", range(1, len(selected) + 1))
    review = ranking.loc[
        ranking["风险等级"].isin(["中", "高"]),
        ["综合排名", identifier_column, "入选状态", "风险等级", "风险提示", "评语摘要", "源行号"],
    ].copy()

    rules_rows = [
        {"类别": "任务识别", "项目": "用户需求", "内容": user_request or "从候选评分表中选出指定数量"},
        {"类别": "字段识别", "项目": "候选标识", "内容": identifier_column},
        {"类别": "字段识别", "项目": "得分字段", "内容": "、".join(score_columns)},
        {"类别": "字段识别", "项目": "评语字段", "内容": "、".join(comment_columns) or "未识别"},
        {"类别": "评分公式", "项目": "基础表现分", "内容": "有效平均分×70%＋最新有效得分×30%"},
        {"类别": "评分公式", "项目": "综合推荐分", "内容": "基础表现分－文本风险扣分－缺失分数扣分＋正向评价加分"},
        {"类别": "缺失处理", "项目": "分数缺失", "内容": "按现有有效分数计算；完整率每缺10%扣0.5分；无有效得分不入选"},
        {"类别": "风险规则", "项目": "严重风险", "内容": "重复/抄袭、诚信或有效性、核心结果错误、严重态度问题、纯AI等只作排序扣分和人工复核提示"},
        {"类别": "人工边界", "项目": "最终决定", "内容": "程序提供可解释推荐，不自动替代赛事资格、原创性和专业结论的人工审核"},
    ]
    for column in score_columns:
        rules_rows.append({"类别": "分值归一", "项目": column, "内容": normalization_notes[column]})
    rules = pd.DataFrame(rules_rows)

    selected_ids = "、".join(selected[identifier_column].astype("string").tolist())
    overview = pd.DataFrame([
        {"指标": "评选任务", "结果": _text(user_request, limit=500) or "候选对象结构化评选", "单位": "", "数据口径": "自然语言自动识别"},
        {"指标": "来源数据表", "结果": source_name, "单位": "", "数据口径": "唯一满足候选标识+得分字段的工作表"},
        {"指标": "候选总数", "结果": len(ranking), "单位": "个", "数据口径": "主表记录数"},
        {"指标": "目标入选数", "结果": top_n, "单位": "个", "数据口径": "从自然语言提取；未提供时默认8"},
        {"指标": "实际建议入选", "结果": selected_count, "单位": "个", "数据口径": "存在有效得分的候选中按综合推荐分排序"},
        {"指标": "建议入选名单", "结果": selected_ids, "单位": "", "数据口径": "顺序即综合推荐顺序"},
        {"指标": "识别得分字段", "结果": len(score_columns), "单位": "列", "数据口径": "字段名+数值可解析率"},
        {"指标": "识别评语字段", "结果": len(comment_columns), "单位": "列", "数据口径": "问题/评语/备注/风险/反馈等字段"},
        {"指标": "需人工复核", "结果": len(review), "单位": "个", "数据口径": "中高风险评语触发；不代表自动淘汰"},
    ])

    dashboard = ranking.head(min(len(ranking), 30)).loc[
        :, [identifier_column, "综合推荐分", "有效平均分", "最新得分", "风险扣分", "入选状态"]
    ].copy()
    outputs: dict[str, pd.DataFrame] = {
        "评选管理总览": overview,
        "建议入选名单": selected,
        "全部候选排序": ranking,
        "风险复核清单": review,
        "评选规则与字段": rules,
    }
    if include_charts:
        outputs["评选图表看板"] = dashboard
    for output in outputs.values():
        output.attrs["toolbox_report_kind"] = "selection_recommendation_report"
    report = {
        "source_table": source_name,
        "identifier_column": identifier_column,
        "score_columns": score_columns,
        "comment_columns": comment_columns,
        "candidate_count": len(ranking),
        "requested_count": top_n,
        "selected_count": selected_count,
        "review_count": len(review),
        "selected_ids": selected[identifier_column].astype("string").tolist(),
        "sheet_count": len(outputs),
        "chart_count": 2 if include_charts else 0,
    }
    return SelectionReportResult(outputs=outputs, report=report)


__all__ = [
    "SelectionReportResult",
    "build_selection_recommendation_report",
    "can_build_selection_report",
    "explicit_selection_count",
    "infer_selection_columns",
    "parse_selection_count",
    "validate_selection_report_params",
]
