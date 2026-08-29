"""Deterministic employee attendance, performance and payroll reporting.

DeepSeek can recognise broad intent, but this allow-listed workflow identifies
the uploaded table roles from their columns and performs every calculation
locally.  The report deliberately labels retention risk as a management proxy,
not as an employment decision or a psychological assessment.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
import math
import re
from typing import Any

import pandas as pd


_NON_WORD = re.compile(r"[\s_\-（）()【】\[\]：:]+")


def _normalise_name(value: Any) -> str:
    return _NON_WORD.sub("", str(value or "")).casefold()


def _find_column(frame: pd.DataFrame, aliases: Sequence[str]) -> str | None:
    normalised = {_normalise_name(column): str(column) for column in frame.columns}
    for alias in aliases:
        key = _normalise_name(alias)
        if key in normalised:
            return normalised[key]
    candidates = {
        column
        for key, column in normalised.items()
        for alias in aliases
        if len(_normalise_name(alias)) >= 2 and _normalise_name(alias) in key
    }
    return next(iter(candidates)) if len(candidates) == 1 else None


def _has_columns(frame: pd.DataFrame, groups: Sequence[Sequence[str]]) -> bool:
    return all(_find_column(frame, aliases) is not None for aliases in groups)


def infer_hr_table_roles(frames: Sequence[pd.DataFrame]) -> dict[str, Any]:
    """Identify HR table roles by columns, supporting multiple period sheets."""

    if not isinstance(frames, Sequence) or isinstance(frames, (str, bytes)):
        raise TypeError("员工经营分析输入必须是表格数组")
    signatures = {
        "employees": (("员工编号", "工号"), ("姓名", "员工姓名"), ("部门",), ("基本工资", "底薪")),
        "attendance": (("员工编号", "工号"), ("出勤天数", "实际出勤"), ("迟到次数", "迟到"), ("请假天数", "请假")),
        "performance": (("员工编号", "工号"), ("完成目标比例", "目标完成率", "绩效完成率"), ("客户评分", "客户满意度", "绩效评分")),
        "adjustments": (("员工编号", "工号"), ("调整类型", "薪资项目", "项目"), ("金额", "调整金额")),
    }
    matches: dict[str, list[int]] = {}
    for role, groups in signatures.items():
        matches[role] = [
            index
            for index, frame in enumerate(frames)
            if isinstance(frame, pd.DataFrame) and _has_columns(frame, groups)
        ]
    if len(matches["employees"]) != 1:
        raise ValueError("无法唯一识别员工基础信息表")
    for role, label in (("attendance", "考勤"), ("performance", "绩效"), ("adjustments", "薪资调整")):
        if not matches[role]:
            raise ValueError(f"无法识别{label}数据表")
    return {
        "employees": matches["employees"][0],
        "attendance": tuple(matches["attendance"]),
        "performance": tuple(matches["performance"]),
        "adjustments": tuple(matches["adjustments"]),
        "notes": tuple(
            index
            for index, frame in enumerate(frames)
            if index not in {matches["employees"][0], *matches["attendance"], *matches["performance"], *matches["adjustments"]}
        ),
    }


def can_build_hr_report(frames: Sequence[pd.DataFrame]) -> bool:
    try:
        infer_hr_table_roles(frames)
    except (TypeError, ValueError):
        return False
    return True


def validate_hr_report_params(params: Mapping[str, Any]) -> None:
    if not isinstance(params, Mapping):
        raise TypeError("员工经营报告参数必须是对象")
    source_names = params.get("source_names")
    if not isinstance(source_names, (list, tuple)) or not source_names or not all(
        isinstance(item, str) and item.strip() for item in source_names
    ):
        raise TypeError("source_names 必须是非空字符串数组")
    for key, default, minimum, maximum in (
        ("expected_workdays", 22, 1, 31),
        ("excellent_score", 85, 50, 100),
        ("attention_score", 70, 0, 95),
    ):
        value = params.get(key, default)
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise TypeError(f"{key} 必须是数字")
        if not minimum <= float(value) <= maximum:
            raise ValueError(f"{key} 必须在 {minimum} 到 {maximum} 之间")
    if float(params.get("attention_score", 70)) >= float(params.get("excellent_score", 85)):
        raise ValueError("attention_score 必须小于 excellent_score")


def _clean_text(value: Any) -> str:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _clean_code(value: Any) -> str:
    return re.sub(r"\s+", "", _clean_text(value)).upper()


def _number(value: Any) -> float:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return float("nan")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = re.sub(r"[¥￥元,%％,\s]", "", _clean_text(value))
    if not text:
        return float("nan")
    try:
        result = float(text)
    except ValueError:
        return float("nan")
    if "%" in _clean_text(value) or "％" in _clean_text(value):
        result /= 100.0
    return result


def _columns(frame: pd.DataFrame, mapping: Mapping[str, Sequence[str]]) -> dict[str, str | None]:
    return {key: _find_column(frame, aliases) for key, aliases in mapping.items()}


def _concat_role(frames: Sequence[pd.DataFrame], indexes: Sequence[int], names: Sequence[str]) -> pd.DataFrame:
    pieces = []
    for index in indexes:
        piece = frames[index].copy(deep=True)
        piece["__源工作表"] = names[index]
        piece["__源行号"] = range(2, len(piece) + 2)
        pieces.append(piece)
    return pd.concat(pieces, ignore_index=True, sort=False)


@dataclass(frozen=True)
class HRReportResult:
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def build_hr_management_report(
    frames: Sequence[pd.DataFrame],
    *,
    source_names: Sequence[str],
    expected_workdays: int = 22,
    excellent_score: float = 85,
    attention_score: float = 70,
) -> HRReportResult:
    """Create a boss-ready, auditable employee operating report."""

    validate_hr_report_params({
        "source_names": source_names,
        "expected_workdays": expected_workdays,
        "excellent_score": excellent_score,
        "attention_score": attention_score,
    })
    if len(frames) != len(source_names):
        raise ValueError("source_names 数量必须与输入表数量一致")
    roles = infer_hr_table_roles(frames)
    names = list(source_names)
    employee_frame = frames[roles["employees"]].copy(deep=True)
    attendance_frame = _concat_role(frames, roles["attendance"], names)
    performance_frame = _concat_role(frames, roles["performance"], names)
    adjustment_frame = _concat_role(frames, roles["adjustments"], names)

    emp_cols = _columns(employee_frame, {
        "code": ("员工编号", "工号"), "name": ("姓名", "员工姓名"), "department": ("部门",),
        "position": ("岗位", "职位"), "entry": ("入职日期",), "base": ("基本工资", "底薪"), "status": ("状态", "员工状态"),
    })
    att_cols = _columns(attendance_frame, {
        "code": ("员工编号", "工号"), "month": ("月份", "期间", "考勤月份"), "days": ("出勤天数", "实际出勤"),
        "late": ("迟到次数", "迟到"), "early": ("早退次数", "早退"), "leave": ("请假天数", "请假"), "overtime": ("加班小时", "加班时数"),
    })
    perf_cols = _columns(performance_frame, {
        "code": ("员工编号", "工号"), "sales": ("销售额", "业绩金额"), "target": ("完成目标比例", "目标完成率", "绩效完成率"),
        "rating": ("客户评分", "客户满意度", "绩效评分"),
    })
    adj_cols = _columns(adjustment_frame, {
        "code": ("员工编号", "工号"), "type": ("调整类型", "薪资项目", "项目"), "amount": ("金额", "调整金额"), "remark": ("备注", "说明"),
    })

    audit_rows: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []

    employees = pd.DataFrame({
        "员工编号": employee_frame[emp_cols["code"]].map(_clean_code),
        "姓名": employee_frame[emp_cols["name"]].map(_clean_text),
        "部门": employee_frame[emp_cols["department"]].map(_clean_text),
        "岗位": employee_frame[emp_cols["position"]].map(_clean_text) if emp_cols["position"] else "",
        "入职日期": pd.to_datetime(employee_frame[emp_cols["entry"]], errors="coerce") if emp_cols["entry"] else pd.NaT,
        "基本工资": employee_frame[emp_cols["base"]].map(_number),
        "状态": employee_frame[emp_cols["status"]].map(_clean_text) if emp_cols["status"] else "在职",
    })
    duplicate_employees = employees[employees["员工编号"].ne("") & employees.duplicated("员工编号", keep="first")]
    for _, row in duplicate_employees.iterrows():
        review_rows.append({"员工编号": row["员工编号"], "姓名": row["姓名"], "事项类型": "主档重复", "严重程度": "高", "事实依据": "员工编号在基础信息中重复", "建议动作": "核对唯一员工主档并合并重复记录", "源工作表": names[roles["employees"]]})
    employees = employees.loc[employees["员工编号"].ne("")].drop_duplicates("员工编号", keep="first").reset_index(drop=True)
    known_codes = set(employees["员工编号"])

    attendance = pd.DataFrame({
        "员工编号": attendance_frame[att_cols["code"]].map(_clean_code),
        "月份": attendance_frame[att_cols["month"]].map(_clean_text) if att_cols["month"] else "未提供",
        "出勤天数": attendance_frame[att_cols["days"]].map(_number),
        "迟到次数": attendance_frame[att_cols["late"]].map(_number),
        "早退次数": attendance_frame[att_cols["early"]].map(_number) if att_cols["early"] else 0.0,
        "请假天数": attendance_frame[att_cols["leave"]].map(_number),
        "加班小时": attendance_frame[att_cols["overtime"]].map(_number) if att_cols["overtime"] else 0.0,
        "源工作表": attendance_frame["__源工作表"],
    }).fillna({"出勤天数": 0, "迟到次数": 0, "早退次数": 0, "请假天数": 0, "加班小时": 0})
    performance = pd.DataFrame({
        "员工编号": performance_frame[perf_cols["code"]].map(_clean_code),
        "销售额": performance_frame[perf_cols["sales"]].map(_number) if perf_cols["sales"] else 0.0,
        "目标完成率": performance_frame[perf_cols["target"]].map(_number),
        "客户评分": performance_frame[perf_cols["rating"]].map(_number),
        "源工作表": performance_frame["__源工作表"],
    })
    adjustments = pd.DataFrame({
        "员工编号": adjustment_frame[adj_cols["code"]].map(_clean_code),
        "调整类型": adjustment_frame[adj_cols["type"]].map(_clean_text),
        "调整金额": adjustment_frame[adj_cols["amount"]].map(_number),
        "备注": adjustment_frame[adj_cols["remark"]].map(_clean_text) if adj_cols["remark"] else "",
        "源工作表": adjustment_frame["__源工作表"],
    }).fillna({"调整金额": 0.0})

    for label, frame in (("考勤", attendance), ("绩效", performance), ("薪资调整", adjustments)):
        for _, row in frame.loc[~frame["员工编号"].isin(known_codes)].iterrows():
            review_rows.append({"员工编号": row["员工编号"], "姓名": "", "事项类型": f"未知员工{label}记录", "严重程度": "高", "事实依据": "员工编号未出现在员工基础信息", "建议动作": "核对工号或补充员工主档后重新运行", "源工作表": row["源工作表"]})
        frame.drop(frame.index[~frame["员工编号"].isin(known_codes)], inplace=True)

    periods = max(int(attendance["月份"].replace("", pd.NA).nunique(dropna=True)), 1)
    att = attendance.groupby("员工编号", as_index=False, observed=True).agg(
        月份数=("月份", "nunique"), 出勤天数=("出勤天数", "sum"), 迟到次数=("迟到次数", "sum"),
        早退次数=("早退次数", "sum"), 请假天数=("请假天数", "sum"), 加班小时=("加班小时", "sum"),
    )
    perf = performance.groupby("员工编号", as_index=False, observed=True).agg(
        销售额=("销售额", "sum"), 目标完成率=("目标完成率", "mean"), 客户评分=("客户评分", "mean"),
    )
    adj = adjustments.groupby("员工编号", as_index=False, observed=True).agg(
        薪资调整金额=("调整金额", "sum"), 调整项目数=("调整类型", "count"),
        调整项目=("调整类型", lambda values: "、".join(dict.fromkeys(item for item in values if item))),
    )
    detail = employees.merge(att, on="员工编号", how="left").merge(perf, on="员工编号", how="left").merge(adj, on="员工编号", how="left")
    for column in ("月份数", "出勤天数", "迟到次数", "早退次数", "请假天数", "加班小时", "销售额", "薪资调整金额", "调整项目数"):
        detail[column] = detail[column].fillna(0.0)
    detail["统计月份数"] = detail["月份数"].where(detail["月份数"].gt(0), periods)
    detail["应出勤天数"] = detail["统计月份数"] * float(expected_workdays)
    detail["缺勤天数"] = (detail["应出勤天数"] - detail["出勤天数"] - detail["请假天数"]).clip(lower=0)
    detail["出勤率"] = detail["出勤天数"].div(detail["应出勤天数"].where(detail["应出勤天数"].gt(0))).clip(upper=1)
    detail["考勤得分"] = (
        100 - detail["迟到次数"] * 5 - detail["早退次数"] * 5 - detail["请假天数"] * 2 - detail["缺勤天数"] * 10
    ).clip(lower=0, upper=100)
    detail["绩效得分"] = (
        detail["目标完成率"].clip(lower=0, upper=1) * 70 + detail["客户评分"].clip(lower=0, upper=5).div(5) * 30
    )
    detail.loc[detail["目标完成率"].isna() | detail["客户评分"].isna(), "绩效得分"] = float("nan")
    detail["综合得分"] = detail["考勤得分"] * 0.3 + detail["绩效得分"] * 0.7
    detail["预计薪资"] = detail["基本工资"] * detail["统计月份数"] + detail["薪资调整金额"]
    detail["在职标记"] = ~detail["状态"].astype("string").str.contains("离职|停职", na=False)

    active_complete = detail["在职标记"] & detail["综合得分"].notna()
    detail["综合排名"] = pd.NA
    detail.loc[active_complete, "综合排名"] = detail.loc[active_complete, "综合得分"].rank(method="min", ascending=False).astype("Int64")
    detail["管理分类"] = "正常"
    excellent = active_complete & detail["综合得分"].ge(float(excellent_score)) & detail["迟到次数"].le(1) & detail["缺勤天数"].eq(0)
    attention = detail["在职标记"] & (
        detail["综合得分"].lt(float(attention_score)) | detail["迟到次数"].ge(5) |
        detail["出勤率"].lt(0.85) | detail["目标完成率"].lt(0.8) | detail["客户评分"].lt(4)
    )
    detail.loc[excellent, "管理分类"] = "表现优秀"
    detail.loc[attention, "管理分类"] = "重点关注"
    detail.loc[~detail["在职标记"], "管理分类"] = "已离职/非在职"
    detail["离职风险代理等级"] = "低"
    high_proxy = detail["在职标记"] & (detail["综合得分"].lt(60) | detail["迟到次数"].ge(8) | detail["出勤率"].lt(0.75))
    medium_proxy = attention & ~high_proxy
    detail.loc[medium_proxy, "离职风险代理等级"] = "中"
    detail.loc[high_proxy, "离职风险代理等级"] = "高"
    detail.loc[~detail["在职标记"], "离职风险代理等级"] = "不适用"

    reasons: list[str] = []
    actions: list[str] = []
    for _, row in detail.iterrows():
        facts = []
        if row["迟到次数"] >= 5: facts.append(f"迟到{int(row['迟到次数'])}次")
        if row["出勤率"] < 0.85: facts.append(f"出勤率{row['出勤率']:.0%}")
        if pd.notna(row["目标完成率"]) and row["目标完成率"] < 0.8: facts.append(f"目标完成率{row['目标完成率']:.0%}")
        if pd.notna(row["客户评分"]) and row["客户评分"] < 4: facts.append(f"客户评分{row['客户评分']:.1f}")
        if pd.notna(row["综合得分"]): facts.append(f"综合得分{row['综合得分']:.1f}")
        reasons.append("；".join(facts) or "考勤与绩效未触发预警")
        if row["管理分类"] == "表现优秀": actions.append("纳入表扬/激励候选，并结合岗位成果做主管复核")
        elif row["管理分类"] == "重点关注": actions.append("安排直属主管面谈，核实考勤与绩效原因并制定30天改进计划")
        elif not row["在职标记"]: actions.append("核对离职日期与当月考勤、薪资结算是否一致")
        else: actions.append("保持常规跟踪")
    detail["事实依据"] = reasons
    detail["建议动作"] = actions

    for _, row in detail.loc[~detail["在职标记"] & detail["出勤天数"].gt(0)].iterrows():
        review_rows.append({"员工编号": row["员工编号"], "姓名": row["姓名"], "事项类型": "离职状态仍有考勤", "严重程度": "高", "事实依据": f"状态为{row['状态']}，但统计期存在{int(row['出勤天数'])}天出勤", "建议动作": "核对离职日期、考勤截止日和薪资结算月份", "源工作表": "员工基础信息/考勤"})
    for _, row in detail.loc[detail["在职标记"] & detail["绩效得分"].isna()].iterrows():
        review_rows.append({"员工编号": row["员工编号"], "姓名": row["姓名"], "事项类型": "在职员工绩效缺失", "严重程度": "中", "事实依据": "目标完成率或客户评分缺失", "建议动作": "补齐绩效记录后重新计算综合得分", "源工作表": "绩效数据"})
    for _, row in detail.loc[detail["薪资调整金额"].abs().gt(detail["基本工资"].abs())].iterrows():
        review_rows.append({"员工编号": row["员工编号"], "姓名": row["姓名"], "事项类型": "薪资调整幅度较大", "严重程度": "中", "事实依据": f"调整金额{row['薪资调整金额']:,.2f}元超过基本工资绝对值", "建议动作": "核对审批单、提成和奖金计算依据", "源工作表": "薪资调整记录"})

    comprehensive_columns = [
        "员工编号", "姓名", "部门", "岗位", "状态", "出勤率", "迟到次数", "早退次数", "请假天数", "缺勤天数",
        "考勤得分", "销售额", "目标完成率", "客户评分", "绩效得分", "基本工资", "薪资调整金额", "预计薪资",
        "综合得分", "综合排名", "管理分类", "离职风险代理等级", "事实依据", "建议动作",
    ]
    comprehensive = detail[comprehensive_columns].sort_values(
        "综合得分", ascending=False, na_position="last", kind="stable"
    )
    comprehensive = comprehensive.reset_index(drop=True)
    excellent_table = detail.loc[excellent, ["员工编号", "姓名", "部门", "岗位", "综合得分", "综合排名", "出勤率", "目标完成率", "客户评分", "预计薪资", "事实依据", "建议动作"]].sort_values("综合得分", ascending=False, kind="stable").reset_index(drop=True)
    attention_table = detail.loc[attention, ["员工编号", "姓名", "部门", "岗位", "综合得分", "出勤率", "迟到次数", "请假天数", "缺勤天数", "目标完成率", "客户评分", "离职风险代理等级", "事实依据", "建议动作"]].sort_values(["离职风险代理等级", "综合得分"], ascending=[False, True], kind="stable").reset_index(drop=True)
    attendance_table = detail[["员工编号", "姓名", "部门", "状态", "统计月份数", "应出勤天数", "出勤天数", "出勤率", "迟到次数", "早退次数", "请假天数", "缺勤天数", "加班小时", "考勤得分"]].sort_values(["状态", "考勤得分"], ascending=[True, True], kind="stable").reset_index(drop=True)
    performance_table = detail[["员工编号", "姓名", "部门", "岗位", "状态", "销售额", "目标完成率", "客户评分", "绩效得分", "综合排名"]].sort_values("绩效得分", ascending=False, na_position="last", kind="stable").reset_index(drop=True)
    salary_table = detail[["员工编号", "姓名", "部门", "岗位", "状态", "基本工资", "统计月份数", "薪资调整金额", "调整项目数", "调整项目", "预计薪资", "销售额"]].copy()
    salary_table["销售额/预计薪资"] = salary_table["销售额"].div(salary_table["预计薪资"].where(salary_table["预计薪资"].gt(0)))
    salary_table = salary_table.sort_values("预计薪资", ascending=False, kind="stable").reset_index(drop=True)

    reviews = pd.DataFrame(review_rows, columns=["员工编号", "姓名", "事项类型", "严重程度", "事实依据", "建议动作", "源工作表"])
    audit_rows.extend([
        {"审计项目": "员工主档", "结果": len(employees), "单位": "人", "处理口径": "员工编号非空并去重"},
        {"审计项目": "统计月份", "结果": periods, "单位": "个月", "处理口径": "根据考勤月份去重识别"},
        {"审计项目": "应出勤天数", "结果": expected_workdays, "单位": "天/人月", "处理口径": "默认22天，可由本地参数调整"},
        {"审计项目": "考勤得分", "结果": "100-迟到×5-早退×5-请假×2-缺勤×10", "单位": "分", "处理口径": "结果限制在0至100分"},
        {"审计项目": "绩效得分", "结果": "目标完成率×70+客户评分/5×30", "单位": "分", "处理口径": "目标完成率最高按100%计分；销售额仅展示不跨岗位计分"},
        {"审计项目": "综合得分", "结果": "考勤30%+绩效70%", "单位": "分", "处理口径": "用于管理排序，不作为劳动人事处分依据"},
        {"审计项目": "优秀阈值", "结果": excellent_score, "单位": "分", "处理口径": "综合得分达标且迟到≤1、无缺勤"},
        {"审计项目": "重点关注阈值", "结果": attention_score, "单位": "分", "处理口径": "综合低于阈值，或迟到≥5、出勤率<85%、目标<80%、客户评分<4"},
        {"审计项目": "离职风险代理", "结果": "高/中/低", "单位": "代理等级", "处理口径": "仅由考勤和绩效预警推导，不代表真实离职意愿"},
        {"审计项目": "人工核验事项", "结果": len(reviews), "单位": "项", "处理口径": "未知工号、离职仍有考勤、在职绩效缺失或薪资调整幅度较大"},
    ])
    audit = pd.DataFrame(audit_rows)

    active = detail[detail["在职标记"]]
    scored = active[active["综合得分"].notna()]
    top_name = "—" if scored.empty else str(scored.sort_values("综合得分", ascending=False).iloc[0]["姓名"])
    overview = pd.DataFrame([
        {"指标": "统计月份", "结果": "、".join(sorted(attendance["月份"].dropna().astype(str).unique())), "单位": "", "数据口径": "考勤记录中的月份"},
        {"指标": "员工总数", "结果": len(employees), "单位": "人", "数据口径": "员工主档按员工编号去重"},
        {"指标": "在职员工数", "结果": int(detail["在职标记"].sum()), "单位": "人", "数据口径": "状态不含离职/停职"},
        {"指标": "预计在职薪资", "结果": float(active["预计薪资"].sum()), "单位": "元", "数据口径": "基本工资×统计月份数+薪资调整金额"},
        {"指标": "平均出勤率", "结果": float(active["出勤率"].mean()), "单位": "%", "数据口径": "在职员工出勤率算术平均"},
        {"指标": "平均目标完成率", "结果": float(active["目标完成率"].mean()), "单位": "%", "数据口径": "在职且有绩效记录员工算术平均"},
        {"指标": "平均综合得分", "结果": float(scored["综合得分"].mean()), "单位": "分", "数据口径": "考勤30%+绩效70%"},
        {"指标": "综合表现第一", "结果": top_name, "单位": "", "数据口径": "在职且绩效完整员工综合得分降序"},
        {"指标": "表现优秀员工", "结果": int(excellent.sum()), "单位": "人", "数据口径": f"综合≥{excellent_score:g}、迟到≤1且无缺勤"},
        {"指标": "重点关注员工", "结果": int(attention.sum()), "单位": "人", "数据口径": "综合/考勤/绩效任一触发保守预警"},
        {"指标": "高风险代理人数", "结果": int(high_proxy.sum()), "单位": "人", "数据口径": "综合<60、迟到≥8或出勤率<75%；非真实离职预测"},
        {"指标": "迟到总次数", "结果": float(active["迟到次数"].sum()), "单位": "次", "数据口径": "仅在职员工统计期汇总"},
        {"指标": "人工核验事项", "结果": len(reviews), "单位": "项", "数据口径": "数据冲突与完整性异常"},
    ])

    dept = active.groupby("部门", as_index=False, observed=True).agg(
        在职人数=("员工编号", "count"), 预计薪资=("预计薪资", "sum"), 平均综合得分=("综合得分", "mean")
    ).sort_values("预计薪资", ascending=False, kind="stable")
    employee_chart = scored[["姓名", "综合得分"]].sort_values("综合得分", ascending=False, kind="stable")
    risk = active.groupby("离职风险代理等级", as_index=False, observed=True)["员工编号"].count().rename(columns={"员工编号": "风险人数"})
    risk["__顺序"] = risk["离职风险代理等级"].map({"高": 1, "中": 2, "低": 3}).fillna(9)
    risk = risk.sort_values("__顺序", kind="stable").drop(columns="__顺序").reset_index(drop=True)
    attendance_chart = active.assign(考勤异常次数=active["迟到次数"] + active["早退次数"])[["姓名", "考勤异常次数"]].sort_values("考勤异常次数", ascending=False, kind="stable")
    chart_rows = max(len(dept), len(employee_chart), len(risk), len(attendance_chart), 1)
    chart_data = pd.DataFrame(index=range(chart_rows))
    for frame, columns in ((dept, ["部门", "在职人数", "预计薪资", "平均综合得分"]), (employee_chart, ["姓名", "综合得分"]), (risk, ["离职风险代理等级", "风险人数"]), (attendance_chart, ["考勤员工", "考勤异常次数"])):
        local = frame.rename(columns={"姓名": "考勤员工"}) if "考勤员工" in columns else frame
        for column in columns:
            chart_data[column] = local[column].reindex(range(chart_rows)) if column in local else pd.NA

    outputs = {
        "管理层人效总览": overview,
        "员工综合分析": comprehensive,
        "表现优秀员工": excellent_table,
        "重点关注员工": attention_table,
        "考勤分析": attendance_table,
        "绩效分析": performance_table,
        "薪资分析": salary_table,
        "人工核验": reviews,
        "数据审计": audit,
        "人力图表看板": chart_data,
    }
    for output in outputs.values():
        output.attrs["toolbox_report_kind"] = "hr_management_report"
    report = {
        "employee_count": len(employees), "active_employee_count": int(detail["在职标记"].sum()),
        "expected_active_payroll": float(active["预计薪资"].sum()), "average_attendance_rate": float(active["出勤率"].mean()),
        "average_composite_score": float(scored["综合得分"].mean()), "top_employee": top_name,
        "excellent_count": int(excellent.sum()), "attention_count": int(attention.sum()),
        "high_proxy_risk_count": int(high_proxy.sum()), "manual_review_count": len(reviews),
        "sheet_count": len(outputs), "chart_count": 4,
    }
    return HRReportResult(outputs=outputs, report=report)


__all__ = [
    "HRReportResult", "build_hr_management_report", "can_build_hr_report",
    "infer_hr_table_roles", "validate_hr_report_params",
]
