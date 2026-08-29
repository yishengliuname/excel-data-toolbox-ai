"""Deterministic finance analysis for AI-planned Excel jobs.

DeepSeek may select a task and map columns, but every number in this module is
computed locally from the uploaded workbook.  The functions deliberately avoid
tax, filing and accounting-policy decisions that require jurisdiction-specific
professional judgement.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any, Mapping

import numpy as np
import pandas as pd


FINANCE_TASKS = frozenset(
    {"ar_aging", "budget_variance", "cash_flow", "financial_ratios", "journal_audit"}
)

_REQUIRED_COLUMNS: dict[str, frozenset[str]] = {
    "ar_aging": frozenset({"due_date", "amount"}),
    "budget_variance": frozenset({"period", "category", "actual", "budget"}),
    "cash_flow": frozenset({"date", "amount"}),
    "financial_ratios": frozenset(),
    "journal_audit": frozenset({"voucher", "debit", "credit"}),
}

_OPTIONAL_COLUMNS: dict[str, frozenset[str]] = {
    "ar_aging": frozenset({"counterparty", "invoice", "paid_amount"}),
    "budget_variance": frozenset(),
    "cash_flow": frozenset({"direction", "category", "counterparty"}),
    "financial_ratios": frozenset(
        {
            "period", "revenue", "gross_profit", "net_profit", "current_assets",
            "current_liabilities", "inventory", "total_assets", "total_liabilities",
            "equity", "operating_cash_flow", "accounts_receivable", "cogs",
        }
    ),
    "journal_audit": frozenset({"date", "account", "description"}),
}


@dataclass(frozen=True)
class FinanceAnalysisResult:
    task: str
    outputs: Mapping[str, pd.DataFrame]
    report: Mapping[str, Any]


def validate_finance_params(params: Mapping[str, Any]) -> None:
    task = params.get("task")
    if task not in FINANCE_TASKS:
        raise ValueError(f"finance.task 必须是 {sorted(FINANCE_TASKS)} 之一")
    columns = params.get("columns")
    if not isinstance(columns, Mapping) or not columns:
        raise ValueError("finance.columns 必须是非空字段映射")
    if len(columns) > 30:
        raise ValueError("finance.columns 最多 30 项")
    allowed = _REQUIRED_COLUMNS[task] | _OPTIONAL_COLUMNS[task]
    unknown = sorted(set(columns) - allowed)
    if unknown:
        raise ValueError(f"finance.columns 含未知业务字段：{unknown}")
    missing = sorted(_REQUIRED_COLUMNS[task] - set(columns))
    if missing:
        raise ValueError(f"finance.columns 缺少字段映射：{missing}")
    for key, value in columns.items():
        if not isinstance(key, str) or not isinstance(value, str) or not value.strip() or len(value) > 256:
            raise ValueError(f"finance.columns.{key} 必须是有效字段名")
    if task == "financial_ratios":
        ratio_inputs = set(columns) - {"period"}
        if len(ratio_inputs) < 2:
            raise ValueError("财务比率分析至少需要映射两个财务指标字段")
    if "as_of_date" in params:
        parsed = pd.to_datetime(params["as_of_date"], errors="coerce")
        if pd.isna(parsed):
            raise ValueError("finance.as_of_date 必须是有效日期")
    if "buckets" in params:
        buckets = params["buckets"]
        if not isinstance(buckets, (list, tuple)) or not buckets or len(buckets) > 10:
            raise ValueError("finance.buckets 必须是 1 至 10 个递增正整数")
        if any(not isinstance(item, int) or isinstance(item, bool) or item <= 0 for item in buckets):
            raise ValueError("finance.buckets 必须是递增正整数")
        if list(buckets) != sorted(set(buckets)):
            raise ValueError("finance.buckets 必须严格递增且不重复")
    if params.get("perspective", "income") not in {"income", "cost"}:
        raise ValueError("finance.perspective 只能是 income 或 cost")
    tolerance = params.get("tolerance", 0.01)
    if not isinstance(tolerance, (int, float)) or isinstance(tolerance, bool) or tolerance < 0:
        raise ValueError("finance.tolerance 必须是非负数字")


def finance_column_names(params: Mapping[str, Any]) -> list[str]:
    columns = params.get("columns")
    if not isinstance(columns, Mapping):
        return []
    return list(dict.fromkeys(value for value in columns.values() if isinstance(value, str)))


def _require_source_columns(frame: pd.DataFrame, columns: Mapping[str, str]) -> None:
    missing = [source for source in columns.values() if source not in frame.columns]
    if missing:
        raise ValueError(f"输入表缺少财务字段：{list(dict.fromkeys(missing))}")


def _number(frame: pd.DataFrame, source: str) -> pd.Series:
    return pd.to_numeric(frame[source], errors="coerce").astype("Float64")


def _safe_divide(numerator: pd.Series, denominator: pd.Series) -> pd.Series:
    result = numerator.astype("Float64").divide(denominator.astype("Float64").replace(0, pd.NA))
    return result.replace([np.inf, -np.inf], pd.NA)


def _aging_bucket(days: Any, thresholds: list[int]) -> str:
    if pd.isna(days):
        return "日期无效"
    value = int(days)
    if value <= 0:
        return "未到期"
    lower = 1
    for upper in thresholds:
        if value <= upper:
            return f"{lower}-{upper}天"
        lower = upper + 1
    return f"{thresholds[-1]}天以上"


def _ar_aging(frame: pd.DataFrame, columns: Mapping[str, str], *, as_of_date: str | None, buckets: list[int]) -> FinanceAnalysisResult:
    detail = frame.copy(deep=True)
    due = pd.to_datetime(detail[columns["due_date"]], errors="coerce").dt.normalize()
    amount = _number(detail, columns["amount"])
    paid = _number(detail, columns["paid_amount"]) if "paid_amount" in columns else pd.Series(0.0, index=detail.index, dtype="Float64")
    as_of = pd.Timestamp(as_of_date).normalize() if as_of_date else pd.Timestamp(date.today()).normalize()
    detail["未结金额"] = (amount.fillna(0) - paid.fillna(0)).clip(lower=0)
    detail["账龄天数"] = (as_of - due).dt.days.astype("Int64")
    detail["账龄区间"] = detail["账龄天数"].map(lambda value: _aging_bucket(value, buckets))
    open_detail = detail.loc[detail["未结金额"] > 0].copy()
    bucket_order = ["未到期"] + [f"{1 if index == 0 else buckets[index - 1] + 1}-{upper}天" for index, upper in enumerate(buckets)] + [f"{buckets[-1]}天以上", "日期无效"]
    bucket_summary = (
        open_detail.groupby("账龄区间", dropna=False, observed=False)["未结金额"]
        .agg(单据数="size", 未结金额="sum").reindex(bucket_order, fill_value=0).reset_index()
    )
    outputs: dict[str, pd.DataFrame] = {"primary": open_detail, "账龄汇总": bucket_summary}
    if "counterparty" in columns:
        counterparty = columns["counterparty"]
        customer = open_detail.pivot_table(index=counterparty, columns="账龄区间", values="未结金额", aggfunc="sum", fill_value=0)
        customer = customer.reindex(columns=bucket_order, fill_value=0)
        customer["未结合计"] = customer.sum(axis=1)
        outputs["客户账龄"] = customer.reset_index().sort_values("未结合计", ascending=False)
    invalid_dates = int(due.isna().sum())
    return FinanceAnalysisResult("ar_aging", outputs, {"as_of_date": str(as_of.date()), "open_items": len(open_detail), "outstanding_total": float(open_detail["未结金额"].sum()), "invalid_date_count": invalid_dates})


def _budget_variance(frame: pd.DataFrame, columns: Mapping[str, str], *, perspective: str) -> FinanceAnalysisResult:
    detail = frame.copy(deep=True)
    actual = _number(detail, columns["actual"])
    budget = _number(detail, columns["budget"])
    detail["差异额"] = actual - budget
    detail["差异率"] = _safe_divide(detail["差异额"], budget.abs())
    favorable = detail["差异额"].ge(0) if perspective == "income" else detail["差异额"].le(0)
    judgment = pd.Series("数据无效", index=detail.index, dtype="string")
    valid = detail["差异额"].notna()
    judgment.loc[valid & favorable.fillna(False)] = "有利"
    judgment.loc[valid & ~favorable.fillna(False)] = "不利"
    detail["差异判断"] = judgment
    keys = [columns["period"], columns["category"]]
    numeric = detail.assign(_actual_amount=actual, _budget_amount=budget)
    summary = numeric.groupby(keys, dropna=False).agg(实际金额=("_actual_amount", "sum"), 预算金额=("_budget_amount", "sum")).reset_index()
    summary["差异额"] = summary["实际金额"] - summary["预算金额"]
    summary["差异率"] = _safe_divide(summary["差异额"], summary["预算金额"].abs())
    period = summary.groupby(columns["period"], dropna=False)[["实际金额", "预算金额", "差异额"]].sum().reset_index()
    period["差异率"] = _safe_divide(period["差异额"], period["预算金额"].abs())
    return FinanceAnalysisResult("budget_variance", {"primary": detail, "预算差异汇总": summary, "期间汇总": period}, {"rows": len(detail), "actual_total": float(actual.fillna(0).sum()), "budget_total": float(budget.fillna(0).sum()), "invalid_amount_rows": int((actual.isna() | budget.isna()).sum()), "perspective": perspective})


def _cash_flow(frame: pd.DataFrame, columns: Mapping[str, str]) -> FinanceAnalysisResult:
    detail = frame.copy(deep=True)
    dates = pd.to_datetime(detail[columns["date"]], errors="coerce")
    amount = _number(detail, columns["amount"])
    if "direction" in columns:
        direction = detail[columns["direction"]].astype("string").str.strip().str.casefold()
        inflow_words = {"收入", "流入", "收款", "借", "in", "inflow", "receipt"}
        outflow_words = {"支出", "流出", "付款", "贷", "out", "outflow", "payment"}
        sign = direction.map(lambda value: 1 if value in inflow_words else (-1 if value in outflow_words else pd.NA)).astype("Float64")
        net = amount.abs() * sign
    else:
        net = amount
    detail["现金流入"] = net.clip(lower=0)
    detail["现金流出"] = (-net.clip(upper=0))
    detail["净现金流"] = net
    detail["月份"] = dates.dt.to_period("M").astype("string")
    valid = detail.loc[dates.notna() & net.notna()].copy()
    monthly = valid.groupby("月份", dropna=False)[["现金流入", "现金流出", "净现金流"]].sum().reset_index().sort_values("月份")
    monthly["累计净现金流"] = monthly["净现金流"].cumsum()
    outputs: dict[str, pd.DataFrame] = {"primary": detail, "月度现金流": monthly}
    if "category" in columns:
        category = valid.groupby(["月份", columns["category"]], dropna=False)[["现金流入", "现金流出", "净现金流"]].sum().reset_index()
        outputs["分类现金流"] = category
    return FinanceAnalysisResult("cash_flow", outputs, {"rows": len(detail), "used_rows": len(valid), "invalid_rows": len(detail) - len(valid), "total_inflow": float(valid["现金流入"].sum()), "total_outflow": float(valid["现金流出"].sum()), "net_cash_flow": float(valid["净现金流"].sum())})


def _financial_ratios(frame: pd.DataFrame, columns: Mapping[str, str]) -> FinanceAnalysisResult:
    result = pd.DataFrame(index=frame.index)
    if "period" in columns:
        result[columns["period"]] = frame[columns["period"]]
    values = {key: _number(frame, source) for key, source in columns.items() if key != "period"}
    formulas = {
        "毛利率": ("gross_profit", "revenue"), "净利率": ("net_profit", "revenue"),
        "流动比率": ("current_assets", "current_liabilities"), "资产负债率": ("total_liabilities", "total_assets"),
        "总资产收益率ROA": ("net_profit", "total_assets"), "净资产收益率ROE": ("net_profit", "equity"),
        "经营现金流比率": ("operating_cash_flow", "current_liabilities"),
        "应收账款周转率": ("revenue", "accounts_receivable"), "存货周转率": ("cogs", "inventory"),
    }
    for label, (numerator, denominator) in formulas.items():
        if numerator in values and denominator in values:
            result[label] = _safe_divide(values[numerator], values[denominator])
    if {"current_assets", "inventory", "current_liabilities"} <= set(values):
        result["速动比率"] = _safe_divide(values["current_assets"] - values["inventory"], values["current_liabilities"])
    if result.shape[1] == (1 if "period" in columns else 0):
        raise ValueError("当前字段组合不足以计算任何受支持的财务比率")
    long_rows: list[dict[str, Any]] = []
    period_name = columns.get("period")
    for column in result.columns:
        if column == period_name:
            continue
        for index, value in result[column].items():
            row = {"指标": column, "指标值": value}
            if period_name:
                row[period_name] = result.at[index, period_name]
            long_rows.append(row)
    long_result = pd.DataFrame(long_rows)
    return FinanceAnalysisResult("financial_ratios", {"primary": result, "财务比率长表": long_result}, {"rows": len(result), "ratio_count": result.shape[1] - (1 if period_name else 0), "invalid_ratio_cells": int(result.drop(columns=[period_name] if period_name else []).isna().sum().sum())})


def _journal_audit(frame: pd.DataFrame, columns: Mapping[str, str], *, tolerance: float) -> FinanceAnalysisResult:
    detail = frame.copy(deep=True)
    debit = _number(detail, columns["debit"])
    credit = _number(detail, columns["credit"])
    detail["借方金额_标准"] = debit
    detail["贷方金额_标准"] = credit
    invalid = debit.isna() | credit.isna() | debit.lt(0) | credit.lt(0)
    if "account" in columns:
        invalid |= detail[columns["account"]].isna() | detail[columns["account"]].astype("string").str.strip().eq("")
    detail["行级异常"] = np.where(invalid, "是", "否")
    voucher = detail.groupby(columns["voucher"], dropna=False).agg(借方合计=("借方金额_标准", "sum"), 贷方合计=("贷方金额_标准", "sum"), 分录行数=(columns["voucher"], "size")).reset_index()
    voucher["借贷差额"] = voucher["借方合计"] - voucher["贷方合计"]
    voucher["平衡状态"] = np.where(voucher["借贷差额"].abs() <= tolerance, "平衡", "不平衡")
    exceptions = detail.loc[invalid].copy()
    unbalanced = voucher.loc[voucher["平衡状态"] == "不平衡"].copy()
    return FinanceAnalysisResult("journal_audit", {"primary": voucher, "凭证明细审计": detail, "不平衡凭证": unbalanced, "异常分录": exceptions}, {"voucher_count": len(voucher), "unbalanced_vouchers": len(unbalanced), "invalid_rows": len(exceptions), "tolerance": tolerance})


def analyze_finance(
    frame: pd.DataFrame,
    *,
    task: str,
    columns: Mapping[str, str],
    as_of_date: str | None = None,
    buckets: list[int] | None = None,
    perspective: str = "income",
    tolerance: float = 0.01,
) -> FinanceAnalysisResult:
    params = {"task": task, "columns": columns, "perspective": perspective, "tolerance": tolerance}
    if as_of_date is not None:
        params["as_of_date"] = as_of_date
    if buckets is not None:
        params["buckets"] = buckets
    validate_finance_params(params)
    _require_source_columns(frame, columns)
    if task == "ar_aging":
        return _ar_aging(frame, columns, as_of_date=as_of_date, buckets=buckets or [30, 60, 90])
    if task == "budget_variance":
        return _budget_variance(frame, columns, perspective=perspective)
    if task == "cash_flow":
        return _cash_flow(frame, columns)
    if task == "financial_ratios":
        return _financial_ratios(frame, columns)
    if task == "journal_audit":
        return _journal_audit(frame, columns, tolerance=float(tolerance))
    raise ValueError(f"不支持的财务任务：{task}")


__all__ = ["FINANCE_TASKS", "FinanceAnalysisResult", "analyze_finance", "finance_column_names", "validate_finance_params"]
