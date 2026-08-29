"""Run the maximum-difficulty V4 acceptance case without using a real API key.

This script exercises the exact post-model boundary: strict JSON validation,
plan preview, isolated dry-run, controlled execution, and workbook export.  It
never contacts DeepSeek; the browser UI is the secure place for a live call.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

from excel_data_toolbox.core import export_tables, load_tables
from excel_data_toolbox.nl_agent import (
    build_table_catalog,
    execute_plan,
    preview_plan,
    validate_plan,
)


PROJECT_DIR = Path(__file__).resolve().parents[1]
CASE_DIR = PROJECT_DIR / "outputs" / "deepseek_v4_production_case"
INPUT_PATH = CASE_DIR / "华辰商贸_2026H1订单回款经营诊断_高难度案例.xlsx"
OUTPUT_PATH = CASE_DIR / "华辰商贸_2026H1_AI自动执行结果.xlsx"
REPORT_PATH = CASE_DIR / "case_execution_report.json"


def _sheet(tables: dict[str, pd.DataFrame], name: str) -> pd.DataFrame:
    matches = [frame for key, frame in tables.items() if key.endswith(f"::{name}")]
    if len(matches) != 1:
        raise RuntimeError(f"案例工作簿缺少唯一工作表：{name}")
    return matches[0]


def _rule(
    rule_id: str,
    rule_type: str,
    column: str,
    params: dict[str, Any],
    message: str,
) -> dict[str, Any]:
    return {
        "rule_id": rule_id,
        "rule_type": rule_type,
        "column": column,
        "severity": "error",
        "params": params,
        "message": message,
        "enabled": True,
    }


def build_max_difficulty_plan() -> dict[str, Any]:
    """Representative strict JSON that a successful DeepSeek call should return."""

    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：订单、回款、客户主数据、商品主数据与区域月度目标表；"
            "处理动作：订单清洗与质量验收、客户名称模糊匹配、订单回款容差对账、经营分析、目标匹配及隐私脱敏；"
            "关键字段/规则/阈值：客户名称相似度 88%，金额容差 0.05 元，到账日期容差 7 天，订单金额使用 IQR 识别异常；"
            "输出：清洗结果、验收明细、匹配与对账分表、趋势/透视/贡献/RFM 分析、目标匹配及脱敏交付表；"
            "人工核验边界：重复键、拆分回款、模糊歧义、待认领款、质量失败和容差外差异均不自动确认。"
        ),
        "message": "计划只调用本机白名单能力；差异、歧义和异常仅标记并进入人工复核。",
        "clarification_questions": [],
        "assumptions": [
            "订单金额为订单侧应收口径，回款金额为银行到账口径",
            "金额容差为 0.05 元，到账日期允许晚 7 天",
            "区域月份字段是月度区域目标的一对一关联键",
        ],
        "warnings": [
            "模糊名称、拆分回款、重复流水和待认领款不得静默自动确认",
            "原始手机号和邮箱只出现在输入与本机内存，交付使用脱敏副本",
        ],
        "steps": [
            {
                "id": "clean_orders",
                "operation": "clean",
                "input_ids": ["orders"],
                "output_name": "01_订单清洗结果",
                "params": {
                    "trim_whitespace": True,
                    "normalize_blank_strings": True,
                    "drop_empty_rows": True,
                    "drop_empty_columns": True,
                    "drop_duplicates": True,
                    "keep_duplicate": "first",
                    "infer_types": False,
                    "missing_strategy": "keep",
                    "reset_index": True,
                },
            },
            {
                "id": "validate_orders",
                "operation": "validate",
                "input_ids": ["$clean_orders"],
                "output_name": "02_订单质量验收",
                "params": {
                    "rules": [
                        _rule("ORD-001", "not_null", "订单号", {"blank_as_null": True}, "订单号不能为空"),
                        _rule("ORD-002", "unique", "订单号", {"ignore_nulls": True, "blank_as_null": True}, "订单号必须唯一"),
                        _rule("CUS-001", "not_null", "客户ID", {"blank_as_null": True}, "客户 ID 不能为空"),
                        _rule("REG-001", "not_null", "区域", {"blank_as_null": True}, "区域不能为空"),
                        _rule("QTY-001", "range", "数量", {"min": 1, "max": 100, "value_type": "numeric", "ignore_nulls": True}, "数量必须在 1 到 100 之间"),
                        _rule("DSC-001", "range", "折扣率", {"min": 0, "max": 0.3, "value_type": "numeric", "ignore_nulls": True}, "折扣率必须在 0 到 30% 之间"),
                        _rule("AMT-001", "range", "订单金额", {"min": -500000, "max": 500000, "value_type": "numeric", "ignore_nulls": False}, "订单金额超出授权业务区间"),
                    ],
                    "include_values": False,
                    "max_value_chars": 80,
                },
            },
            {
                "id": "customer_match",
                "operation": "fuzzy_lookup",
                "input_ids": ["$clean_orders", "customers"],
                "output_name": "03_客户名称标准化建议",
                "params": {
                    "source_key": "客户名称原值",
                    "lookup_key": "客户标准名称",
                    "value_columns": ["客户标准名称", "客户层级", "行业"],
                    "threshold": 0.88,
                    "ambiguous_gap": 0.03,
                },
            },
            {
                "id": "reconcile_cash",
                "operation": "reconcile",
                "input_ids": ["$clean_orders", "payments"],
                "output_name": "04_订单回款高级对账",
                "params": {
                    "left_amount": "订单金额",
                    "right_amount": "回款金额",
                    "left_date": "下单日期",
                    "right_date": "到账日期",
                    "left_key_columns": ["订单号"],
                    "right_key_columns": ["订单号"],
                    "left_secondary_columns": ["客户名称原值"],
                    "right_secondary_columns": ["付款方原值"],
                    "amount_tolerance": 0.05,
                    "date_tolerance_days": 7,
                    "enable_split_candidates": True,
                    "max_candidates_per_row": 30,
                    "max_candidate_pairs": 150000,
                    "max_split_combinations": 30000,
                },
            },
            {
                "id": "monthly_summary",
                "operation": "summary",
                "input_ids": ["$clean_orders"],
                "output_name": "05_区域月份经营汇总",
                "params": {
                    "by": ["区域月份"],
                    "aggregations": {"订单金额": "sum", "订单号": "nunique", "数量": "sum"},
                    "dropna": False,
                    "sort": True,
                },
            },
            {
                "id": "target_match",
                "operation": "join",
                "input_ids": ["$monthly_summary", "targets"],
                "output_name": "06_区域月份实际对目标",
                "params": {"on": "区域月份", "how": "left", "suffixes": ["_实际", "_目标"], "validate": "one_to_one"},
            },
            {
                "id": "sales_trend",
                "operation": "trend",
                "input_ids": ["$clean_orders"],
                "output_name": "07_月度区域销售趋势",
                "params": {
                    "date_column": "下单日期",
                    "value_columns": ["订单金额", "数量"],
                    "frequency": "month",
                    "aggregation": {"订单金额": "sum", "数量": "sum"},
                    "group_by": ["区域"],
                    "period_column": "月份",
                },
            },
            {
                "id": "channel_pivot",
                "operation": "pivot",
                "input_ids": ["$clean_orders"],
                "output_name": "08_区域渠道销售透视",
                "params": {"index": ["区域"], "columns": ["渠道"], "values": "订单金额", "aggregation": "sum", "fill_value": 0, "margins": True, "margins_name": "合计"},
            },
            {
                "id": "amount_outliers",
                "operation": "outliers",
                "input_ids": ["$clean_orders"],
                "output_name": "09_订单金额异常",
                "params": {"columns": ["订单金额"], "method": "iqr", "iqr_multiplier": 1.5},
            },
            {
                "id": "category_contribution",
                "operation": "contribution",
                "input_ids": ["$clean_orders"],
                "output_name": "10_商品销售贡献",
                "params": {"category_columns": ["SKU", "商品名称原值"], "value_column": "订单金额", "aggregation": "sum", "pareto_threshold": 0.8, "top_n": 24, "include_other": False},
            },
            {
                "id": "customer_rfm",
                "operation": "rfm",
                "input_ids": ["$clean_orders"],
                "output_name": "11_RFM客户价值分群",
                "params": {"customer_column": "客户ID", "date_column": "下单日期", "amount_column": "订单金额", "transaction_column": "订单号", "reference_date": "2026-07-01", "quantiles": 5},
            },
            {
                "id": "privacy_copy",
                "operation": "mask",
                "input_ids": ["$clean_orders"],
                "output_name": "12_订单交付脱敏版",
                "params": {"columns": {"客户名称原值": "name", "手机号": "phone", "邮箱": "email"}},
            },
        ],
    }


def main() -> None:
    loaded = load_tables(INPUT_PATH)
    tables = {
        "orders": _sheet(loaded, "订单明细"),
        "payments": _sheet(loaded, "回款流水"),
        "customers": _sheet(loaded, "客户主数据"),
        "products": _sheet(loaded, "商品主数据"),
        "targets": _sheet(loaded, "区域月度目标"),
    }
    display_names = {
        "orders": "订单明细",
        "payments": "回款流水",
        "customers": "客户主数据",
        "products": "商品主数据",
        "targets": "区域月度目标",
    }
    catalog = build_table_catalog(tables, display_names=display_names)
    plan = validate_plan(build_max_difficulty_plan(), catalog)
    if not plan.executable:
        raise RuntimeError(f"案例计划不可执行：{plan.to_dict()}")

    preview = preview_plan(plan, tables)
    dry_run = execute_plan(plan, tables, dry_run=True)
    executed = execute_plan(plan, tables, dry_run=False)

    clean_report = executed.reports["clean_orders"]
    reconciliation_report = executed.reports["reconcile_cash"]
    validation_report = executed.reports["validate_orders"]
    checks = [
        {"检查": "计划状态", "期望": "ready", "实际": plan.status, "通过": plan.status == "ready"},
        {"检查": "执行步骤数", "期望": 12, "实际": len(plan.steps), "通过": len(plan.steps) == 12},
        {"检查": "清洗前订单行数", "期望": 968, "实际": clean_report["rows_before"], "通过": clean_report["rows_before"] == 968},
        {"检查": "清洗后订单行数", "期望": 960, "实际": clean_report["rows_after"], "通过": clean_report["rows_after"] == 960},
        {"检查": "删除完全重复行", "期望": 8, "实际": clean_report["duplicate_rows_removed"], "通过": clean_report["duplicate_rows_removed"] == 8},
        {"检查": "质量规则执行", "期望": 7, "实际": validation_report["rule_count"], "通过": validation_report["rule_count"] == 7},
        {"检查": "对账产生摘要", "期望": "非空", "实际": len(reconciliation_report), "通过": bool(reconciliation_report)},
        {"检查": "生成结果表", "期望": ">=20", "实际": len(executed.tables), "通过": len(executed.tables) >= 20},
        {"检查": "干跑与正式输出表结构一致", "期望": True, "实际": list(dry_run.tables) == list(executed.tables), "通过": list(dry_run.tables) == list(executed.tables)},
    ]
    if not all(bool(item["通过"]) for item in checks):
        raise RuntimeError(f"案例验收失败：{checks}")

    report_rows = []
    for step in plan.steps:
        report_rows.append(
            {
                "步骤": step.id,
                "操作": step.operation,
                "输出名称": step.output_name,
                "输入": "；".join(step.input_ids),
                "状态": "通过",
                "摘要(JSON)": json.dumps(executed.to_dict()["reports"].get(step.id, {}), ensure_ascii=False, sort_keys=True),
            }
        )
    export_payload = {
        "00_案例执行报告": pd.DataFrame(report_rows),
        "00_自动验收结果": pd.DataFrame(checks),
        **dict(executed.tables),
    }
    export_tables(export_payload, OUTPUT_PATH, include_log=False, overwrite=True)

    report = {
        "case": "华辰商贸 2026H1 订单—回款经营诊断",
        "network_call": False,
        "reason": "安全验收脚本覆盖模型返回 JSON 后的完整本地链路；真实模型调用由本地网页密码框触发",
        "input_workbook": str(INPUT_PATH),
        "output_workbook": str(OUTPUT_PATH),
        "plan": plan.to_dict(),
        "preview": preview.to_dict(),
        "execution": executed.to_dict(),
        "checks": checks,
        "all_checks_passed": True,
    }
    REPORT_PATH.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "status": "passed",
                "steps": len(plan.steps),
                "output_tables": len(executed.tables),
                "input_rows": sum(len(frame) for frame in tables.values()),
                "output": str(OUTPUT_PATH),
                "report": str(REPORT_PATH),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
