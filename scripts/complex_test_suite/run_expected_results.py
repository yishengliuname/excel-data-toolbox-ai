from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pandas as pd

from excel_data_toolbox.core import export_tables, load_tables
from excel_data_toolbox.nl_agent import (
    AgentPlan,
    build_table_catalog,
    execute_plan,
    preview_plan,
    validate_plan,
)


PROJECT_DIR = Path(__file__).resolve().parents[2]
OUTPUT_DIR = PROJECT_DIR / "outputs" / "complex_test_suite_20260822"


def _sheet(raw: dict[str, pd.DataFrame], suffix: str) -> pd.DataFrame:
    for key, frame in raw.items():
        if key.endswith(f"::{suffix}"):
            return frame.copy(deep=True)
    raise KeyError(f"找不到工作表：{suffix}")


def _validation_rule(
    rule_id: str,
    rule_type: str,
    column: str,
    params: dict[str, Any],
    message: str,
    severity: str = "error",
) -> dict[str, Any]:
    return {
        "rule_id": rule_id,
        "rule_type": rule_type,
        "column": column,
        "severity": severity,
        "params": params,
        "message": message,
        "enabled": True,
    }


def inventory_plan() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：华东销售出库、华南销售出库、采购入库、月初库存、月末盘点与 SKU 主数据；"
            "处理动作：销售追加清洗、质量验收、主数据补全、汇总、趋势、透视、贡献与异常检测，采购清洗验收与汇总，库存新旧比对；"
            "关键字段/规则/阈值：出库单号唯一，销售数量 -20 至 500，采购数量 1 至 500，销售金额使用 1.5 倍 IQR，库存按门店编码+SKU 比对；"
            "输出：质量失败、经营分析、采购汇总、库存新增/删除/修改/未变化及异常明细；"
            "人工核验边界：缺关键键、负数采购、新品未建档、重复库存键和异常大单均不自动修正。"
        ),
        "message": "计划仅调用本机白名单操作，所有结果另存新表。",
        "clarification_questions": [],
        "assumptions": [],
        "warnings": ["退货负数量不等同于错误", "库存变化只做识别，不推断责任"],
        "steps": [
            {
                "id": "sales_concat",
                "operation": "concat",
                "input_ids": ["sales_east", "sales_south"],
                "output_name": "01_两区销售合并",
                "params": {"join": "outer", "ignore_index": True, "source_column": "来源表序号"},
            },
            {
                "id": "sales_clean",
                "operation": "clean",
                "input_ids": ["$sales_concat"],
                "output_name": "02_销售清洗结果",
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
                "id": "sales_validate",
                "operation": "validate",
                "input_ids": ["$sales_clean"],
                "output_name": "03_销售质量验收",
                "params": {
                    "include_values": False,
                    "rules": [
                        _validation_rule("S01", "not_null", "出库单号", {"blank_as_null": True}, "出库单号不能为空"),
                        _validation_rule("S02", "unique", "出库单号", {"ignore_nulls": True, "blank_as_null": True}, "出库单号必须唯一"),
                        _validation_rule("S03", "not_null", "门店编码", {"blank_as_null": True}, "门店编码不能为空"),
                        _validation_rule("S04", "not_null", "SKU", {"blank_as_null": True}, "SKU 不能为空"),
                        _validation_rule("S05", "range", "销售数量", {"min": -20, "max": 500, "value_type": "numeric", "ignore_nulls": False}, "销售数量超出授权范围"),
                        _validation_rule("S06", "range", "销售金额", {"min": -100000, "max": 1000000, "value_type": "numeric", "ignore_nulls": False}, "销售金额超出授权范围"),
                    ],
                },
            },
            {
                "id": "sales_lookup",
                "operation": "lookup",
                "input_ids": ["$sales_clean", "sku_master"],
                "output_name": "04_销售SKU补全",
                "params": {
                    "source_key": "SKU",
                    "lookup_key": "SKU",
                    "value_columns": ["商品标准名称", "品类", "供应商", "标准进价", "安全库存", "是否在售"],
                    "keep_lookup_duplicate": False,
                    "add_match_column": True,
                    "match_column": "SKU匹配状态",
                },
            },
            {
                "id": "sales_summary",
                "operation": "summary",
                "input_ids": ["$sales_lookup"],
                "output_name": "05_门店SKU销售汇总",
                "params": {"by": ["门店编码", "SKU"], "aggregations": {"销售数量": "sum", "销售金额": "sum", "出库单号": "nunique"}, "dropna": False, "sort": True},
            },
            {
                "id": "sales_trend",
                "operation": "trend",
                "input_ids": ["$sales_lookup"],
                "output_name": "06_区域月度销售趋势",
                "params": {"date_column": "销售日期", "value_columns": ["销售数量", "销售金额"], "frequency": "month", "aggregation": {"销售数量": "sum", "销售金额": "sum"}, "group_by": ["区域"], "period_column": "月份"},
            },
            {
                "id": "sales_pivot",
                "operation": "pivot",
                "input_ids": ["$sales_lookup"],
                "output_name": "07_品类门店销售透视",
                "params": {"index": ["品类"], "columns": ["门店名称"], "values": "销售金额", "aggregation": "sum", "fill_value": 0, "margins": True, "margins_name": "合计"},
            },
            {
                "id": "sales_contribution",
                "operation": "contribution",
                "input_ids": ["$sales_lookup"],
                "output_name": "08_SKU销售贡献",
                "params": {"category_columns": ["品类", "SKU"], "value_column": "销售金额", "aggregation": "sum", "pareto_threshold": 0.8, "top_n": 36, "include_other": False},
            },
            {
                "id": "sales_outliers",
                "operation": "outliers",
                "input_ids": ["$sales_clean"],
                "output_name": "09_销售金额异常",
                "params": {"columns": ["销售金额"], "method": "iqr", "iqr_multiplier": 1.5},
            },
            {
                "id": "receipt_clean",
                "operation": "clean",
                "input_ids": ["receipts"],
                "output_name": "10_采购入库清洗",
                "params": {"trim_whitespace": True, "normalize_blank_strings": True, "drop_empty_rows": True, "drop_empty_columns": True, "drop_duplicates": True, "keep_duplicate": "first", "infer_types": False, "missing_strategy": "keep", "reset_index": True},
            },
            {
                "id": "receipt_validate",
                "operation": "validate",
                "input_ids": ["$receipt_clean"],
                "output_name": "11_采购质量验收",
                "params": {
                    "include_values": False,
                    "rules": [
                        _validation_rule("P01", "not_null", "入库单号", {"blank_as_null": True}, "入库单号不能为空"),
                        _validation_rule("P02", "unique", "入库单号", {"ignore_nulls": True, "blank_as_null": True}, "入库单号必须唯一"),
                        _validation_rule("P03", "not_null", "SKU", {"blank_as_null": True}, "SKU 不能为空"),
                        _validation_rule("P04", "range", "入库数量", {"min": 1, "max": 500, "value_type": "numeric", "ignore_nulls": False}, "入库数量必须在 1 到 500"),
                        _validation_rule("P05", "range", "含税金额", {"min": 0.01, "max": 1000000, "value_type": "numeric", "ignore_nulls": False}, "采购金额必须为正且在授权范围"),
                    ],
                },
            },
            {
                "id": "receipt_summary",
                "operation": "summary",
                "input_ids": ["$receipt_clean"],
                "output_name": "12_门店SKU采购汇总",
                "params": {"by": ["门店编码", "SKU"], "aggregations": {"入库数量": "sum", "含税金额": "sum", "入库单号": "nunique"}, "dropna": False, "sort": True},
            },
            {
                "id": "supplier_cluster",
                "operation": "fuzzy_cluster",
                "input_ids": ["$receipt_clean"],
                "output_name": "13_供应商名称聚类",
                "params": {"column": "供应商原值", "threshold": 0.88, "max_unique": 200},
            },
            {
                "id": "closing_clean",
                "operation": "clean",
                "input_ids": ["closing_inventory"],
                "output_name": "14_月末库存清洗",
                "params": {"trim_whitespace": True, "normalize_blank_strings": True, "drop_empty_rows": True, "drop_empty_columns": True, "drop_duplicates": True, "duplicate_subset": ["门店编码", "SKU"], "keep_duplicate": "first", "infer_types": False, "missing_strategy": "keep", "reset_index": True},
            },
            {
                "id": "inventory_compare",
                "operation": "compare",
                "input_ids": ["opening_inventory", "$closing_clean"],
                "output_name": "15_月初月末库存比对",
                "params": {"key_columns": ["门店编码", "SKU"], "compare_columns": ["库存数量", "库存金额"], "suffixes": ["_月初", "_月末"], "include_unchanged": True},
            },
            {
                "id": "closing_outliers",
                "operation": "outliers",
                "input_ids": ["$closing_clean"],
                "output_name": "16_月末库存异常",
                "params": {"columns": ["库存数量", "库存金额"], "method": "iqr", "iqr_multiplier": 1.5},
            },
        ],
    }


def hr_plan() -> dict[str, Any]:
    return {
        "schema_version": 1,
        "status": "ready",
        "summary": (
            "输入范围：员工主数据、2026年7月考勤、2026年6月薪资、2026年7月薪资和部门预算；"
            "处理动作：考勤与薪资清洗验收、员工级汇总、主数据补全、部门汇总与预算匹配、跨月比对、异常/相关性分析及隐私脱敏；"
            "关键字段/规则/阈值：员工编号为主键，实际出勤 0 至 16 小时，实发工资 0 至 10 万，实发工资使用 1.5 倍 IQR；"
            "输出：质量失败、员工考勤汇总、部门预算对照、薪资新增/删除/修改/未变化、异常与脱敏交付表；"
            "人工核验边界：缺员工号、超长工时、重复薪资版本、负工资和异常高薪均不自动修正。"
        ),
        "message": "计划仅调用本机白名单操作，敏感明细只输出脱敏副本。",
        "clarification_questions": [],
        "assumptions": [],
        "warnings": ["工资差异不等同于计算错误", "所有身份及账户数据均为虚构"],
        "steps": [
            {
                "id": "attendance_clean",
                "operation": "clean",
                "input_ids": ["attendance"],
                "output_name": "01_考勤清洗结果",
                "params": {"trim_whitespace": True, "normalize_blank_strings": True, "drop_empty_rows": True, "drop_empty_columns": True, "drop_duplicates": True, "keep_duplicate": "first", "infer_types": False, "missing_strategy": "keep", "reset_index": True},
            },
            {
                "id": "attendance_validate",
                "operation": "validate",
                "input_ids": ["$attendance_clean"],
                "output_name": "02_考勤质量验收",
                "params": {
                    "include_values": False,
                    "rules": [
                        _validation_rule("A01", "not_null", "员工编号", {"blank_as_null": True}, "员工编号不能为空"),
                        _validation_rule("A02", "range", "实际出勤小时", {"min": 0, "max": 16, "value_type": "numeric", "ignore_nulls": False}, "实际出勤超出范围"),
                        _validation_rule("A03", "range", "加班小时", {"min": 0, "max": 12, "value_type": "numeric", "ignore_nulls": False}, "加班小时超出范围"),
                        _validation_rule("A04", "range", "迟到分钟", {"min": 0, "max": 360, "value_type": "numeric", "ignore_nulls": False}, "迟到分钟超出范围"),
                        _validation_rule("A05", "range", "缺勤小时", {"min": 0, "max": 8, "value_type": "numeric", "ignore_nulls": False}, "缺勤小时超出范围"),
                    ],
                },
            },
            {
                "id": "attendance_summary",
                "operation": "summary",
                "input_ids": ["$attendance_clean"],
                "output_name": "03_员工考勤汇总",
                "params": {"by": ["员工编号"], "aggregations": {"考勤日期": "count", "实际出勤小时": "sum", "加班小时": "sum", "迟到分钟": "sum", "缺勤小时": "sum"}, "dropna": False, "sort": True},
            },
            {
                "id": "attendance_lookup",
                "operation": "lookup",
                "input_ids": ["$attendance_summary", "employee_master"],
                "output_name": "04_考勤员工信息补全",
                "params": {"source_key": "员工编号", "lookup_key": "员工编号", "value_columns": ["姓名", "部门", "工作城市", "在职状态"], "keep_lookup_duplicate": False, "add_match_column": True, "match_column": "员工匹配状态"},
            },
            {
                "id": "payroll_clean",
                "operation": "clean",
                "input_ids": ["payroll_july"],
                "output_name": "05_七月薪资清洗",
                "params": {"trim_whitespace": True, "normalize_blank_strings": True, "drop_empty_rows": True, "drop_empty_columns": True, "drop_duplicates": True, "duplicate_subset": ["员工编号"], "keep_duplicate": "last", "infer_types": False, "missing_strategy": "keep", "reset_index": True},
            },
            {
                "id": "payroll_validate",
                "operation": "validate",
                "input_ids": ["$payroll_clean"],
                "output_name": "06_七月薪资质量验收",
                "params": {
                    "include_values": False,
                    "rules": [
                        _validation_rule("W01", "not_null", "员工编号", {"blank_as_null": True}, "员工编号不能为空"),
                        _validation_rule("W02", "unique", "员工编号", {"ignore_nulls": True, "blank_as_null": True}, "员工编号必须唯一"),
                        _validation_rule("W03", "range", "实发工资", {"min": 0, "max": 100000, "value_type": "numeric", "ignore_nulls": False}, "实发工资超出范围"),
                    ],
                },
            },
            {
                "id": "payroll_lookup",
                "operation": "lookup",
                "input_ids": ["$payroll_clean", "employee_master"],
                "output_name": "07_七月薪资员工补全",
                "params": {"source_key": "员工编号", "lookup_key": "员工编号", "value_columns": ["部门", "工作城市", "身份证号", "手机号", "邮箱", "在职状态"], "keep_lookup_duplicate": False, "add_match_column": True, "match_column": "员工匹配状态"},
            },
            {
                "id": "department_summary",
                "operation": "summary",
                "input_ids": ["$payroll_lookup"],
                "output_name": "08_部门薪资汇总",
                "params": {"by": ["部门"], "aggregations": {"实发工资": "sum", "加班工资": "sum", "员工编号": "nunique"}, "dropna": False, "sort": True},
            },
            {
                "id": "budget_join",
                "operation": "join",
                "input_ids": ["$department_summary", "payroll_budget"],
                "output_name": "09_部门薪资预算对照",
                "params": {"on": "部门", "how": "left", "suffixes": ["_实际", "_预算"], "validate": "one_to_one"},
            },
            {
                "id": "payroll_compare",
                "operation": "compare",
                "input_ids": ["payroll_june", "$payroll_clean"],
                "output_name": "10_六月七月薪资比对",
                "params": {"key_columns": ["员工编号"], "compare_columns": ["基本工资", "岗位津贴", "实发工资"], "suffixes": ["_六月", "_七月"], "include_unchanged": True},
            },
            {
                "id": "payroll_outliers",
                "operation": "outliers",
                "input_ids": ["$payroll_clean"],
                "output_name": "11_七月实发工资异常",
                "params": {"columns": ["实发工资"], "method": "iqr", "iqr_multiplier": 1.5},
            },
            {
                "id": "payroll_correlation",
                "operation": "correlation",
                "input_ids": ["$payroll_clean"],
                "output_name": "12_薪资字段相关性",
                "params": {"columns": ["基本工资", "岗位津贴", "加班工资", "考勤扣款", "绩效奖金", "社保公积金", "个税", "实发工资"], "method": "pearson", "min_periods": 10},
            },
            {
                "id": "payroll_mask",
                "operation": "mask",
                "input_ids": ["$payroll_lookup"],
                "output_name": "13_七月薪资脱敏交付",
                "params": {"columns": {"姓名": "name", "身份证号": "id", "手机号": "phone", "邮箱": "email", "银行卡号": "id"}},
            },
            {
                "id": "department_split",
                "operation": "split",
                "input_ids": ["$payroll_lookup"],
                "output_name": "14_部门薪资拆分",
                "params": {"by": ["部门"], "drop_group_columns": False},
            },
        ],
    }


def _summary_frame(case_name: str, plan: AgentPlan, result: Any) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"项目": "案例", "结果": case_name},
            {"项目": "计划状态", "结果": plan.status},
            {"项目": "标准话术", "结果": plan.summary},
            {"项目": "步骤数", "结果": len(plan.steps)},
            {"项目": "结果表数", "结果": len(result.tables)},
            {"项目": "执行方式", "结果": "本机白名单执行器（未调用真实 DeepSeek API）"},
            {"项目": "数据性质", "结果": "全部为程序生成的虚构测试数据"},
        ]
    )


def _report_frame(plan: AgentPlan, result: Any) -> pd.DataFrame:
    rows = []
    for step in plan.steps:
        report = result.reports.get(step.id, {})
        rows.append(
            {
                "步骤": step.id,
                "操作": step.operation,
                "输出名称": step.output_name,
                "输入": "；".join(step.input_ids),
                "状态": "通过",
                "结果摘要(JSON)": json.dumps(report, ensure_ascii=False, sort_keys=True, default=str),
            }
        )
    return pd.DataFrame(rows)


def run_case(
    *,
    case_name: str,
    input_filename: str,
    output_filename: str,
    table_mapping: dict[str, str],
    plan_payload: dict[str, Any],
) -> dict[str, Any]:
    raw = load_tables(OUTPUT_DIR / input_filename)
    tables = {table_id: _sheet(raw, sheet_name) for table_id, sheet_name in table_mapping.items()}
    catalog = build_table_catalog(tables, display_names=table_mapping)
    plan = validate_plan(plan_payload, catalog)
    if not plan.executable:
        raise RuntimeError(f"{case_name} 计划不可执行：{plan.to_dict()}")
    preview = preview_plan(plan, tables)
    dry_run = execute_plan(plan, tables, dry_run=True)
    result = execute_plan(plan, tables, dry_run=False)
    if list(dry_run.tables) != list(result.tables):
        raise RuntimeError(f"{case_name} 预演与正式执行表结构不一致")
    payload = {
        "00_执行摘要": _summary_frame(case_name, plan, result),
        "00_步骤执行报告": _report_frame(plan, result),
        **dict(result.tables),
    }
    output_path = OUTPUT_DIR / output_filename
    export_tables(payload, output_path, include_log=False, overwrite=True)
    report = {
        "case": case_name,
        "network_call": False,
        "input": str(OUTPUT_DIR / input_filename),
        "output": str(output_path),
        "plan": plan.to_dict(),
        "preview": preview.to_dict(),
        "execution": result.to_dict(),
        "dry_run_table_names_match": True,
    }
    report_path = OUTPUT_DIR / f"{output_path.stem}_report.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return {
        "case": case_name,
        "input_rows": sum(len(frame) for frame in tables.values()),
        "steps": len(plan.steps),
        "result_tables": len(result.tables),
        "output": str(output_path),
        "report": str(report_path),
        "reports": result.reports,
    }


def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    inventory = run_case(
        case_name="连锁门店库存、采购与销售异常诊断",
        input_filename="案例02_连锁库存销售异常_输入.xlsx",
        output_filename="案例02_连锁库存销售异常_标准结果.xlsx",
        table_mapping={
            "sales_east": "华东销售出库",
            "sales_south": "华南销售出库",
            "receipts": "采购入库",
            "opening_inventory": "月初库存",
            "closing_inventory": "月末盘点",
            "sku_master": "SKU主数据",
        },
        plan_payload=inventory_plan(),
    )
    hr = run_case(
        case_name="考勤、薪资、预算核验与隐私脱敏",
        input_filename="案例03_考勤薪资核验脱敏_输入.xlsx",
        output_filename="案例03_考勤薪资核验脱敏_标准结果.xlsx",
        table_mapping={
            "employee_master": "员工主数据",
            "attendance": "2026-07考勤明细",
            "payroll_june": "2026-06薪资",
            "payroll_july": "2026-07薪资",
            "payroll_budget": "部门薪资预算",
        },
        plan_payload=hr_plan(),
    )
    summary = {"inventory": inventory, "hr": hr}
    (OUTPUT_DIR / "expected_results_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, default=str), encoding="utf-8"
    )
    print(json.dumps({key: {field: value for field, value in data.items() if field != "reports"} for key, data in summary.items()}, ensure_ascii=False))


if __name__ == "__main__":
    main()
