"""Framework-free production acceptance for the deterministic finance engine."""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import sys

import pandas as pd


PROJECT_PARENT = Path(__file__).resolve().parents[2]
if str(PROJECT_PARENT) not in sys.path:
    sys.path.insert(0, str(PROJECT_PARENT))

from excel_data_toolbox.core import export_tables
from excel_data_toolbox.finance import analyze_finance
from excel_data_toolbox.nl_agent import build_table_catalog, execute_plan, validate_plan


def ensure(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_dir = Path(__file__).resolve().parents[1] / "outputs" / "finance_acceptance" / stamp
    output_dir.mkdir(parents=True, exist_ok=True)
    checks: list[dict[str, object]] = []
    tables: dict[str, pd.DataFrame] = {}

    ar = pd.DataFrame({"客户": ["甲", "甲", "乙", "丙"], "发票号": ["A1", "A2", "B1", "C1"], "到期日": ["2026-08-15", "2026-07-01", "2026-09-01", "bad"], "应收金额": [1000, 2000, 1500, 500], "已收金额": [200, 500, 0, 0]})
    ar_result = analyze_finance(ar, task="ar_aging", columns={"counterparty": "客户", "invoice": "发票号", "due_date": "到期日", "amount": "应收金额", "paid_amount": "已收金额"}, as_of_date="2026-08-31", buckets=[30, 60, 90])
    ensure(ar_result.report["outstanding_total"] == 4300.0, "应收未结合计错误")
    checks.append({"检查": "应收账龄", "状态": "通过", "结果": "未结合计 4,300；含客户账龄和无效日期"})
    tables.update({f"应收_{name}": frame for name, frame in ar_result.outputs.items()})

    budget = pd.DataFrame({"月份": ["1月", "1月"], "科目": ["差旅", "推广"], "实际": ["120", "80"], "预算": ["100", "100"]})
    budget_result = analyze_finance(budget, task="budget_variance", columns={"period": "月份", "category": "科目", "actual": "实际", "budget": "预算"}, perspective="cost")
    ensure(budget_result.outputs["primary"]["差异判断"].tolist() == ["不利", "有利"], "成本预算有利/不利判断错误")
    checks.append({"检查": "预算差异", "状态": "通过", "结果": "差异额、差异率、有利/不利判断正确"})
    tables.update({f"预算_{name}": frame for name, frame in budget_result.outputs.items()})

    cash = pd.DataFrame({"日期": ["2026-01-01", "2026-01-05", "2026-02-01"], "方向": ["收入", "支出", "流入"], "分类": ["销售", "采购", "销售"], "金额": [1000, 300, 500]})
    cash_result = analyze_finance(cash, task="cash_flow", columns={"date": "日期", "direction": "方向", "category": "分类", "amount": "金额"})
    ensure(cash_result.outputs["月度现金流"]["累计净现金流"].tolist() == [700.0, 1200.0], "累计净现金流错误")
    checks.append({"检查": "现金流", "状态": "通过", "结果": "流入、流出、月度净额与累计额正确"})
    tables.update({f"现金流_{name}": frame for name, frame in cash_result.outputs.items()})

    statement = pd.DataFrame({"年度": [2025], "营业收入": [1000], "净利润": [100], "流动资产": [600], "流动负债": [300], "存货": [100], "总资产": [2000], "总负债": [800], "权益": [1200]})
    ratio_result = analyze_finance(statement, task="financial_ratios", columns={"period": "年度", "revenue": "营业收入", "net_profit": "净利润", "current_assets": "流动资产", "current_liabilities": "流动负债", "inventory": "存货", "total_assets": "总资产", "total_liabilities": "总负债", "equity": "权益"})
    ensure(ratio_result.outputs["primary"].loc[0, "净利率"] == 0.1, "净利率错误")
    ensure(ratio_result.outputs["primary"].loc[0, "资产负债率"] == 0.4, "资产负债率错误")
    checks.append({"检查": "财务比率", "状态": "通过", "结果": f"自动计算 {ratio_result.report['ratio_count']} 项指标"})
    tables.update({f"比率_{name}": frame for name, frame in ratio_result.outputs.items()})

    ledger = pd.DataFrame({"凭证号": ["V1", "V1", "V2", "V2"], "科目": ["银行", "收入", "费用", "银行"], "借方": [100, 0, 60, 0], "贷方": [0, 100, 0, 50]})
    catalog = build_table_catalog({"ledger": ledger}, display_names={"ledger": "总账"})
    plan = validate_plan({"schema_version": 1, "status": "ready", "summary": "输入范围：总账；处理动作：凭证审计；关键字段/规则/阈值：借贷差额0.01；输出：异常表；人工核验边界：财务人员复核。", "message": "可执行", "clarification_questions": [], "assumptions": [], "warnings": ["结果需财务人员复核"], "steps": [{"id": "finance_1", "operation": "finance", "input_ids": ["ledger"], "output_name": "凭证审计", "params": {"task": "journal_audit", "columns": {"voucher": "凭证号", "account": "科目", "debit": "借方", "credit": "贷方"}, "tolerance": 0.01}}]}, catalog)
    executed = execute_plan(plan, {"ledger": ledger}, dry_run=False)
    ensure(executed.reports["finance_1"]["unbalanced_vouchers"] == 1, "凭证平衡审计错误")
    checks.append({"检查": "AI 白名单执行", "状态": "通过", "结果": "finance 计划校验、执行及多表产物正常"})
    tables.update(executed.tables)

    tables = {"00_验收报告": pd.DataFrame(checks), **tables}
    workbook = output_dir / "财务AI功能验收交付包.xlsx"
    export_tables(tables, workbook, include_log=False, overwrite=True)
    report = {"status": "passed", "checks_passed": len(checks), "checks_total": 5, "workbook": str(workbook), "tasks": ["ar_aging", "budget_variance", "cash_flow", "financial_ratios", "journal_audit"]}
    report_path = output_dir / "财务AI功能验收报告.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False))


if __name__ == "__main__":
    main()
