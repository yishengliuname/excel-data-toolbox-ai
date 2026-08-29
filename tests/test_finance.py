from __future__ import annotations

import pandas as pd

from excel_data_toolbox.finance import analyze_finance
from excel_data_toolbox.nl_agent import build_table_catalog, execute_plan, validate_plan


def test_ar_aging_generates_detail_bucket_and_customer_outputs() -> None:
    frame = pd.DataFrame(
        {
            "客户": ["甲", "甲", "乙", "丙"],
            "发票号": ["A1", "A2", "B1", "C1"],
            "到期日": ["2026-08-15", "2026-07-01", "2026-09-01", "bad"],
            "应收金额": [1000, 2000, 1500, 500],
            "已收金额": [200, 500, 0, 0],
        }
    )
    result = analyze_finance(
        frame,
        task="ar_aging",
        columns={"counterparty": "客户", "invoice": "发票号", "due_date": "到期日", "amount": "应收金额", "paid_amount": "已收金额"},
        as_of_date="2026-08-31",
        buckets=[30, 60, 90],
    )
    assert set(result.outputs) == {"primary", "账龄汇总", "客户账龄"}
    assert result.report["outstanding_total"] == 4300.0
    assert result.report["invalid_date_count"] == 1
    # 2026-07-01 到 2026-08-31 相差 61 天，应进入 61-90 天区间。
    assert set(result.outputs["primary"]["账龄区间"]) == {"1-30天", "61-90天", "未到期", "日期无效"}


def test_budget_variance_marks_cost_overspend_unfavourable() -> None:
    frame = pd.DataFrame({"月份": ["1月", "1月"], "科目": ["差旅", "推广"], "实际": ["120", "80"], "预算": ["100", "100"]})
    result = analyze_finance(frame, task="budget_variance", columns={"period": "月份", "category": "科目", "actual": "实际", "budget": "预算"}, perspective="cost")
    assert result.outputs["primary"]["差异判断"].tolist() == ["不利", "有利"]
    assert result.outputs["期间汇总"].loc[0, "差异额"] == 0


def test_cash_flow_monthly_and_category_totals() -> None:
    frame = pd.DataFrame({"日期": ["2026-01-01", "2026-01-05", "2026-02-01"], "方向": ["收入", "支出", "流入"], "分类": ["销售", "采购", "销售"], "金额": [1000, 300, 500]})
    result = analyze_finance(frame, task="cash_flow", columns={"date": "日期", "direction": "方向", "category": "分类", "amount": "金额"})
    monthly = result.outputs["月度现金流"]
    assert monthly["净现金流"].tolist() == [700.0, 500.0]
    assert monthly["累计净现金流"].tolist() == [700.0, 1200.0]
    assert result.report["total_outflow"] == 300.0


def test_financial_ratios_compute_supported_metrics_without_inventing_values() -> None:
    frame = pd.DataFrame({"年度": [2025], "营业收入": [1000], "净利润": [100], "流动资产": [600], "流动负债": [300], "存货": [100], "总资产": [2000], "总负债": [800], "权益": [1200]})
    result = analyze_finance(frame, task="financial_ratios", columns={"period": "年度", "revenue": "营业收入", "net_profit": "净利润", "current_assets": "流动资产", "current_liabilities": "流动负债", "inventory": "存货", "total_assets": "总资产", "total_liabilities": "总负债", "equity": "权益"})
    ratios = result.outputs["primary"]
    assert ratios.loc[0, "净利率"] == 0.1
    assert ratios.loc[0, "流动比率"] == 2.0
    assert ratios.loc[0, "速动比率"] == 5 / 3
    assert ratios.loc[0, "资产负债率"] == 0.4


def test_journal_audit_and_ai_allowlist_execution() -> None:
    frame = pd.DataFrame({"凭证号": ["V1", "V1", "V2", "V2"], "科目": ["银行", "收入", "费用", "银行"], "借方": [100, 0, 60, 0], "贷方": [0, 100, 0, 50]})
    catalog = build_table_catalog({"ledger": frame}, display_names={"ledger": "总账"})
    plan = validate_plan(
        {"schema_version": 1, "status": "ready", "summary": "财务凭证审计", "message": "可执行", "clarification_questions": [], "assumptions": [], "warnings": ["结果需财务人员复核"], "steps": [{"id": "finance_1", "operation": "finance", "input_ids": ["ledger"], "output_name": "凭证审计", "params": {"task": "journal_audit", "columns": {"voucher": "凭证号", "account": "科目", "debit": "借方", "credit": "贷方"}, "tolerance": 0.01}}]},
        catalog,
    )
    executed = execute_plan(plan, {"ledger": frame}, dry_run=False)
    assert "凭证审计_不平衡凭证" in executed.tables
    assert executed.reports["finance_1"]["unbalanced_vouchers"] == 1
