# Showcase prompts

These are copyable, synthetic-data prompts for screenshots, demos, and issue discussions. They are deliberately phrased as business requests rather than implementation instructions.

## Sales and profit quality

> Analyze the selected sales workbooks together. First show the source sheets, row counts, duplicate keys, missing fields, and any conflicting definitions. Then calculate revenue, cost, gross profit, margin, product and salesperson rankings, monthly trend, and high-revenue/low-margin items. Put uncertain joins and records needing review in a separate sheet. Create a manager-ready workbook without overwriting the source.

## Finance and cash

> Reconcile invoices, receivables, payments, and expenses as of the latest reliable date. Use only fields that exist in the files. Build aging buckets, overdue customer risk, cash inflow/outflow by month, budget variance if budgets are present, and an evidence-based action list. Do not infer a payment date from a text note.

## Inventory and purchasing

> Calculate current stock as opening balance plus receipts minus issues plus adjustments. Compare each item with safety stock, recent demand, and target days where available. Identify stockouts, slow-moving stock, negative balances, and purchase-price anomalies. Separate facts, assumptions, and approvals needed from the recommended actions.

## HR and labor efficiency

> Combine employee master data, attendance, overtime, performance, and payroll when keys and periods match. Report headcount, overtime rate, output per labor hour, pay changes, missing records, and employees needing review. Avoid ranking employees when the underlying period or KPI definitions are not comparable.

## Restaurant or service operations

> Link store, POS, delivery settlement, refunds, purchasing, recipe/BOM, waste, labor, fixed costs, and reviews where the keys support it. Explain whether growth creates contribution profit, where platform fees or discounts erode margin, which products need attention, and where waste, labor, or complaints cluster. Mark scale or unit mismatches for confirmation.

## Manufacturing and management diagnosis

> Build a management diagnosis from the available tables. Find the largest evidence-backed risks across customers, sales, costs, inventory, cash, quality, and people. For every risk provide evidence, impact, owner, priority, next action, and an acceptance metric. If the data cannot support a conclusion, state the missing field and request manual confirmation.

## Turn a showcase into a contribution

When a prompt reveals a missing capability, open a GitHub issue with:

- the anonymized table schema and a few synthetic rows;
- the business question and expected output;
- what the current result got wrong;
- privacy constraints and a reproducible test case.

Never attach customer workbooks, API keys, access tokens, or identifiable personal data.
