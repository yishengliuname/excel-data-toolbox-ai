# Five-minute showcase

This script is designed for a README GIF, a release note, or a short product post. It uses the built-in synthetic data and never requires a customer workbook.

## 1. Start the local app

```powershell
python -m excel_data_toolbox.server
```

Open `http://127.0.0.1:8501/`, click **加载演示数据**, and keep the task local.

## 2. Paste one business request

```text
Analyze the selected sales data. First show the source sheets, row counts,
missing values, and duplicate order IDs. Then calculate revenue, cost, profit,
margin, monthly trend, product contribution, and salesperson ranking. Flag
high-revenue/low-margin items and records that need manual review. Export a
manager-ready workbook without changing the source data.
```

## 3. Show the proof, not only the chart

The strongest demo sequence is:

1. The source list and the selected sheets.
2. The natural-language request.
3. The plan preview and local safety checks.
4. The generated KPI, risk, and human-review sections.
5. A native Excel chart with axis titles and a downloadable workbook.

Use a synthetic workbook or the built-in demo only. Do not record API keys, customer names, phone numbers, screenshots of private files, or full cell values sent to an external service.

## 4. A 60-second post

```text
Messy Excel should not require a maze of formulas. Excel Data Toolbox AI
turns a plain-language request into a validated, auditable workbook while
keeping deterministic processing local. DeepSeek is optional.

Try the synthetic demo and tell us which domain pack you need next:
https://github.com/yishengliuname/excel-data-toolbox-ai
```

## 5. What to measure

After publishing, record weekly GitHub views, unique visitors, clones, stars, issue activity, and the top referral source. A post that gets views but no clones needs a clearer quick start; a post that gets clones but no issues may need better onboarding or a smaller first contribution.
