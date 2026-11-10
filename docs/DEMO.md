# Order reconciliation demo / 电商订单对账演示

This demo uses two fictional tables. It runs on your computer without an API key.

## Run it / 如何试用

1. Download the [Windows release](https://github.com/yishengliuname/excel-data-toolbox-ai/releases/latest), extract the ZIP, and run `BiaogeKuaichuAI.exe`. You can also [run from source](../README.md#run-from-source).
2. Click **试用电商订单对账** below the command box. No files or prompt are needed. The app opens a separate demo task with six order rows and five payout rows, matches exact order IDs, and flags any nonzero amount difference.
3. Click **下载演示对账 Excel** and open the workbook. The release page also provides the workbook as `Synthetic-Order-Reconciliation.xlsx`.

下载并解压 Windows 包，打开程序后点击 **试用电商订单对账**，再点 **下载演示对账 Excel**。对账规则按订单号精确匹配，金额差异阈值为 0。演示用单独任务，无需上传文件。

## Verified results / 已验证结果

| Result / 结果 | Count / 数量 | What to check / 如何核对 |
| --- | ---: | --- |
| Exact matches / 金额一致 | 2 orders | `DEMO-001`, `DEMO-005` |
| Amount difference / 金额差异 | 1 order | `DEMO-002` has 158 due and 148 paid, a shortfall of 10 |
| Orders only / 仅在订单表 | 1 order | `DEMO-003` |
| Payouts only / 仅在结算表 | 1 order | `DEMO-006` |
| Quarantined duplicates / 重复键隔离 | 3 rows | `DEMO-004` has two order rows and one payout row, all kept for review |

The workbook includes both input sheets, a reconciliation summary, matched and unmatched records, an amount-difference review sheet, and duplicate evidence. Start with **演示对账_对账摘要**, then open **演示对账_待人工确认** to see the `-10` difference. The Windows release check generates the workbook and reads it back before publishing.

## Sales demo / 销售分析演示

For a natural-language example, click **加载演示数据** and run the prefilled request to analyze sales by month and region. When the job completes, download its Excel report. This example also runs without an API key.

## Help test it / 参与试用

If anything is unclear, tell us your OS, app version, and the first step that was difficult. Feedback in English or Chinese is welcome. [Open a GitHub issue](https://github.com/yishengliuname/excel-data-toolbox-ai/issues/new/choose)

## Recording guide / 录制指引

Record a 15–25 second clip using the packaged app on a fresh profile. Show **试用电商订单对账** (0–5 s), the summary and export link (5–11 s), then the Excel summary and `DEMO-002` (11–23 s). Keep the labels readable at the final resolution.

Use built-in demo data only. Hide keys, customer files, local private paths, notifications, and browser profile details. See the [demo asset specification](DEMO_ASSETS.md).
