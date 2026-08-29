# Excel Data Toolbox AI

[![CI](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/ci.yml/badge.svg)](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/ci.yml)
[![CodeQL](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/codeql.yml/badge.svg)](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/codeql.yml)
[![License: Apache-2.0](https://img.shields.io/badge/License-Apache--2.0-blue.svg)](LICENSE)

**[中文 README](README.md)**

An open-source, local-first AI workspace for turning messy Excel work into an auditable deliverable:

> Natural-language request → business plan → safe local execution → validated Excel report

[Five-minute demo](docs/DEMO.md) · [Contributing](CONTRIBUTING.md) · [Contributor roadmap](docs/CONTRIBUTOR_ROADMAP.md)

Upload one or more Excel/CSV files, describe the outcome you need, and let the system select the appropriate cleaning, joining, reconciliation, analysis, visualization, or delivery workflow. DeepSeek is optional: the deterministic local engine can still inspect, transform, validate, and export data without an API key.

## Why this project

Most spreadsheet tools either force users to write formulas or return a generic chart. This project separates understanding from execution:

- AI interprets ordinary business language and maps it to a structured plan.
- A local whitelist validates fields, types, aggregations, resource limits, and dangerous operations before anything runs.
- Deterministic Python code performs the calculation and creates the workbook.
- Evidence gaps, ambiguous matches, and decisions requiring business judgement are explicitly marked for human review.
- Every delivery can be reopened and checked for sheet names, dimensions, fingerprints, and key totals.

## Highlights

- **One-command workflows** for cleaning, combining, matching, reconciliation, pivots, trends, rankings, anomaly detection, RFM, and business diagnostics.
- **Adaptive analysis compiler** that infers domain, table roles, grain, standard concepts, evidence gaps, metrics, and useful visualizations from the request and uploaded schemas. It does not assume every customer is an e-commerce or restaurant customer.
- **Business packs** for sales, finance, inventory, HR, e-commerce, restaurants, manufacturing, and general management reporting.
- **16 native visualizations** including bar, grouped/stacked bar, line, area, donut, waterfall, funnel, treemap, histogram, scatter/regression, box plot, heatmap, radar, and Gantt charts.
- **Professional deliverables** with management dashboards, ranking and contribution tables, risk registers, action plans, data-quality findings, audit notes, and human-review queues.
- **Advanced adapters** that generate safe VBA/static-scan packages, read-only SQLite/ODBC queries, Power BI star-schema/DAX/PBIP artifacts, PDF/table extraction, OCR, and DuckDB summaries where the local dependencies are available.
- **Privacy by default**: customer files and generated outputs stay in isolated task directories; the optional AI request contains the requirement and schema catalogue by default, not full cell values.

## Quick start

Requires Python 3.11 or newer.

### Windows PowerShell

```powershell
git clone https://github.com/yishengliuname/excel-data-toolbox-ai.git
cd excel-data-toolbox-ai
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[automation]"
Copy-Item .env.example .env
python -m excel_data_toolbox.server
```

### Linux or macOS

```bash
git clone https://github.com/yishengliuname/excel-data-toolbox-ai.git
cd excel-data-toolbox-ai
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e ".[automation]"
cp .env.example .env
python -m excel_data_toolbox.server
```

Open `http://127.0.0.1:8501/`. `DEEPSEEK_API_KEY` is optional for local deterministic workflows. If you enable DeepSeek, keep the key in `.env` or the operating-system secret store and never commit it.

## Try it with natural language

Select the sheets you want to use and enter a request such as:

> Combine the January, February, and March sales sheets. Remove exact duplicate order IDs, preserve all source rows in an audit sheet, compare revenue, cost, and profit by product, region, and salesperson, flag missing or conflicting fields for review, and create a management-ready workbook with a monthly trend chart and an action plan. Do not overwrite the source files.

Other useful starting points:

```text
Find overdue receivables as of 2026-08-31. Use due date, receivable amount,
paid amount, customer, and invoice number. Create aging detail, an aging
summary, and a customer risk table. Do not invent missing dates.
```

```text
For the inventory workbook, calculate opening stock + receipts - issues +
adjustments, compare with safety stock and target days, identify stockouts
and overstock, and separate facts, assumptions, and approvals needed from the
recommended actions.
```

See [copyable showcase prompts](examples/README.md) for sales, finance, inventory, HR, restaurant, and manufacturing examples. The examples use synthetic data only.

## A delivery that is useful to a manager

Depending on the data and request, the generated workbook can contain:

1. Management overview and key indicators.
2. Source and data-quality audit, including missing, duplicate, mixed-type, and conflicting fields.
3. Domain-specific analysis such as product/channel/customer/region, costs and margins, inventory turns, labor efficiency, or cash conversion.
4. Risk items with P0/P1/P2 priority, evidence, impact, owner, and recommended next step.
5. An action plan with acceptance metrics and a human-review queue.
6. Native charts chosen from the available fields and the business question, not from a fixed dashboard template.

The engine never silently deletes an anomaly or turns a guess into a fact. When a unit, definition, duplicate, or judgement cannot be established from the file, the result says what must be confirmed.

## Security boundary

- The local server listens on `127.0.0.1` by default.
- Original files are not overwritten; each task has its own input, output, log, and review area.
- Formula-injection protection, file-size/row limits, read-only SQL checks, and output re-open validation are enabled by default.
- VBA is statically scanned and does not bypass Office macro trust, passwords, MFA, or permissions.
- Power BI artifacts can be generated and locally checked; publishing still requires the customer’s valid tenant, license, service principal, and workspace permissions.
- This is an analysis and automation aid, not a substitute for an accountant, auditor, tax adviser, legal adviser, or a business owner’s approval.

## Configure DeepSeek (optional)

Copy `.env.example` to `.env` and set:

```text
DEEPSEEK_API_KEY=
DEEPSEEK_MODEL=deepseek-v4-flash
```

The key is read server-side only. It is not written into plans, logs, generated workbooks, screenshots, or the repository. Run `python scripts/check_secrets.py` before every public push.

## Development

New to the project? Start with the [five-minute demo](docs/DEMO.md) and the [contributor roadmap](docs/CONTRIBUTOR_ROADMAP.md). Pick an issue labelled `good first issue` or `help wanted`, or discuss an idea in [GitHub Discussions](https://github.com/yishengliuname/excel-data-toolbox-ai/discussions) before opening a pull request.

```bash
python scripts/check_secrets.py
python -m ruff check . --select E9,F63,F7,F82
python -m pytest -q
python -m build
```

Please read [CONTRIBUTING.md](CONTRIBUTING.md), [the architecture notes](docs/ARCHITECTURE.md), [the domain-pack guide](docs/ADDING_DOMAIN_PACK.md), and [the privacy model](docs/PRIVACY.md) before opening a pull request. Use the GitHub issue templates for bugs, feature requests, and new domain packs. Report security vulnerabilities privately according to [SECURITY.md](SECURITY.md).

## Current limits

The project is designed to be extensible rather than to claim that every file can be interpreted perfectly. Unsupported or ambiguous cases remain visible as manual-review items. Macros, subjective approval decisions, proprietary formats, unavailable OCR/database drivers, and cloud permissions may require customer confirmation or a separate handoff.

## License

Released under the [Apache License 2.0](LICENSE). You may use, modify, and commercialize the project subject to the license notices and conditions.
