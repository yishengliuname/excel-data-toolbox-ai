# Excel Data Toolbox AI

Turn messy Excel files into clean, analyzed, ready-to-deliver workbooks with one prompt.

AI understands your request. Deterministic local code handles the spreadsheet.

English | [简体中文](README_ZH.md)

[![CI](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/ci.yml/badge.svg)](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/ci.yml)
[![CodeQL](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/codeql.yml/badge.svg)](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/codeql.yml)
[![License: Apache-2.0](https://img.shields.io/badge/License-Apache--2.0-blue.svg)](LICENSE)
[![Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-3776AB.svg)](https://www.python.org/downloads/)

## What it does

- Clean messy Excel and CSV data.
- Merge, match, and reconcile business tables.
- Analyze sales, finance, inventory, and operational data.
- Generate native Excel charts and management-ready reports.
- Validate outputs without overwriting the original workbook.

## Why not just let an LLM edit Excel directly?

| LLM-only workflow | Excel Data Toolbox AI |
|---|---|
| Model guesses operations | AI creates a structured plan |
| Model may modify data directly | A local whitelist validates operations |
| Hard to audit | Execution and validation are traceable |
| Calculations may be hallucinated | Deterministic Python performs calculations |
| The original workbook may change | Source files remain untouched |

The model interprets the request; it does not receive arbitrary code or file-system access. Ambiguous fields, unsafe parameters, and missing evidence are stopped or marked for review before delivery.

## Quick start

### Windows release

For the shortest setup, open [Releases](https://github.com/yishengliuname/excel-data-toolbox-ai/releases), download `Excel-Data-Toolbox-AI-Windows-x64.zip`, extract it, and run `BiaogeKuaichuAI.exe`.

The Windows package does not require Git or a separate Python installation. If a release is not published yet, use the source installation below.

### Run from source

Python 3.11 or newer is required.

Windows PowerShell:

```powershell
git clone https://github.com/yishengliuname/excel-data-toolbox-ai.git
cd excel-data-toolbox-ai
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[automation]"
python -m excel_data_toolbox.server
```

Linux or macOS:

```bash
git clone https://github.com/yishengliuname/excel-data-toolbox-ai.git
cd excel-data-toolbox-ai
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e '.[automation]'
python -m excel_data_toolbox.server
```

Open `http://127.0.0.1:8501/`, select **Load demo data / 加载演示数据**, and try the workflow below. An AI provider is optional; local deterministic workflows work without an API key.

## See it in action

**Input:** the built-in fictional `sales.xlsx`-style dataset.

**Prompt:**

> Analyze revenue, cost, and profit by region and product. Find low-margin items, create a monthly trend chart, and export a management-ready workbook. Do not modify the original file.

**Output:**

- KPI summary
- Monthly trend
- Product contribution
- Low-margin warnings
- Native Excel charts
- Auditable Excel workbook

The demo is generated locally and contains no customer names, credentials, or private files. See the repeatable [demo guide](docs/DEMO.md) and the privacy-safe [recording specification](docs/DEMO_ASSETS.md).

If this project saves you time on real Excel work, consider giving it a ⭐. It helps more people discover the project.

## Architecture

```text
Natural-language request
          ↓
Intent and domain planning
          ↓
Safe parameter validation
          ↓
Deterministic local execution
          ↓
Output validation
          ↓
Excel workbook or report
```

Understanding and execution are deliberately separated. The optional model produces a constrained plan; local Python validates the plan, performs spreadsheet operations, and reopens exported files for delivery checks.

## Key capabilities

- Excel data cleaning, standardization, deduplication, merging, matching, splitting, and masking.
- Tolerance-aware reconciliation with duplicate-key isolation and manual-review candidates.
- Adaptive analysis based on the request, workbook structure, field semantics, and available evidence.
- Native Excel tables, formatting, charts, management summaries, risk lists, and action plans.
- Formula-injection protection, per-task storage, undo/redo, and output re-open validation.
- Reusable recipes, data-quality rules, processing lineage, and auditable delivery notes.

## Advanced capabilities

These capabilities are available when the workbook and installed dependencies provide enough evidence:

- Finance: receivables aging, budget variance, cash flow, ratios, and journal-entry checks.
- Operations: sales, inventory, HR, e-commerce, restaurant, and manufacturing analysis packs.
- Analytics: pivots, trends, contribution, anomalies, correlations, regression, and RFM segmentation.
- Files and data: PDF table extraction, image OCR, DuckDB summaries, and read-only SQLite/ODBC queries.
- Engineering handoffs: statically scanned VBA packages and Power BI model/DAX/PBIP artifacts.

Power BI publishing still requires valid Microsoft tenant, license, service-principal, and workspace permissions. OCR requires a compatible Tesseract installation. Subjective approvals and unsupported formats remain explicit human-review boundaries.

## Privacy and safety

- The server binds to `127.0.0.1` by default.
- Original workbooks are not overwritten.
- Each task has isolated input, output, audit, and review storage.
- Formula injection, unsafe SQL, dangerous VBA, resource limits, and output integrity are checked locally.
- Customer data, generated outputs, logs, databases, and `.env` files are excluded from the repository.
- Secrets are read server-side and are not written to plans, workbooks, or operation logs.
- Without an AI provider, no model request is made.

When an optional provider is enabled, the default planning request contains the user's instruction and selected table schema catalogue, not complete cell contents. Read [Privacy](docs/PRIVACY.md) and [Security](SECURITY.md) before using private business data.

## AI provider configuration (optional)

DeepSeek is currently supported as an optional planning provider. Copy `.env.example` to `.env` only if you want to enable it:

```text
DEEPSEEK_API_KEY=
DEEPSEEK_MODEL=deepseek-v4-flash
```

Restart the local app after configuration. Never commit `.env`, paste an active key into an issue, or include customer data in a public reproduction.

## Development

```bash
python -m pip install -e '.[automation,dev]'
python scripts/check_secrets.py
python -m ruff check . --select E9,F63,F7,F82
python scripts/check_docs_links.py
python -m pytest -q
python -m build
```

Useful project guides:

- [Architecture](docs/ARCHITECTURE.md)
- [Repeatable demo](docs/DEMO.md)
- [Maintainer release guide](docs/GITHUB_RELEASE.md)
- [Domain-pack guide](docs/ADDING_DOMAIN_PACK.md)
- [Roadmap](docs/ROADMAP.md)

## Contributing

Small documentation, testing, accessibility, packaging, and synthetic-demo improvements are welcome. Start with [Your first contribution](CONTRIBUTING.md#your-first-contribution) and the prepared [good first issues](docs/GOOD_FIRST_ISSUES.md).

Use the issue templates for reproducible bugs, real Excel-task proposals, and domain-pack vocabulary. Report security vulnerabilities privately according to [SECURITY.md](SECURITY.md). Do not attach customer workbooks or active credentials.

## Roadmap

The current focus is reliability, packaging, cross-platform verification, and contributor experience—not adding a new customer-specific branch for every workbook. See the public [roadmap](docs/ROADMAP.md).

## License

Released under the [Apache License 2.0](LICENSE). You may use, modify, and commercialize the project subject to the license and notice requirements. The software is provided as-is and is not accounting, tax, legal, audit, HR, or investment advice.
