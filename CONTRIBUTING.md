# Contributing

Thank you for helping improve Excel Data Toolbox AI. Documentation, tests, accessibility, packaging, safe data handling, and focused bug fixes are welcome.

## Your first contribution

1. Run the project from source.
2. Load the built-in synthetic demo data.
3. Pick a prepared [good first issue](docs/GOOD_FIRST_ISSUES.md).
4. Create a short branch from `main`.
5. Run the required checks.
6. Open a focused pull request.

This should be enough to start. The sections below describe the project safety boundary.

## Before you begin

- Never submit customer workbooks, screenshots, databases, logs, personal information, or proprietary templates.
- Never place API keys, passwords, tokens, private keys, or connection strings in code, issues, pull requests, commits, or fixtures.
- Keep facts, inferences, recommendations, and human approvals separate in financial, HR, audit, and risk outputs.
- AI output must continue through the local whitelist and parameter validation; it must not execute arbitrary code, SQL, or file operations.

## Local development

```bash
git clone https://github.com/yishengliuname/excel-data-toolbox-ai.git
cd excel-data-toolbox-ai
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[automation,dev]"
```

Linux or macOS:

```bash
source .venv/bin/activate
python -m pip install -e '.[automation,dev]'
```

DeepSeek is optional. Local deterministic cleaning, analysis, reporting, demo, and tests must work without an API key.

## Required checks

```bash
python scripts/check_secrets.py
python -m ruff check . --select E9,F63,F7,F82
python scripts/check_docs_links.py
python -m pytest -q
python -m build
```

All test data must be generated in code or have clear permission and provenance. Do not bypass repository safety checks with force-add flags.

## Pull-request scope

1. Search existing issues before starting.
2. For a larger change, describe the real Excel task and evidence boundary first.
3. Create a focused branch such as `fix/chart-axis-title`.
4. Add or update automated tests.
5. Explain the input shape, expected output, risk boundary, and exact verification commands.
6. Use synthetic screenshots for chart or layout changes.

New industries should normally extend `domain_packs.json`, not add a customer-name branch. Add a dedicated calculation module only when the business rule is stable, auditable, and covered by tests. See [Adding a domain pack](docs/ADDING_DOMAIN_PACK.md).

## Pull-request acceptance

- Source files remain unchanged.
- Historical outputs cannot become new-task inputs by accident.
- Ratios, balances, amounts, and scores use explicit aggregation semantics.
- Missing evidence is reported rather than invented.
- Exported workbooks reopen with expected sheets, dimensions, fingerprints, and totals.
- New dependencies have a documented purpose.

By contributing, you agree that your contribution is licensed under Apache-2.0.
