# Contributor roadmap

The project is intentionally broad inside, but each contribution should be small, testable, and reusable across customer domains. Do not add a customer-specific branch or commit real files.

## Good first issues

### Chart accessibility

Add explicit x-axis and y-axis titles to every native chart, with a test for bar, line, and scatter charts. Acceptance: titles come from the compiled semantic roles or an explicit user request; no hard-coded sales-only labels.

### Filename safety

Add regression coverage for Chinese, Japanese, and emoji filenames on Windows and Linux. Acceptance: the display name is readable, the stored path is safe, and the original file is not renamed.

### Output isolation

Add a test proving that a report in `outputs/` cannot be selected as a new task input unless the user explicitly selects it. Acceptance: the test covers a previous report with dashboard-style sheets.

### Domain fixture

Add a small synthetic fixture for one domain pack such as inventory, finance, HR, restaurant, or manufacturing. Acceptance: the fixture contains only generated data, a reproducible request, expected metrics, and a passing test.

### Documentation

Improve one setup page or add a translated example. Acceptance: a new user can follow it without an API key and without seeing private data.

## Help-wanted issues

- Add chart theme presets that preserve contrast and readable labels.
- Add a finance aging example with explicit date and currency assumptions.
- Add a restaurant contribution-margin fixture covering discounts, platform fees, waste, and labor.
- Add a performance benchmark for 100k-row CSV summaries.
- Add a Power BI artifact validation fixture without publishing to a tenant.

## Pull request checklist

```text
- [ ] The change works for more than one business domain or is clearly scoped.
- [ ] No customer data, API key, token, screenshot, or private URL is included.
- [ ] Tests cover the new behavior and the ambiguous/empty case.
- [ ] Original inputs remain unchanged and output isolation is preserved.
- [ ] Facts, inference, recommendation, and manual review remain distinct.
- [ ] `check_secrets.py`, Ruff, pytest, and build pass locally.
```

## Contributor workflow

```bash
git clone https://github.com/yishengliuname/excel-data-toolbox-ai.git
cd excel-data-toolbox-ai
git checkout -b feat/short-description
# make a focused change and add tests
git add .
git commit -m "feat: describe the change"
git push origin feat/short-description
```

Open a pull request against `main`. CI and CodeQL run automatically. For a security concern, do not open a public issue; follow [SECURITY.md](../SECURITY.md).
