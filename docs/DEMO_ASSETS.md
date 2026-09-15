# Demo asset specification

Only publish assets created from the built-in synthetic dataset. Never use a customer workbook as a visual example.

## Current status

No screenshot or GIF is committed yet. The automated browser available during the productization pass could not render a reliable screenshot, so the README deliberately contains no placeholder or broken asset. Add `hero.png` only after following the process below.

## Recommended asset set

| File | Required scene | Acceptance criteria |
|---|---|---|
| `hero.png` | App with synthetic data loaded | Product name, one-command input, and selected fictional table are readable |
| `workflow.png` | Structured plan and safety check | No key, local private path, or customer data |
| `dashboard.png` | KPI and chart result | Axis title, unit, legend, and warning are readable |
| `result-workbook.png` | Exported workbook in Excel | Native chart and audit/summary sheet visible |
| `demo.gif` | Full 15–30 second flow | Follows the recording guide in `docs/DEMO.md` |

Do not add a README link until the referenced asset exists and has passed the privacy checklist.

## Regeneration notes

1. Set `BIAOGE_USER_DATA` to a new temporary directory.
2. Start the app with `--no-browser` on a non-default local port.
3. Call the built-in `/api/demo` endpoint with an empty JSON object.
4. Capture only the local application window.
5. Run `python scripts/check_secrets.py` before committing the asset.
