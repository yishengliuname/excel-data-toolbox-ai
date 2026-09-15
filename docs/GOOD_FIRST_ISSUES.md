# Prepared good first issues

These are real, bounded contribution candidates. Copy them into GitHub Issues after confirming that no equivalent issue already exists.

## 1. Verify the Windows release on a clean Windows 10 machine

**Labels:** `good first issue`, `help wanted`, `testing`

**Scope:** Download the latest ZIP, start the executable, load synthetic demo data, export one workbook, and report the exact result.

**Acceptance criteria:**

- Record Windows edition and package version.
- Confirm `/health`, demo loading, export, and workbook reopening.
- Attach only synthetic screenshots.
- Document any antivirus warning without bypassing it.

## 2. Add keyboard-focus visibility to the single-command interface

**Labels:** `good first issue`, `help wanted`, `accessibility`

**Scope:** Improve visible `:focus-visible` styles for buttons, the upload control, table selectors, and the command input in `web/`.

**Acceptance criteria:**

- Every interactive element is reachable by keyboard.
- Focus is visible against both white and green backgrounds.
- Existing hover styles and layout remain unchanged.
- Add a short manual verification note to the pull request.

## 3. Add a synthetic workbook with locale-formatted numbers

**Labels:** `good first issue`, `help wanted`, `testing`

**Scope:** Extend tests with fictional values such as `1,234.50`, `1.234,50`, percentages, blanks, and negative parentheses.

**Acceptance criteria:**

- No static `.xlsx` fixture is committed; construct it in the test.
- Expected parse results are explicit.
- Ambiguous formats produce a review warning instead of a guessed value.
- The full test suite remains green.

## 4. Improve the Linux/macOS source-install troubleshooting guide

**Labels:** `good first issue`, `help wanted`, `documentation`

**Scope:** Verify source installation on one current Linux distribution or macOS version and document only observed problems.

**Acceptance criteria:**

- Record OS and Python version.
- Verify local server, synthetic demo, and one export.
- Keep optional OCR/ODBC notes separate from the base install.
- All internal Markdown links pass `scripts/check_docs_links.py`.

## 5. Add a reduced-motion mode to public demo styling

**Labels:** `good first issue`, `help wanted`, `accessibility`

**Scope:** Audit CSS transitions/animations and add `prefers-reduced-motion` handling where needed.

**Acceptance criteria:**

- No functional behavior changes.
- The default visual appearance remains unchanged.
- Reduced-motion users receive no nonessential animation.
- Include before/after notes using synthetic data only.
