# Maintainer release guide

Use this checklist for every public release. The current package version is defined in `pyproject.toml`; the application version and release tag must match it.

## 1. Verify the working tree

```powershell
git status --short
git diff --check
```

Do not release from a tree containing `.env`, customer files, outputs, logs, databases, caches, or generated binaries.

## 2. Run safety and quality gates

```powershell
python -m pip install -e ".[automation,dev]"
python scripts/check_secrets.py
python -m ruff check . --select E9,F63,F7,F82
python scripts/check_docs_links.py
python -m pytest -q
python -m build
```

## 3. Verify the version

For v9.9.0:

```powershell
python scripts/check_release_version.py --tag v9.9.0
```

The check compares the tag, `pyproject.toml`, and the application version exposed by `/health`.

## 4. Build the Windows package

On Windows:

```powershell
.\scripts\build_windows.ps1 -RecreateEnvironment
```

The script creates an isolated build environment, builds the existing PyInstaller specification, runs a local HTTP/demo/export smoke test, and creates `dist/Excel-Data-Toolbox-AI-Windows-x64.zip`.

The ZIP must not contain `.env`, customer data, outputs, logs, test caches, or development sources.

## 5. Create and push the tag

Only tag a commit after the checks above pass and CI on `main` is green.

```powershell
git tag -a v9.9.0 -m "Excel Data Toolbox AI v9.9.0"
git push origin v9.9.0
```

Pushing a `v*` tag triggers `.github/workflows/release.yml`. Add matching notes at `docs/RELEASE_NOTES_<tag>.md` before tagging (for example, `docs/RELEASE_NOTES_v9.9.0.md`). The workflow repeats the safety scan, static checks, tests, version check, Python build, Windows package build, and executable smoke test before creating the GitHub Release with the Windows ZIP, wheel, and source distribution.

## 6. Verify the GitHub Release

1. Confirm the workflow is green.
2. Download `Excel-Data-Toolbox-AI-Windows-x64.zip` from a clean Windows computer.
3. Extract it outside the repository.
4. Run `BiaogeKuaichuAI.exe`.
5. Confirm `/health` reports the release version.
6. Load synthetic demo data.
7. Export a workbook and open it in desktop Excel or LibreOffice.
8. Confirm no API key is required for the local demo.

Do not call a release complete until the clean-machine download test passes.

## Release safety notes

- Revoke any credential that has ever appeared in a public commit, issue, screenshot, or Actions log.
- Keep Actions permissions read-only except for the release job's `contents: write` permission.
- Publish binaries only as GitHub Release assets; never commit `.exe`, `.zip`, `build/`, or `dist/`.
- Use only synthetic data in release screenshots and smoke tests.
