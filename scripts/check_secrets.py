"""Fail safely when publishable files contain likely secrets or customer data.

The scanner prints only file names and rule identifiers.  It never prints a
matched credential.  GitHub secret scanning remains the second, history-aware
layer after publication.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import re
import subprocess
import sys


EXCLUDED_DIRECTORIES = {
    ".git", ".venv", "venv", "env", "__pycache__", ".pytest_cache",
    ".test_cache", ".ruff_cache", "build", "dist", "outputs", "user_data",
    "acceptance_reports", "releases", "downloads", ".artifact_review",
    ".artifact_work", ".tmp",
}
PROHIBITED_SUFFIXES = {
    ".xlsx", ".xlsm", ".xls", ".csv", ".tsv", ".parquet", ".sqlite",
    ".sqlite3", ".db", ".duckdb", ".log", ".p12", ".pfx", ".jks", ".key", ".pem",
}
TEXT_SUFFIXES = {
    ".py", ".pyi", ".js", ".mjs", ".html", ".css", ".json", ".toml",
    ".ini", ".cfg", ".yaml", ".yml", ".md", ".txt", ".ps1", ".cmd",
    ".bat", ".sh", ".xml", ".spec", "",
}
ALLOWED_EXAMPLE_VALUES = {
    "", "changeme", "change-me", "your-key", "your-secret", "example",
    "placeholder", "<secret>", "<api-key>", "<your-key>", "dummy",
}

PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("private-key", re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----")),
    ("github-token", re.compile(r"\bgh[pousr]_[A-Za-z0-9_]{30,}\b")),
    ("generic-sk-token", re.compile(r"\bsk-[A-Za-z0-9_-]{20,}\b")),
    ("jwt-token", re.compile(r"\beyJ[A-Za-z0-9_-]{20,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b")),
    ("cloud-access-key", re.compile(r"\b(?:AKIA|ASIA)[A-Z0-9]{16}\b")),
    (
        "assigned-secret",
        re.compile(
            r"(?i)\b(?:api[_-]?key|client[_-]?secret|access[_-]?token|refresh[_-]?token|password|passwd)"
            r"\s*[:=]\s*[\"']([^\"'\r\n]{12,})[\"']"
        ),
    ),
)


def _tracked_files(root: Path) -> list[Path] | None:
    if not (root / ".git").exists():
        return None
    process = subprocess.run(
        # CI sandboxes can mount the checkout under a different OS owner.  A
        # per-process wildcard keeps this read-only probe portable without
        # mutating the user's global Git configuration.
        ["git", "-c", "safe.directory=*", "-C", str(root), "ls-files", "-z"],
        check=False, capture_output=True,
    )
    # A freshly initialized repository has no index entries yet.  Treat that
    # state like a non-git checkout so the pre-commit gate still scans the
    # working tree instead of silently passing an empty file list.
    if process.returncode != 0 or not process.stdout.strip():
        return None
    return [root / item.decode("utf-8", errors="surrogateescape") for item in process.stdout.split(b"\0") if item]


def _publishable_files(root: Path) -> list[Path]:
    tracked = _tracked_files(root)
    if tracked is not None:
        return [path for path in tracked if path.is_file()]
    files: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        relative = path.relative_to(root)
        if any(part in EXCLUDED_DIRECTORIES or part.startswith("tmp_") for part in relative.parts[:-1]):
            continue
        if relative.name == ".env" or (relative.name.startswith(".env.") and relative.name != ".env.example"):
            continue
        if path.suffix.casefold() in PROHIBITED_SUFFIXES:
            continue
        files.append(path)
    return files


def scan(root: Path) -> list[tuple[str, str]]:
    findings: list[tuple[str, str]] = []
    for path in _publishable_files(root):
        relative = path.relative_to(root).as_posix()
        if path.name == ".env":
            findings.append((relative, "forbidden-env-file"))
            continue
        if path.suffix.casefold() in PROHIBITED_SUFFIXES:
            findings.append((relative, "forbidden-data-or-secret-file"))
            continue
        if path.suffix.casefold() not in TEXT_SUFFIXES or path.stat().st_size > 2_000_000:
            continue
        try:
            text = path.read_text(encoding="utf-8-sig")
        except (OSError, UnicodeDecodeError):
            continue
        for rule, pattern in PATTERNS:
            for match in pattern.finditer(text):
                matched_text = match.group(0).casefold()
                if "unit-test" in matched_text or "fake-" in matched_text:
                    continue
                if rule == "assigned-secret":
                    value = match.group(1).strip().casefold()
                    if (
                        value in ALLOWED_EXAMPLE_VALUES or "${" in value or "your_" in value
                        or "unit-test" in value or value.startswith("fake-")
                    ):
                        continue
                findings.append((relative, rule))
                break
    return sorted(set(findings))


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan publishable repository files without printing secret values")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    args = parser.parse_args()
    root = args.root.resolve()
    findings = scan(root)
    if findings:
        print("Secret/customer-data gate: FAILED")
        for filename, rule in findings:
            print(f"- {filename}: {rule}")
        print("Remove the file/value, rotate any exposed credential, then run this check again.")
        return 1
    print(f"Secret/customer-data gate: PASSED ({root})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
