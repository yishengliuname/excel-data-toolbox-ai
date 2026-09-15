"""Check repository-local links in public Markdown documents."""

from __future__ import annotations

from pathlib import Path
import re
from urllib.parse import unquote, urlsplit


ROOT = Path(__file__).resolve().parents[1]
LINK_PATTERN = re.compile(r"!?\[[^\]]*\]\(([^)]+)\)")
SCHEMES = {"http", "https", "mailto", "tel", "data"}


def markdown_files() -> list[Path]:
    top_level = [
        ROOT / "README.md",
        ROOT / "README_ZH.md",
        ROOT / "README_EN.md",
        ROOT / "CONTRIBUTING.md",
        ROOT / "SUPPORT.md",
        ROOT / "SECURITY.md",
        ROOT / "CHANGELOG.md",
    ]
    return [path for path in top_level if path.exists()] + sorted((ROOT / "docs").rglob("*.md"))


def local_target(source: Path, raw_target: str) -> Path | None:
    target = raw_target.strip().strip("<>").split(maxsplit=1)[0]
    if not target or target.startswith("#"):
        return None
    parsed = urlsplit(target)
    if parsed.scheme.lower() in SCHEMES or parsed.netloc:
        return None
    relative = unquote(parsed.path)
    if not relative:
        return None
    return (ROOT / relative.lstrip("/")) if relative.startswith("/") else (source.parent / relative)


def main() -> None:
    failures: list[str] = []
    checked = 0
    for source in markdown_files():
        text = source.read_text(encoding="utf-8")
        for line_number, line in enumerate(text.splitlines(), start=1):
            for match in LINK_PATTERN.finditer(line):
                target = local_target(source, match.group(1))
                if target is None:
                    continue
                checked += 1
                if not target.exists():
                    failures.append(
                        f"{source.relative_to(ROOT)}:{line_number}: missing {target.relative_to(ROOT)}"
                    )
    if failures:
        raise SystemExit("Broken internal Markdown links:\n" + "\n".join(failures))
    print(f"Documentation link check passed: {checked} internal links")


if __name__ == "__main__":
    main()
