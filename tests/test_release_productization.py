from __future__ import annotations

from pathlib import Path
import subprocess
import sys
import tomllib


ROOT = Path(__file__).resolve().parents[1]


def current_version() -> str:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        return str(tomllib.load(handle)["project"]["version"])


def run_script(name: str, *arguments: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(ROOT / "scripts" / name), *arguments],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )


def test_release_version_matches_package_and_runtime() -> None:
    version = current_version()
    result = run_script("check_release_version.py", "--tag", f"v{version}")
    assert result.returncode == 0, result.stderr
    assert f"v{version}" in result.stdout


def test_release_version_rejects_mismatched_tag() -> None:
    result = run_script("check_release_version.py", "--tag", "v0.0.0")
    assert result.returncode != 0
    assert "release version mismatch" in result.stderr


def test_public_documentation_has_no_broken_internal_links() -> None:
    result = run_script("check_docs_links.py")
    assert result.returncode == 0, result.stdout + result.stderr


def test_pyinstaller_spec_includes_runtime_assets() -> None:
    spec = (ROOT / "BiaogeKuaichuAI.spec").read_text(encoding="utf-8")
    assert "('web', 'web')" in spec
    assert "('domain_packs.json', '.')" in spec
