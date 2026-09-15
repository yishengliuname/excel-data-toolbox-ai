"""Verify that release tag, package metadata, and runtime version agree."""

from __future__ import annotations

import argparse
import ast
from pathlib import Path
import re
import tomllib


ROOT = Path(__file__).resolve().parents[1]
TAG_PATTERN = re.compile(r"^v(?P<version>\d+\.\d+\.\d+(?:[A-Za-z0-9.-]+)?)$")


def project_version() -> str:
    with (ROOT / "pyproject.toml").open("rb") as handle:
        return str(tomllib.load(handle)["project"]["version"])


def runtime_version() -> str:
    module = ast.parse((ROOT / "server.py").read_text(encoding="utf-8"))
    for node in module.body:
        if not isinstance(node, (ast.Assign, ast.AnnAssign)):
            continue
        targets = node.targets if isinstance(node, ast.Assign) else [node.target]
        if not any(isinstance(target, ast.Name) and target.id == "APP_VERSION" for target in targets):
            continue
        value = node.value
        if isinstance(value, ast.Constant) and isinstance(value.value, str):
            return value.value
    raise ValueError("server.py must define APP_VERSION as a string literal")


def check(tag: str) -> str:
    match = TAG_PATTERN.fullmatch(tag.strip())
    if not match:
        raise ValueError("release tag must use v<major>.<minor>.<patch>, for example v9.9.0")
    tag_version = match.group("version")
    package_version = project_version()
    app_version = runtime_version()
    if len({tag_version, package_version, app_version}) != 1:
        raise ValueError(
            "release version mismatch: "
            f"tag={tag_version}, pyproject.toml={package_version}, server.APP_VERSION={app_version}"
        )
    return tag_version


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--tag", help="Git tag such as v9.9.0")
    parser.add_argument("--print-version", action="store_true", help="Print the package version")
    args = parser.parse_args()
    if args.print_version:
        print(project_version())
        return
    if not args.tag:
        parser.error("--tag is required unless --print-version is used")
    version = check(args.tag)
    print(f"Release version check passed: v{version}")


if __name__ == "__main__":
    main()
