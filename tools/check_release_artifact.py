#!/usr/bin/env python3
"""Scan generated workbook packages for public-release metadata hazards."""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_SUFFIXES = {".xlsx", ".xltx"}


def forbidden_needles() -> list[str]:
    return [
        "C:" + "\\" + "Users",
        "/" + "Users/",
        "." + "codex",
        "public" + "-exports",
        "release" + "_artifacts",
        "Jared " + "Bidlow",
        "#REF!",
        "#VALUE!",
        "#N/A",
        "#NAME?",
        "#DIV/0!",
        "HYPERLINK(",
    ]


def discover_workbooks(paths: list[Path]) -> list[Path]:
    workbooks: list[Path] = []
    for path in paths:
        resolved = path if path.is_absolute() else ROOT / path
        if resolved.is_dir():
            workbooks.extend(
                child
                for child in resolved.rglob("*")
                if child.suffix.lower() in WORKBOOK_SUFFIXES and not child.name.startswith("~$")
            )
        elif resolved.suffix.lower() in WORKBOOK_SUFFIXES and not resolved.name.startswith("~$"):
            workbooks.append(resolved)
    return sorted(set(workbooks))


def scan_workbook(path: Path) -> list[str]:
    findings: list[str] = []
    needles = forbidden_needles()
    try:
        with zipfile.ZipFile(path) as package:
            for member in package.namelist():
                if not (member.endswith(".xml") or member.endswith(".rels")):
                    continue
                data = package.read(member).decode("utf-8", errors="replace")
                upper_data = data.upper()
                for needle in needles:
                    haystack = upper_data if needle == "HYPERLINK(" else data
                    target = needle.upper() if needle == "HYPERLINK(" else needle
                    if target in haystack:
                        findings.append(f"{path.name}:{member}: contains {needle}")
    except zipfile.BadZipFile:
        findings.append(f"{path}: not a valid workbook zip package")
    except PermissionError as exc:
        findings.append(f"{path}: cannot read package: {exc}")
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description="Scan generated workbook artifacts.")
    parser.add_argument("paths", nargs="*", type=Path, default=[Path("release_artifacts")])
    args = parser.parse_args()

    workbooks = discover_workbooks(args.paths)
    if not workbooks:
        print("No workbook artifacts found.")
        return 0

    findings: list[str] = []
    for workbook in workbooks:
        findings.extend(scan_workbook(workbook))

    if findings:
        print("Release artifact scan failed:")
        for finding in findings:
            print(f"- {finding}")
        return 1

    print(f"Release artifact scan passed: {len(workbooks)} workbook package(s) checked.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
