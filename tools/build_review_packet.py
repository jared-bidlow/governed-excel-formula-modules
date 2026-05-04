#!/usr/bin/env python3
"""Build a compact review packet for branch and release review."""

from __future__ import annotations

import csv
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

import report_feature_status


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "release_artifacts" / "review_packet" / "review_packet.md"


def run_command(args: list[str]) -> tuple[int, str]:
    try:
        completed = subprocess.run(
            args,
            cwd=ROOT,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            check=False,
        )
        return completed.returncode, completed.stdout.strip()
    except FileNotFoundError as exc:
        return 127, str(exc)


def git_output(args: list[str]) -> str:
    code, output = run_command(["git", *args])
    if code != 0:
        return "Git metadata unavailable; generated from working tree files."
    return output or "(none)"


def command_summary(command: list[str], label: str) -> str:
    code, output = run_command(command)
    last_lines = "\n".join(output.splitlines()[-12:]) if output else "(no output)"
    return f"### {label}\n\nCommand: `{' '.join(command)}`\n\nExit code: `{code}`\n\n```text\n{last_lines}\n```\n"


def read_manifest_summary() -> str:
    path = ROOT / "samples" / "workbook_manifest.tsv"
    if not path.exists():
        return "Workbook manifest missing."
    with path.open("r", encoding="utf-8", newline="") as handle:
        rows = list(csv.DictReader(handle, delimiter="\t"))
    visible = [row["SheetName"] for row in rows if row.get("Visibility") == "visible"]
    hidden = [row["SheetName"] for row in rows if row.get("Visibility") == "hidden"]
    optional = [row["SheetName"] for row in rows if row.get("Presence") == "OptionalLegacy"]
    return "\n".join(
        [
            f"- Manifest rows: {len(rows)}",
            f"- Visible sheets: {', '.join(visible) if visible else '(none)'}",
            f"- Hidden/admin sheets: {len(hidden)}",
            f"- Optional legacy sheets: {', '.join(optional) if optional else '(none)'}",
        ]
    )


def power_query_summary() -> str:
    query_dir = ROOT / "samples" / "power-query" / "budget-input"
    if not query_dir.exists():
        return "Budget input Power Query folder missing."
    files = sorted(path.name for path in query_dir.glob("*.m"))
    selector = (query_dir / "qBudget_Source_Selected.m").read_text(encoding="utf-8", errors="replace")
    modes = [mode for mode in ["CurrentWorkbook", "AzureSQL", "Dataverse", "FabricSqlEndpoint"] if mode in selector]
    return "\n".join(
        [
            f"- Budget input templates: {', '.join(files)}",
            f"- Adapter selector modes: {', '.join(modes) if modes else '(selector missing modes)'}",
            "- CurrentWorkbook is the normal operator path.",
            "- Other source adapters are public-safe placeholders; credentials and tenant-specific endpoints are not included.",
        ]
    )


def asset_status_summary(results: list[report_feature_status.FeatureResult]) -> str:
    asset_results = [result for result in results if result.category == "Assets"]
    if not asset_results:
        return "No asset feature rows found."
    return "\n".join(
        f"- `{result.feature_id}`: expected {result.expected_status}, actual {result.actual_status}"
        for result in asset_results
    )


def main() -> int:
    output_path = DEFAULT_OUTPUT
    output_path.parent.mkdir(parents=True, exist_ok=True)

    feature_results = report_feature_status.load_results()
    feature_markdown = report_feature_status.format_markdown(feature_results)

    sections = [
        "# v0.5 Review Packet",
        "",
        f"Generated UTC: `{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')}`",
        "",
        "## Git State",
        "",
        f"- Branch: `{git_output(['branch', '--show-current'])}`",
        f"- HEAD: `{git_output(['rev-parse', '--short', 'HEAD'])}`",
        "",
        "### Working Tree",
        "",
        "```text",
        git_output(["status", "--short"]),
        "```",
        "",
        "### Changed Files",
        "",
        "```text",
        git_output(["diff", "--name-status"]),
        "```",
        "",
        "## Feature Status",
        "",
        feature_markdown,
        "",
        "## Validation Summaries",
        "",
        command_summary([sys.executable, "-S", "tools/audit_capex_module.py"], "Static Audit"),
        command_summary([sys.executable, "-S", "tools/lint_formulas.py", "modules/*.formula.txt"], "Formula Lint"),
        command_summary([sys.executable, "-S", "tools/report_feature_status.py"], "Feature Status"),
        "",
        "## Workbook Manifest Summary",
        "",
        read_manifest_summary(),
        "",
        "## Power Query Adapter Summary",
        "",
        power_query_summary(),
        "",
        "## Asset Workflow Status",
        "",
        asset_status_summary(feature_results),
        "",
        "## Public-Release Artifact Checks",
        "",
        "- Source builder contains package sanitization and public package assertions.",
        "- `tools/check_release_artifact.py` can scan generated `.xlsx` / `.xltx` packages without Excel COM.",
        "- Generated artifacts stay under ignored `release_artifacts/` and are not source artifacts.",
        "",
        "## Built / Scaffolded / Missing",
        "",
        "See the Feature Status section above for the generated Built, Scaffolded, Missing, and Mismatch lists.",
        "",
    ]

    output_path.write_text("\n".join(sections).replace("\r\n", "\n"), encoding="utf-8")
    print(f"Wrote review packet: {output_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
