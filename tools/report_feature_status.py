#!/usr/bin/env python3
"""Report built/scaffolded/missing feature evidence for the public template."""

from __future__ import annotations

import argparse
import csv
import sys
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_STATUS_PATH = ROOT / "docs" / "feature_status.tsv"

VALID_STATUSES = {"Built", "Scaffolded", "Missing", "Mismatch"}


@dataclass
class FeatureResult:
    feature_id: str
    feature_name: str
    expected_status: str
    actual_status: str
    category: str
    detail: str
    notes: str

    @property
    def ok(self) -> bool:
        return not (self.expected_status == "Built" and self.actual_status == "Missing")


def read_text(relative_path: str) -> str:
    path = ROOT / relative_path
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return ""


def path_exists(relative_path: str) -> bool:
    return (ROOT / relative_path).exists()


def evaluate_pattern(pattern: str) -> tuple[bool, str]:
    if pattern.startswith("contains:"):
        payload = pattern.removeprefix("contains:")
        if "::" not in payload:
            return False, "contains pattern missing :: separator"
        relative_path, token_blob = payload.split("::", 1)
        text = read_text(relative_path)
        tokens = [token for token in token_blob.split("||") if token]
        missing = [token for token in tokens if token not in text]
        if not text:
            return False, f"{relative_path} missing or unreadable"
        if missing:
            return False, f"{relative_path} missing token(s): {', '.join(missing)}"
        return True, f"{relative_path} contains {len(tokens)} token(s)"

    if pattern.startswith("all_exists:"):
        paths = [path for path in pattern.removeprefix("all_exists:").split("||") if path]
        missing = [path for path in paths if not path_exists(path)]
        if missing:
            return False, f"missing path(s): {', '.join(missing)}"
        return True, f"{len(paths)} path(s) exist"

    if pattern.startswith("exists:"):
        relative_path = pattern.removeprefix("exists:")
        exists = path_exists(relative_path)
        return exists, f"{relative_path} {'exists' if exists else 'missing'}"

    return False, f"unsupported evidence pattern: {pattern}"


def actual_status(expected_status: str, evidence_ok: bool) -> str:
    if expected_status == "Missing":
        return "Mismatch" if evidence_ok else "Missing"
    if expected_status == "Scaffolded":
        return "Scaffolded" if evidence_ok else "Missing"
    if expected_status == "Built":
        return "Built" if evidence_ok else "Missing"
    return "Mismatch"


def load_results(status_path: Path = DEFAULT_STATUS_PATH) -> list[FeatureResult]:
    with status_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle, delimiter="\t")
        results: list[FeatureResult] = []
        for row in reader:
            expected = row["ExpectedStatus"].strip()
            if expected not in VALID_STATUSES:
                evidence_ok = False
                detail = f"invalid ExpectedStatus: {expected}"
                actual = "Mismatch"
            else:
                evidence_ok, detail = evaluate_pattern(row["EvidencePattern"].strip())
                actual = actual_status(expected, evidence_ok)
            results.append(
                FeatureResult(
                    feature_id=row["FeatureId"].strip(),
                    feature_name=row["FeatureName"].strip(),
                    expected_status=expected,
                    actual_status=actual,
                    category=row["Category"].strip(),
                    detail=detail,
                    notes=row["Notes"].strip(),
                )
            )
    return results


def format_console(results: list[FeatureResult]) -> str:
    lines = ["Feature status:"]
    for result in results:
        marker = "OK" if result.ok else "FAIL"
        lines.append(
            f"{marker:<4} {result.actual_status:<10} expected={result.expected_status:<10} "
            f"{result.feature_id} - {result.feature_name}"
        )
        lines.append(f"     {result.detail}")
    counts = {status: sum(1 for result in results if result.actual_status == status) for status in sorted(VALID_STATUSES)}
    lines.append(
        "Summary: "
        + ", ".join(f"{counts[status]} {status}" for status in ("Built", "Scaffolded", "Missing", "Mismatch"))
    )
    return "\n".join(lines)


def format_markdown(results: list[FeatureResult]) -> str:
    lines = [
        "# Feature Status",
        "",
        "| Feature | Category | Expected | Actual | Evidence |",
        "| --- | --- | --- | --- | --- |",
    ]
    for result in results:
        lines.append(
            f"| `{result.feature_id}` | {result.category} | {result.expected_status} | "
            f"{result.actual_status} | {result.detail} |"
        )
    lines.extend(
        [
            "",
            "## Built",
            *[f"- `{result.feature_id}`: {result.feature_name}" for result in results if result.actual_status == "Built"],
            "",
            "## Scaffolded",
            *[
                f"- `{result.feature_id}`: {result.feature_name}"
                for result in results
                if result.actual_status == "Scaffolded"
            ],
            "",
            "## Missing",
            *[
                f"- `{result.feature_id}`: {result.feature_name}"
                for result in results
                if result.actual_status == "Missing"
            ],
            "",
            "## Mismatch",
            *[
                f"- `{result.feature_id}`: expected {result.expected_status}, evidence says {result.actual_status}"
                for result in results
                if result.actual_status == "Mismatch"
            ],
        ]
    )
    return "\n".join(lines).rstrip() + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description="Report feature implementation status.")
    parser.add_argument("--markdown", type=Path, help="Optional Markdown output path.")
    args = parser.parse_args()

    results = load_results()
    print(format_console(results))

    if args.markdown:
        output_path = args.markdown if args.markdown.is_absolute() else ROOT / args.markdown
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(format_markdown(results), encoding="utf-8")
        print(f"Wrote Markdown feature status: {output_path}")

    missing_built = [result for result in results if not result.ok]
    return 1 if missing_built else 0


if __name__ == "__main__":
    sys.exit(main())
