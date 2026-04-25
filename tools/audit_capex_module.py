#!/usr/bin/env python3
"""Static audit for the public Excel formula-module template.

This audit is text-only. It does not open or edit Excel files.
"""

from __future__ import annotations

import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULES = ROOT / "modules"
MAX_NAMED_FORMULA_CHARS = 8192


def joined(*parts: str) -> str:
    """Build safety needles without self-matching this audit file."""
    return "".join(parts)


FORBIDDEN_TEXT = [
    joined("F", "M", "D", "C"),
    joined("P", "P", "L", " Corporation"),
    joined("C:", "\\", "Truth"),
    joined("C:", "\\", "Users", "\\", "e187258"),
    joined("One", "Drive - ", "P", "P", "L"),
    joined("2026 Capital Budget v", ".xlsx"),
    joined("Expanded_", "Validation_Workbook"),
    joined("Records", "Governance"),
    joined("q", "Share_", "ReviewGate_", "LinkedIn"),
    joined("q", "Share_", "PublicOutput_", "LinkedIn"),
    joined("Actual_", "Documents"),
    joined("202", "6 Budget"),
    joined("Meeting", " Note Flow & KPIs"),
    joined("Notes", " to Apply"),
    joined("tbl", "TO_APPLY"),
    joined("SP", " Number"),
    joined("ER", " Number"),
    joined("PM", " Notes"),
    joined("Start", " Timeline"),
    joined("New", "Meeting", "Notes"),
    joined("PM", "Notes"),
    joined("SP", "Number"),
    joined("ER", "Number"),
    joined("SP", "#"),
    joined("ER", "#"),
    joined("Start", "Timeline"),
    joined("ER", "Flag"),
    joined("204", "01"),
    joined("150", "00"),
    joined("153", "00000"),
    joined("430", "0000"),
    joined("Example", " BU"),
]

FORBIDDEN_PATTERNS = [
    (re.compile(r"[A-Za-z]:\\"), "local Windows path"),
    (re.compile(r"https?://", re.I), "URL"),
    (re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I), "email address"),
]

FORBIDDEN_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xls",
    ".zip",
}

REQUIRED_FORMULAS = {
    "modules/get.formula.txt": [
        "GetFinanceBlock",
        "GetBudgetDetailRows",
        "GetProjections",
        "GetActuals12",
    ],
    "modules/kind.formula.txt": [
        "CapConsumeMask",
        "CapByBU",
        "FlagsOut",
        "HiddenMaskLogic",
    ],
    "modules/Capex_Tracker_2026.formula.txt": [
        "CAPEX_REPORT",
    ],
    "modules/analysis.formula.txt": [
        "BU_CAP_SCORECARD_AXIS",
        "BU_CAP_SCORECARD",
        "REFORECAST_QUEUE_AXIS",
        "REFORECAST_QUEUE",
    ],
}

NAMED_FORMULA_BUDGETS = [
    ("modules/Capex_Tracker_2026.formula.txt", "CAPEX_REPORT"),
    ("modules/analysis.formula.txt", "BU_CAP_SCORECARD_AXIS"),
    ("modules/analysis.formula.txt", "BU_CAP_SCORECARD"),
    ("modules/analysis.formula.txt", "REFORECAST_QUEUE_AXIS"),
    ("modules/analysis.formula.txt", "REFORECAST_QUEUE"),
]


@dataclass
class Result:
    status: str
    file: str
    check: str
    detail: str
    suggestion: str = ""

    def render(self) -> str:
        out = f"{self.status:<5} {self.file} - {self.check}: {self.detail}"
        if self.suggestion:
            out += f"\n      suggestion: {self.suggestion}"
        return out


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        return path.read_text(encoding="utf-8", errors="replace")
    except FileNotFoundError:
        return ""


def rel(path: Path) -> str:
    return str(path.relative_to(ROOT)).replace("\\", "/")


def tracked_files() -> list[Path]:
    try:
        output = subprocess.check_output(
            ["git", "ls-files", "--cached", "--others", "--exclude-standard"],
            cwd=ROOT,
            text=True,
            stderr=subprocess.DEVNULL,
        )
        return [ROOT / line.strip() for line in output.splitlines() if line.strip()]
    except Exception:
        return [path for path in ROOT.rglob("*") if path.is_file() and ".git" not in path.parts]


def add(results: list[Result], ok: bool, file: str, check: str, detail: str, suggestion: str = "") -> None:
    results.append(Result("PASS" if ok else "FAIL", file, check, detail, "" if ok else suggestion))


def strip_block_comments(text: str) -> str:
    return re.sub(r"/\*.*?\*/", "", text, flags=re.S)


def balance_check(path: Path) -> list[Result]:
    text = read_text(path)
    label = rel(path)
    if not text:
        return [Result("FAIL", label, "file exists/readable", "missing or empty file")]

    clean = strip_block_comments(text)
    pairs = {"(": ")", "[": "]", "{": "}"}
    closes = {v: k for k, v in pairs.items()}
    stack: list[tuple[str, int, int]] = []
    in_string = False
    line = 1
    col = 0
    i = 0

    while i < len(clean):
        ch = clean[i]
        col += 1
        if ch == "\n":
            line += 1
            col = 0
            i += 1
            continue
        if ch == '"':
            if in_string and i + 1 < len(clean) and clean[i + 1] == '"':
                i += 2
                col += 1
                continue
            in_string = not in_string
            i += 1
            continue
        if in_string:
            i += 1
            continue
        if ch in pairs:
            stack.append((ch, line, col))
        elif ch in closes:
            if not stack or stack[-1][0] != closes[ch]:
                return [Result("FAIL", label, "balanced brackets/quotes", f"unexpected {ch!r} at {line}:{col}")]
            stack.pop()
        i += 1

    if in_string:
        return [Result("FAIL", label, "balanced brackets/quotes", "unterminated string literal")]
    if stack:
        ch, open_line, open_col = stack[-1]
        return [Result("FAIL", label, "balanced brackets/quotes", f"unclosed {ch!r} from {open_line}:{open_col}")]
    return [Result("PASS", label, "balanced brackets/quotes", "parentheses, square brackets, curly braces, and quotes are balanced")]


def extract_named_formula(text: str, name: str) -> str:
    pattern = rf"(?ms)^{re.escape(name)}\s*=\s*.*?;\s*(?=\n[A-Za-z_][A-Za-z0-9_]*\s*=|\Z)"
    match = re.search(pattern, text)
    return match.group(0).strip() if match else ""


def has_formula(text: str, name: str) -> bool:
    return re.search(rf"(?m)^{re.escape(name)}\s*=", text) is not None


def check_required_regex(results: list[Result], file: str, text: str, check: str, pattern: str, suggestion: str) -> None:
    found = re.search(pattern, text, flags=re.S) is not None
    add(results, found, file, check, "pattern present" if found else "required pattern missing", suggestion)


def audit_public_safety(results: list[Result], files: list[Path]) -> None:
    for path in files:
        label = rel(path)
        add(
            results,
            path.suffix.lower() not in FORBIDDEN_EXTENSIONS,
            label,
            "no workbook or generated binary tracked",
            "extension is allowed",
            "Remove workbook/generated binaries from the public template.",
        )

        if path.suffix.lower() not in {".md", ".txt", ".py", ".formula", ".tsv", ".csv"} and path.name not in {"AGENTS.md", "README.md", "ApplyNotes", "BadgeReportExampleOnly"}:
            continue
        text = read_text(path)
        for needle in FORBIDDEN_TEXT:
            add(
                results,
                needle not in text,
                label,
                f"public safety forbids {needle}",
                "forbidden text absent",
                f"Remove or replace private/internal text: {needle}",
            )
        for pattern, label_name in FORBIDDEN_PATTERNS:
            add(
                results,
                pattern.search(text) is None,
                label,
                f"public safety forbids {label_name}",
                f"{label_name} absent",
                f"Remove or replace public-unsafe {label_name}.",
            )


def audit_formula_files(results: list[Result]) -> None:
    for path in sorted(MODULES.glob("*.formula.txt")):
        results.extend(balance_check(path))
    root_formula_like = ROOT / "ApplyNotes"
    if root_formula_like.exists():
        results.extend(balance_check(root_formula_like))

    for file_name, names in REQUIRED_FORMULAS.items():
        path = ROOT / file_name
        text = read_text(path)
        for name in names:
            add(
                results,
                has_formula(text, name),
                file_name,
                f"defines {name}",
                "formula definition present",
                f"Keep {name} importable from {file_name}.",
            )

    for file_name, name in NAMED_FORMULA_BUDGETS:
        path = ROOT / file_name
        formula = extract_named_formula(read_text(path), name)
        if not formula:
            add(results, False, file_name, f"{name} char budget", "formula not found", f"Define {name}.")
            continue
        add(
            results,
            len(formula) <= MAX_NAMED_FORMULA_CHARS,
            file_name,
            f"{name} char budget",
            f"{len(formula)} chars",
            f"Split {name} into helpers before it exceeds {MAX_NAMED_FORMULA_CHARS} chars.",
        )


def audit_docs(results: list[Result]) -> None:
    readme = read_text(ROOT / "README.md")
    operating = read_text(ROOT / "docs" / "operating_contract.md")
    planning = read_text(ROOT / "docs" / "planning_plugins.md")
    scenarios = read_text(ROOT / "docs" / "scenario_matrix.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")
    release = read_text(ROOT / "docs" / "public_release_checklist.md")
    starter = read_text(ROOT / "docs" / "starter_workbook.md")
    starter_table = read_text(ROOT / "samples" / "planning_table_starter.tsv")
    starter_rows = [line.split("\t") for line in starter_table.splitlines() if line.strip()]

    check_required_regex(
        results,
        "README.md",
        readme,
        "README presents public template",
        r"public template for treating complex Excel workbook logic as source code",
        "Describe the repository as a public formula-module template.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README states no workbook binaries",
        r"does not include any real workbook",
        "Keep the public workbook-binary boundary visible.",
    )
    check_required_regex(
        results,
        "docs/operating_contract.md",
        operating,
        "operating contract states source-code pattern",
        r"formula modules are edited in plain text",
        "Document the formula-module source-control contract.",
    )
    check_required_regex(
        results,
        "docs/planning_plugins.md",
        planning,
        "planning plugins document reforecast pivot",
        r"`Reforecast Queue by <Group>, Job, and Action` is a `PIVOTBY` matrix",
        "Document the reforecast pivot summary.",
    )
    check_required_regex(
        results,
        "docs/scenario_matrix.md",
        scenarios,
        "scenario matrix covers public safety",
        r"Public Safety",
        "Include public safety scenarios.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records public sanitization",
        r"Public template sanitization started",
        "Record the public-template sanitization pass.",
    )
    check_required_regex(
        results,
        "docs/public_release_checklist.md",
        release,
        "release checklist requires clean-history export",
        r"clean-history export",
        "Document that the public repo must be exported with fresh history.",
    )
    check_required_regex(
        results,
        "docs/public_release_checklist.md",
        release,
        "release checklist blocks private workbook material",
        r"no private workbook files",
        "Document the no-private-workbook release gate.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide explains paste target",
        r"Planning Table!A2",
        "Document where to paste the starter table.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide explains monthly triples",
        r"Blank values are acceptable\. Missing columns are not\.",
        "Document why the wide monthly finance block exists.",
    )
    check_required_regex(
        results,
        "samples/planning_table_starter.tsv",
        starter_table,
        "starter table has annual projection",
        r"2026 Projected",
        "Keep the paste-ready Planning Table starter aligned with get.GetFinanceBlock.",
    )
    check_required_regex(
        results,
        "samples/planning_table_starter.tsv",
        starter_table,
        "starter table has monthly triples",
        r"January Projected\tJanuary Actuals\tJanuary.*December Projected\tDecember Actuals\tDecember",
        "Keep projected, actuals, and budget columns for all twelve months.",
    )
    check_required_regex(
        results,
        "samples/planning_table_starter.tsv",
        starter_table,
        "starter table uses fictional BU keys",
        r"BU-A: Sample Unit.*BU-B: Sample Unit",
        "Use fictional sample BU keys instead of real business-unit codes.",
    )
    add(
        results,
        bool(starter_rows) and all(len(row) == 67 for row in starter_rows),
        "samples/planning_table_starter.tsv",
        "starter table row width",
        "all rows have 67 tab-delimited columns" if starter_rows else "starter table is empty",
        "Keep every starter row aligned to the 67-column Planning Table contract.",
    )


def audit_reforecast_contract(results: list[Result]) -> None:
    analysis = read_text(MODULES / "analysis.formula.txt")
    reforecast = extract_named_formula(analysis, "REFORECAST_QUEUE")
    checks = [
        ("composes axis helper", r"Axis,\s*REFORECAST_QUEUE_AXIS\(GroupChoiceRaw\)"),
        ("defines synthetic job key", r"JobKey,\s*CHOOSECOLS\(Rows,\s*5\).*CHOOSECOLS\(Rows,\s*6\).*CHOOSECOLS\(Rows,\s*7\)"),
        ("uses PIVOTBY summary", r"GroupJobActionPivot,\s*PIVOTBY\("),
        ("uses selected group and job row fields", r"VSTACK\(GroupLabel,\s*GroupVals\).*VSTACK\(\"Job\",\s*JobKey\)"),
        ("uses decision dollars value", r"VSTACK\(\"Decision Dollars\",\s*DecisionDollars\)"),
        ("uses row subtotal depth 2", r"SUM,\s*3,\s*2,"),
        ("keeps ranked detail section", r"Reforecast Queue Detail"),
    ]
    for check, pattern in checks:
        check_required_regex(
            results,
            "modules/analysis.formula.txt",
            reforecast,
            f"REFORECAST_QUEUE {check}",
            pattern,
            "Keep the public reforecast queue contract intact.",
        )


def main() -> int:
    results: list[Result] = []
    files = tracked_files()

    audit_public_safety(results, files)
    audit_formula_files(results)
    audit_docs(results)
    audit_reforecast_contract(results)

    for result in results:
        print(result.render())

    fail_count = sum(1 for result in results if result.status == "FAIL")
    warn_count = sum(1 for result in results if result.status == "WARN")
    pass_count = sum(1 for result in results if result.status == "PASS")
    print(f"\nSummary: {pass_count} PASS, {warn_count} WARN, {fail_count} FAIL")
    return 1 if fail_count else 0


if __name__ == "__main__":
    sys.exit(main())
