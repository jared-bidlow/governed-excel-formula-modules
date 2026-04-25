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
    joined("202", "6 Capital Budget v", ".xlsx"),
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
    joined("Capex_", "Tracker_", "2026"),
    joined("CAPEX", "_REPORT"),
    joined("2026 ", "Projected"),
    joined("CapByBU", "_Keys"),
    joined("CapByBU", "_Vals"),
    joined("Cap", "Ex", "Cap"),
]

FORBIDDEN_PATTERNS = [
    (re.compile(r"[A-Za-z]:\\"), "local Windows path"),
    (re.compile(r"https?://", re.I), "URL"),
    (re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I), "email address"),
]

SCAN_TEXT_SUFFIXES = {
    ".css",
    ".csv",
    ".html",
    ".js",
    ".json",
    ".md",
    ".ps1",
    ".py",
    ".svg",
    ".tsv",
    ".txt",
    ".xml",
    ".formula",
}

ALLOWED_ADDIN_URL_PREFIXES = (
    "http://schemas.microsoft.com/office/appforoffice/1.1",
    "http://www.w3.org/2000/svg",
    "http://www.w3.org/2001/XMLSchema-instance",
    "https://localhost:3000",
    "https://appsforoffice.microsoft.com/lib/1/hosted/office.js",
)

PATTERN_DEFINITION_FILES = {
    Path("tools/audit_capex_module.py"),
}

FORBIDDEN_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xls",
    ".zip",
}

REQUIRED_FORMULAS = {
    "modules/controls.formula.txt": [
        "PM_Filter_Dropdowns",
        "Future_Filter_Mode",
        "HideClosed_Status",
        "Burndown_Cut_Target",
    ],
    "modules/get.formula.txt": [
        "TRIMRANGE_KEEPBLANKS",
        "GetFinanceBlock",
        "GetBudgetDetailRows",
        "GetProjections",
        "GetActuals12",
    ],
    "modules/kind.formula.txt": [
        "RBYROW",
        "CapTable",
        "PortfolioCap",
        "CapConsumeMask",
        "CapByBU",
        "FlagsOut",
        "HiddenMaskLogic",
    ],
    "modules/capital_planning_report.formula.txt": [
        "CAPITAL_PLANNING_REPORT",
    ],
    "modules/analysis.formula.txt": [
        "BU_CAP_SCORECARD_AXIS",
        "BU_CAP_SCORECARD",
        "REFORECAST_QUEUE_AXIS",
        "REFORECAST_QUEUE",
    ],
}

NAMED_FORMULA_BUDGETS = [
    ("modules/capital_planning_report.formula.txt", "CAPITAL_PLANNING_REPORT"),
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

        if path.suffix.lower() not in SCAN_TEXT_SUFFIXES and path.name not in {"AGENTS.md", "README.md", "ApplyNotes", "BadgeReportExampleOnly"}:
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
        rel_path = path.relative_to(ROOT)
        if rel_path in PATTERN_DEFINITION_FILES:
            continue
        for pattern, label_name in FORBIDDEN_PATTERNS:
            if label_name == "URL" and "addin" in rel_path.parts:
                urls = re.findall(r"https?://[^\s\"'<>]+", text, flags=re.I)
                add(
                    results,
                    all(url.startswith(ALLOWED_ADDIN_URL_PREFIXES) for url in urls),
                    label,
                    "public safety allows only add-in development URLs",
                    "add-in URLs are allowlisted",
                    "Use only the local add-in development host or Office.js CDN URL in add-in files.",
                )
                continue
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
    review = read_text(ROOT / "docs" / "technical_review_guide.md")
    push_helper = read_text(ROOT / "tools" / "push_public.ps1")
    addin_smoke = read_text(ROOT / "tools" / "start_addin_smoke_test.ps1")
    addin_server = read_text(ROOT / "tools" / "start_addin_dev_server.ps1")
    addin_stop = read_text(ROOT / "tools" / "stop_addin_smoke_test.ps1")
    gitignore = read_text(ROOT / ".gitignore")
    package_json = read_text(ROOT / "package.json")
    starter_table = read_text(ROOT / "samples" / "planning_table_starter.tsv")
    cap_starter = read_text(ROOT / "samples" / "cap_setup_starter.tsv")
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
        "README.md",
        readme,
        "README defends Excel runtime",
        r"Excel is the right runtime for this pattern",
        "Explain why Excel is the runtime for the public template.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README states application boundary",
        r"multi-user transactions, permissions, APIs, durable storage",
        "Keep the public Excel rationale honest about application boundaries.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents add-in smoke helper",
        r"start_addin_smoke_test\.ps1",
        "Tell users the one-command Office.js smoke-test path.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents npm smoke helper",
        r"npm run addin:smoke",
        "Tell users the npm Office.js smoke-test path.",
    )
    for check, pattern in [
        ("defines smoke script", r'"addin:smoke"\s*:\s*"powershell .*start_addin_smoke_test\.ps1"'),
        ("defines dev-server script", r'"dev-server"\s*:\s*"powershell .*start_addin_dev_server\.ps1"'),
        ("declares Office debugging tool", r'"office-addin-debugging"\s*:'),
    ]:
        check_required_regex(
            results,
            "package.json",
            package_json,
            f"package metadata {check}",
            pattern,
            "Keep package metadata aligned with the Office.js smoke-test helpers.",
        )
    check_required_regex(
        results,
        ".gitignore",
        gitignore,
        "gitignore excludes node local tooling",
        r"node_modules/",
        "Keep local Office.js npm tooling out of source control.",
    )
    check_required_regex(
        results,
        ".gitignore",
        gitignore,
        "gitignore excludes npm lockfile",
        r"package-lock\.json",
        "Keep generated npm lock metadata out of this public formula-template repo.",
    )
    check_required_regex(
        results,
        ".gitignore",
        gitignore,
        "gitignore excludes local add-in cert files",
        r"\.office-addin-dev-certs/",
        "Keep generated local certificates out of source control.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README points technical reviewers to guide",
        r"docs/technical_review_guide\.md",
        "Surface the technical-review path from the README.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README explains cap setup",
        r"Cap Setup.*samples/cap_setup_starter\.tsv.*kind\.CapByBU",
        "Tell public users how to set BU caps without editing formula modules.",
    )
    check_required_regex(
        results,
        "docs/technical_review_guide.md",
        review,
        "technical review guide states systems pattern",
        r"treats workbook logic as governed source code",
        "Keep the reviewer guide focused on the governed-workbook systems pattern.",
    )
    check_required_regex(
        results,
        "docs/technical_review_guide.md",
        review,
        "technical review guide lists review path",
        r"Read `README\.md`.*Inspect `tools/audit_capex_module\.py`",
        "Give reviewers a concrete file-by-file path through the repo.",
    )
    check_required_regex(
        results,
        "docs/technical_review_guide.md",
        review,
        "technical review guide states public boundary",
        r"This repo intentionally does not include:.*workbook binaries.*production data",
        "Keep the public/private boundary visible in the reviewer guide.",
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
        "docs/operating_contract.md",
        operating,
        "operating contract states runtime position",
        r"Excel is the runtime because this pattern is meant for planning teams",
        "Document why Excel is the public template runtime.",
    )
    check_required_regex(
        results,
        "docs/operating_contract.md",
        operating,
        "operating contract keeps caps workbook-driven",
        r"BU cap values should be changed in the workbook's `Cap Setup` sheet",
        "Keep cap limits in workbook input tables, not module constants.",
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
        "change log records technical review guide",
        r"Technical review guide added",
        "Record the reviewer-facing documentation layer.",
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
        "docs/change_log.md",
        changelog,
        "change log records cap setup contract",
        r"Workbook-driven cap setup",
        "Record the workbook-driven cap setup change.",
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
        "docs/public_release_checklist.md",
        release,
        "release checklist documents push helper",
        r"tools\\push_public\.ps1",
        "Document the local push helper for the public export repo.",
    )
    for check, pattern in [
        ("runs audit", r"python tools\\audit_capex_module\.py"),
        ("runs formula lint", r"python tools\\lint_formulas\.py modules\\\*\.formula\.txt"),
        ("runs whitespace check", r"git diff --check"),
        ("fetches before push", r"git fetch origin"),
        ("rebases when behind", r"git rebase \$upstream"),
        ("pushes selected branch", r"git push origin \$Branch"),
    ]:
        check_required_regex(
            results,
            "tools/push_public.ps1",
            push_helper,
            f"push helper {check}",
            pattern,
            "Keep the public push helper guarded by validation and remote sync.",
        )
    for file_name, text, checks in [
        (
            "tools/start_addin_smoke_test.ps1",
            addin_smoke,
            [
                ("runs static audit", r"python tools\\audit_capex_module\.py"),
                ("runs formula lint", r"python tools\\lint_formulas\.py modules\\\*\.formula\.txt"),
                ("uses Excel desktop sideload", r"office-addin-debugging start.*--app excel"),
                ("recovers installed Node path", r"Use-InstalledNodePath.*ProgramFiles.*nodejs"),
                ("falls back without npm", r"npm is not on PATH.*sideload addin\\manifest\.xml manually"),
                ("starts server helper", r"start_addin_dev_server\.ps1"),
            ],
        ),
        (
            "tools/start_addin_dev_server.ps1",
            addin_server,
            [
                ("creates trusted local certificate", r"New-SelfSignedCertificate.*Import-Certificate"),
                ("uses local certificate package", r"localhost\.pfx.*localhost\.cer"),
                ("serves repo root with TLS stream", r"TcpListener.*SslStream"),
                ("bounds stalled local requests", r"RequestTimeoutMs.*ReceiveTimeout.*AuthenticateAsServerAsync"),
                ("serves taskpane files", r"addin/taskpane\.html"),
            ],
        ),
        (
            "tools/stop_addin_smoke_test.ps1",
            addin_stop,
            [
                ("stops Office debugging session", r"office-addin-debugging stop"),
                ("stops fallback server by port", r"Get-NetTCPConnection -LocalPort \$Port.*Stop-Process"),
                ("recovers installed Node path", r"Use-InstalledNodePath.*ProgramFiles.*nodejs"),
            ],
        ),
    ]:
        for check, pattern in checks:
            check_required_regex(
                results,
                file_name,
                text,
                f"add-in helper {check}",
                pattern,
                "Keep the Office.js smoke-test helpers usable from a clean checkout.",
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
        "starter guide explains cap setup paste target",
        r"Cap Setup!A2",
        "Document where to paste the cap setup starter table.",
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
        r"Annual Projected",
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
    check_required_regex(
        results,
        "samples/cap_setup_starter.tsv",
        cap_starter,
        "cap setup starter has required columns",
        r"\ABU\tCap\r?\n",
        "Provide a paste-ready BU cap starter table.",
    )
    check_required_regex(
        results,
        "samples/cap_setup_starter.tsv",
        cap_starter,
        "cap setup starter uses fake BU examples",
        r"BU-A\t1200000.*BU-B\t800000",
        "Keep cap starter data fake and generic.",
    )


def audit_cap_setup_contract(results: list[Result]) -> None:
    kind = read_text(MODULES / "kind.formula.txt")
    report = read_text(MODULES / "capital_planning_report.formula.txt")

    check_required_regex(
        results,
        "modules/kind.formula.txt",
        kind,
        "cap setup reads workbook sheet",
        r"CapTable\s*=\s*LAMBDA\(.*'Cap Setup'!\$A\$2:\$B\$100",
        "Read BU caps from the public workbook's Cap Setup sheet.",
    )
    check_required_regex(
        results,
        "modules/kind.formula.txt",
        kind,
        "portfolio cap sums cap table",
        r"PortfolioCap\s*=\s*SUM\(N\(CHOOSECOLS\(CapTable\(\),\s*2\)\)\)",
        "Define the portfolio cap from Cap Setup values.",
    )
    check_required_regex(
        results,
        "modules/kind.formula.txt",
        kind,
        "BU cap lookup uses BU key path",
        r"CapByBU\s*=\s*LAMBDA\(bu,.*TEXTBEFORE\(bu & \"\",\s*\":\".*XLOOKUP\(TRIM\(buKey\),\s*CapBUKeys,\s*CapBUVals",
        "Use the BU code before a colon as the cap-table lookup key.",
    )
    check_required_regex(
        results,
        "modules/kind.formula.txt",
        kind,
        "hidden burn helper guards empty group",
        r"HiddenBurnByBU\s*=\s*LAMBDA\(.*IFERROR\(GROUPBY\(BUCode,\s*HiddenBurnAmtVec,\s*SUM,\s*0,\s*0,\s*1,\s*HiddenMask\),\s*HSTACK\(\"\",\s*0\)\)",
        "Treat an empty hidden-burn grouping as zero so starter subtotals do not spill errors.",
    )
    check_required_regex(
        results,
        "modules/capital_planning_report.formula.txt",
        report,
        "main report guards BU cap lookup",
        r"GrpCap,\s*IF\(.*IsGroupBU,\s*IFERROR\(kind\.CapByBU\(GrpBuKey\),\s*0\)",
        "Keep subtotal cap math from surfacing #VALUE or #N/A when a BU key is missing.",
    )


def audit_addin_contract(results: list[Result]) -> None:
    manifest = read_text(ROOT / "addin" / "manifest.xml")
    taskpane = read_text(ROOT / "addin" / "taskpane.js")
    taskpane_html = read_text(ROOT / "addin" / "taskpane.html")
    addin_doc = read_text(ROOT / "docs" / "office_addin.md")
    readme = read_text(ROOT / "README.md")
    operating = read_text(ROOT / "docs" / "operating_contract.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")

    manifest_checks = [
        ("is task pane app", r"xsi:type=\"TaskPaneApp\""),
        ("uses supported manifest version", r"<Version>1\.0\.0\.0</Version>"),
        ("targets Excel workbook", r"<Host Name=\"Workbook\""),
        ("uses read-write document permission", r"<Permissions>ReadWriteDocument</Permissions>"),
        ("uses PNG icon", r"<IconUrl DefaultValue=\"https://localhost:3000/addin/assets/icon-32\.png\""),
        ("uses PNG high resolution icon", r"<HighResolutionIconUrl DefaultValue=\"https://localhost:3000/addin/assets/icon-64\.png\""),
        ("has local taskpane source", r"<SourceLocation DefaultValue=\"https://localhost:3000/addin/taskpane\.html\""),
    ]
    for check, pattern in manifest_checks:
        check_required_regex(
            results,
            "addin/manifest.xml",
            manifest,
            f"Office.js manifest {check}",
            pattern,
            "Keep the Office.js manifest usable for local sideload testing.",
        )

    taskpane_checks = [
        ("loads Office.js", r"Office\.onReady"),
        ("uses Excel.run", r"Excel\.run"),
        ("creates starter sheets", r"Planning Table.*Cap Setup.*Planning Review"),
        ("loads workbook controls", r"../modules/controls\.formula\.txt"),
        ("loads formula modules", r"../modules/kind\.formula\.txt.*../modules/analysis\.formula\.txt"),
        ("installs workbook names", r"context\.workbook\.names\.add"),
        ("installs qualified module names", r"name:\s*`\$\{moduleFile\.prefix\}\.\$\{item\.name\}`"),
        ("handles unqualified alias collisions", r"unqualifiedAliases"),
        ("validates required names", r"requiredNames"),
        ("validates workbook control names", r"PM_Filter_Dropdowns.*Future_Filter_Mode.*HideClosed_Status.*Burndown_Cut_Target"),
        ("validates workbook-local compatibility helpers", r"TRIMRANGE_KEEPBLANKS.*RBYROW"),
        ("strips module comments", r"stripBlockComments"),
    ]
    for check, pattern in taskpane_checks:
        check_required_regex(
            results,
            "addin/taskpane.js",
            taskpane,
            f"task pane {check}",
            pattern,
            "Keep the add-in as a formula-module installer and validator.",
        )

    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane imports Office.js",
        r"appsforoffice\.microsoft\.com/lib/1/hosted/office\.js",
        "Load the official Office.js host library.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs state installer boundary",
        r"installer and validator.*does not replace the formula modules",
        "Document that JavaScript is not the calculation engine.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README mentions Office.js starter",
        r"Office\.js Add-In Starter",
        "Surface the add-in packaging path in the README.",
    )
    check_required_regex(
        results,
        "docs/operating_contract.md",
        operating,
        "operating contract states add-in boundary",
        r"Office\.js add-in under `addin/` is a packaging and installation layer",
        "Keep the add-in role separate from formula logic.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records manifest validation fix",
        r"Fix Office manifest validation",
        "Record the Office manifest validation fix.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Office.js starter",
        r"Office\.js add-in starter",
        "Record the add-in starter change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records add-in smoke helper",
        r"Automated add-in smoke-test helper",
        "Record the automated add-in smoke-test helper.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records npm smoke metadata",
        r"Add npm smoke-test package metadata",
        "Record the npm smoke-test package metadata.",
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
    audit_cap_setup_contract(results)
    audit_addin_contract(results)
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
