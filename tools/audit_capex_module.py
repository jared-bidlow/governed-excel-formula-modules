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
    ".bas",
    ".csv",
    ".html",
    ".js",
    ".json",
    ".m",
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

ALLOWED_PUBLIC_URLS_BY_PATH = {
    Path("samples/ontology_namespaces_starter.tsv"): (
        "https://w3id.org/rec#",
        "https://brickschema.org/schema/Brick#",
    ),
}

PATTERN_DEFINITION_FILES = {
    Path("tools/audit_capex_module.py"),
}

FORBIDDEN_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xls",
    ".xltx",
    ".xltm",
    ".zip",
}

ANALYSIS_PUBLIC_FORMULAS = [
    "PM_SPEND_REPORT",
    "WORKING_BUDGET_SCREEN",
    "REFORECAST_QUEUE_AXIS",
    "REFORECAST_QUEUE",
    "BU_CAP_SCORECARD_AXIS",
    "BU_CAP_SCORECARD",
    "BURNDOWN_AXIS",
    "BURNDOWN_BASE",
    "BURNDOWN_WS_HORIZON",
    "BURNDOWN_WS_STATUS",
    "BURNDOWN_WS_SIGNAL",
    "BURNDOWN_DIRECTOR_SUMMARY_FROM_AXIS",
    "BURNDOWN_DIRECTOR_SUMMARY",
    "BURNDOWN_SCREEN_TOP_FROM_AXIS",
    "BURNDOWN_SCREEN_DETAIL_FROM_AXIS",
    "BURNDOWN_SCREEN",
]

ASSET_PUBLIC_FORMULAS = [
    "ASSET_REGISTER_START_HERE",
    "ASSET_REGISTER_STATUS",
    "ASSET_REGISTER_ISSUES",
    "ASSET_REGISTER_FIELD_GUIDE",
    "ASSET_START_HERE",
    "ASSET_WORKFLOW_STATUS",
    "ASSET_NEXT_ACTIONS",
    "ASSET_TABLE_MAP",
    "ASSET_GLOSSARY",
    "ASSET_REVIEW_QUEUE",
    "PROJECT_PROMOTION_QUEUE",
    "ASSET_MAPPING_ISSUES",
    "ASSET_CHANGE_ISSUES",
    "INSTALLED_WITHOUT_EVIDENCE",
    "REPLACEMENT_SOURCE_TARGET_ISSUES",
]

ASSET_FINANCE_PUBLIC_FORMULAS = [
    "FINANCE_START_HERE",
    "FINANCE_READINESS_STATUS",
    "CLASSIFIED_MODEL_INPUTS",
    "DEPRECIATION_SCHEDULE",
    "FUNDING_REQUIREMENTS",
    "FINANCE_TOTALS",
    "CHART_FEEDS",
]

SOURCE_PUBLIC_FORMULAS = [
    "SOURCE_STATUS",
    "SOURCE_SCHEMA_STATUS",
    "SOURCE_REFRESH_STATUS",
    "SOURCE_ROW_HEALTH",
    "SOURCE_LINEAGE",
    "SOURCE_RECONCILIATION_QUEUE",
]

ONTOLOGY_PUBLIC_FORMULAS = [
    "ONTOLOGY_START_HERE",
    "CLASS_MAP",
    "RELATIONSHIP_MAP",
    "SEMANTIC_MAPPING_STATUS",
    "ONTOLOGY_ISSUES",
    "TRIPLE_EXPORT_QUEUE",
    "JSONLD_EXPORT_HELP",
]

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
    "modules/analysis.formula.txt": ANALYSIS_PUBLIC_FORMULAS,
    "modules/assets.formula.txt": ASSET_PUBLIC_FORMULAS,
    "modules/asset_finance.formula.txt": ASSET_FINANCE_PUBLIC_FORMULAS,
    "modules/source.formula.txt": SOURCE_PUBLIC_FORMULAS,
    "modules/ontology.formula.txt": ONTOLOGY_PUBLIC_FORMULAS,
    "modules/ready.formula.txt": [
        "ColumnOrBlank",
        "InternalEligible",
        "Maturity",
        "Stage",
        "ChargeableFlag",
        "InternalReady3",
        "InternalJobs_Export",
    ],
}

MODULE_PREFIX_FILES = {
    "Controls": "modules/controls.formula.txt",
    "get": "modules/get.formula.txt",
    "kind": "modules/kind.formula.txt",
    "CapitalPlanning": "modules/capital_planning_report.formula.txt",
    "Analysis": "modules/analysis.formula.txt",
    "defer": "modules/defer.formula.txt",
    "Notes": "modules/notes.formula.txt",
    "Phasing": "modules/phasing.formula.txt",
    "Ready": "modules/ready.formula.txt",
    "Search": "modules/search.formula.txt",
    "Source": "modules/source.formula.txt",
    "Assets": "modules/assets.formula.txt",
    "AssetFinance": "modules/asset_finance.formula.txt",
    "Ontology": "modules/ontology.formula.txt",
}

NAMED_FORMULA_BUDGETS = [
    ("modules/capital_planning_report.formula.txt", "CAPITAL_PLANNING_REPORT"),
] + [("modules/analysis.formula.txt", name) for name in ANALYSIS_PUBLIC_FORMULAS] + [
    ("modules/assets.formula.txt", name) for name in ASSET_PUBLIC_FORMULAS
] + [
    ("modules/asset_finance.formula.txt", name) for name in ASSET_FINANCE_PUBLIC_FORMULAS
] + [
    ("modules/source.formula.txt", name) for name in SOURCE_PUBLIC_FORMULAS
] + [
    ("modules/ontology.formula.txt", name) for name in ONTOLOGY_PUBLIC_FORMULAS
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


def read_tsv_rows(relative_path: str) -> list[list[str]]:
    text = read_text(ROOT / relative_path)
    return [line.split("\t") for line in text.splitlines() if line.strip()]


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


def extract_named_formula_bodies(text: str) -> list[tuple[str, str]]:
    pattern = r"(?ms)^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?);\s*(?=\n[A-Za-z_][A-Za-z0-9_]*\s*=|\Z)"
    return [(match.group(1), match.group(2).strip()) for match in re.finditer(pattern, text)]


def compact_formula_body(text: str) -> str:
    source = strip_block_comments(text)
    out: list[str] = []
    in_string = False
    in_quoted_sheet = False
    i = 0
    while i < len(source):
        ch = source[i]
        if ch == '"' and not in_quoted_sheet:
            out.append(ch)
            if in_string and i + 1 < len(source) and source[i + 1] == '"':
                out.append(source[i + 1])
                i += 2
                continue
            in_string = not in_string
            i += 1
            continue
        if ch == "'" and not in_string:
            out.append(ch)
            if in_quoted_sheet and i + 1 < len(source) and source[i + 1] == "'":
                out.append(source[i + 1])
                i += 2
                continue
            in_quoted_sheet = not in_quoted_sheet
            i += 1
            continue
        if not in_string and not in_quoted_sheet and ch.isspace():
            i += 1
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def has_formula(text: str, name: str) -> bool:
    return re.search(rf"(?m)^{re.escape(name)}\s*=", text) is not None


def strip_formula_literals(text: str) -> str:
    source = strip_block_comments(text)
    out: list[str] = []
    in_string = False
    in_quoted_sheet = False
    i = 0
    while i < len(source):
        ch = source[i]
        if ch == '"' and not in_quoted_sheet:
            if in_string and i + 1 < len(source) and source[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
            i += 1
            continue
        if ch == "'" and not in_string:
            if in_quoted_sheet and i + 1 < len(source) and source[i + 1] == "'":
                i += 2
                continue
            in_quoted_sheet = not in_quoted_sheet
            i += 1
            continue
        out.append(ch if not in_string and not in_quoted_sheet else " ")
        i += 1
    return "".join(out)


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
            if label_name == "URL" and rel_path in ALLOWED_PUBLIC_URLS_BY_PATH:
                urls = re.findall(r"https?://[^\s\"'<>]+", text, flags=re.I)
                allowed_prefixes = ALLOWED_PUBLIC_URLS_BY_PATH[rel_path]
                add(
                    results,
                    all(url.startswith(allowed_prefixes) for url in urls),
                    label,
                    "public safety allows only approved public namespace URLs",
                    "public namespace URLs are allowlisted",
                    "Use only approved public REC/Brick namespace identifiers in ontology starter files.",
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


def audit_qualified_formula_references(results: list[Result]) -> None:
    module_bodies: dict[str, list[tuple[str, str]]] = {}
    definitions: dict[str, set[str]] = {}
    for prefix, file_name in MODULE_PREFIX_FILES.items():
        bodies = extract_named_formula_bodies(read_text(ROOT / file_name))
        module_bodies[file_name] = bodies
        definitions[prefix] = {name for name, _body in bodies}

    prefix_pattern = "|".join(re.escape(prefix) for prefix in sorted(MODULE_PREFIX_FILES, key=len, reverse=True))
    reference_pattern = re.compile(rf"(?<![A-Za-z0-9_])({prefix_pattern})\.([A-Za-z_][A-Za-z0-9_]*)\b")

    for file_name, bodies in sorted(module_bodies.items()):
        reference_count = 0
        missing: list[str] = []
        for source_name, body in bodies:
            for match in reference_pattern.finditer(strip_formula_literals(body)):
                prefix, target = match.groups()
                reference_count += 1
                if target not in definitions[prefix]:
                    missing.append(f"{source_name}: {prefix}.{target}")
        add(
            results,
            not missing,
            file_name,
            "qualified formula references resolve",
            f"{reference_count} qualified references resolved" if not missing else f"missing references: {'; '.join(missing[:5])}",
            "Define the referenced formula in the target module or update the qualified reference.",
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

    audit_qualified_formula_references(results)

    defer = read_text(ROOT / "modules" / "defer.formula.txt")
    defer_audit = extract_named_formula(defer, "Audit")
    add(
        results,
        "get.GetBudgetDetailRows()" in defer_audit and "get.GetBudgetActiveRows()" not in defer_audit,
        "modules/defer.formula.txt",
        "defer.Audit uses defined budget detail helper",
        "defer.Audit points to get.GetBudgetDetailRows()",
        "Keep defer.Audit aligned to the get module's defined budget row helper.",
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

    for path in sorted(MODULES.glob("*.formula.txt")):
        lengths = [
            (name, len("=" + compact_formula_body(body)))
            for name, body in extract_named_formula_bodies(read_text(path))
        ]
        max_name, max_len = max(lengths, key=lambda item: item[1]) if lengths else ("", 0)
        add(
            results,
            max_len <= MAX_NAMED_FORMULA_CHARS,
            rel(path),
            "installed formula bodies fit Excel save limit",
            f"max compacted formula is {max_name} at {max_len} chars",
            "Keep Office.js-installed formulas under Excel's 8192-character save limit after compaction.",
        )

    ready = read_text(ROOT / "modules" / "ready.formula.txt")
    check_required_regex(
        results,
        "modules/ready.formula.txt",
        ready,
        "Ready helper uses header-driven ChargeableFlag",
        r"ChargeableFlag\s*=\s*Ready\.ColumnOrBlank\(\"Chargeable\"\)",
        "Keep Ready chargeability tied to the Chargeable header, not a column letter.",
    )
    check_required_regex(
        results,
        "modules/ready.formula.txt",
        ready,
        "Ready eligibility uses Internal Eligible directly",
        r"InternalEligible\s*=\s*Ready\.ColumnOrBlank\(\"Internal Eligible\"\)",
        "Do not reintroduce an older visible Eligible fallback column.",
    )
    add(
        results,
        'ColumnOrBlank("Eligible")' not in ready and 'HeaderPos("Eligible")' not in ready,
        "modules/ready.formula.txt",
        "Ready module does not fallback to visible Eligible column",
        "legacy fallback absent",
        "Use Internal Eligible as the sole readiness eligibility input.",
    )
    check_required_regex(
        results,
        "modules/ready.formula.txt",
        ready,
        "Ready helper uses chargeability parameter name",
        r"InternalReady3\s*=\s*LAMBDA\(eligible,\s*maturity,\s*stage,\s*chargeableFlag",
        "Use Chargeable wording for the fourth InternalReady3 input.",
    )

    assets = read_text(ROOT / "modules" / "assets.formula.txt")
    check_required_regex(
        results,
        "modules/assets.formula.txt",
        assets,
        "Assets promotion queue uses dropdown-normalized review status",
        r"PROJECT_PROMOTION_QUEUE.*constCol\(\"review\"\).*constCol\(FALSE\)",
        "Keep generated asset promotion queue values aligned with asset validation lists.",
    )

    asset_finance = read_text(ROOT / "modules" / "asset_finance.formula.txt")
    check_required_regex(
        results,
        "modules/asset_finance.formula.txt",
        asset_finance,
        "AssetFinance reads loaded model input bridge",
        r"tblAssetEvidence_ModelInputs.*PresentWithClassifiedEvidence.*AssetFinance\.CLASSIFIED_MODEL_INPUTS",
        "Keep AssetFinance pointed at the loaded qAssetEvidence_ModelInputs table.",
    )
    check_required_regex(
        results,
        "modules/asset_finance.formula.txt",
        asset_finance,
        "AssetFinance exposes finance outputs",
        r"DEPRECIATION_SCHEDULE.*FUNDING_REQUIREMENTS.*FINANCE_TOTALS.*CHART_FEEDS",
        "Keep the v0.4 depreciation, funding, totals, and chart-ready outputs importable.",
    )
    check_required_regex(
        results,
        "modules/asset_finance.formula.txt",
        asset_finance,
        "AssetFinance uses finance assumptions",
        r"tblAssetFinanceAssumptions.*UsefulLifeYears.*FundingRequirementRule.*ChartGroup",
        "Keep asset finance assumptions source-controlled and operator editable.",
    )
    check_required_regex(
        results,
        "modules/asset_finance.formula.txt",
        asset_finance,
        "AssetFinance surfaces unsupported depreciation methods",
        r"DEPRECIATION_SCHEDULE.*DepreciationIssue.*RawMethod.*DepreciationMethod.*Method, IF\(TRIM\(RawMethod & \"\"\) = \"\", \"Straight-line\", RawMethod\).*SupportedMethod.*AnnualDepreciation, IF\(SupportedMethod, Amount / UsefulLife, \"\"\).*Unsupported DepreciationMethod:",
        "Keep unsupported DepreciationMethod values visible with blank depreciation amounts.",
    )
    check_required_regex(
        results,
        "modules/asset_finance.formula.txt",
        asset_finance,
        "AssetFinance surfaces unsupported funding rules",
        r"FUNDING_REQUIREMENTS.*FundingIssue.*RawRules.*FundingRequirementRule.*Rules, IF\(TRIM\(RawRules & \"\"\) = \"\", \"Fund full classified amount\", RawRules\).*SupportedRule.*RequirementAmount, IF\(SupportedRule, GroupedAmount, \"\"\).*Unsupported FundingRequirementRule:",
        "Keep unsupported FundingRequirementRule values visible with blank funding amounts.",
    )
    asset_finance_bodies = dict(extract_named_formula_bodies(asset_finance))
    chart_feeds = asset_finance_bodies.get("CHART_FEEDS", "")
    add(
        results,
        "AssetFinance.FUNDING_REQUIREMENTS" in chart_feeds
        and "AssetFinance.CLASSIFIED_MODEL_INPUTS" not in chart_feeds
        and "CHOOSECOLS(FundRows, 6)" in chart_feeds
        and "AssetFinance.DEPRECIATION_SCHEDULE" in chart_feeds
        and "CHOOSECOLS(DepSourceRows, 9)" in chart_feeds,
        "modules/asset_finance.formula.txt",
        "AssetFinance chart feeds use supported output amounts",
        "chart feeds read supported output amounts",
        "Keep chart feeds aligned to supported depreciation and funding outputs, not raw classified inputs.",
    )
    for forbidden_source in [
        "tblAssetEvidenceSource",
        "tblAssetEvidenceRules",
        "tblAssetEvidenceOverrides",
    ]:
        add(
            results,
            forbidden_source not in asset_finance,
            "modules/asset_finance.formula.txt",
            f"AssetFinance does not read raw setup table {forbidden_source}",
            "raw setup table absent",
            "AssetFinance outputs must read tblAssetEvidence_ModelInputs, not raw setup tables.",
        )
    check_required_regex(
        results,
        "modules/ready.formula.txt",
        ready,
        "Ready helper has self-contained execution stage list",
        r"lstExecStages\s*=\s*\{\"Execution\";\s*\"Demo\";\s*\"Procurement\";\s*\"Install\";\s*\"Commissioning\"\}",
        "Keep Ready independent from private workbook list sheets.",
    )
    add(
        results,
        not has_formula(ready, "JobFlag"),
        "modules/ready.formula.txt",
        "Ready module does not define stale JobFlag helper",
        "stale helper absent",
        "Use Ready.ChargeableFlag for chargeability.",
    )
    add(
        results,
        'ColOrBlank("Internal Ready")' not in ready and "InternalReadyRaw" not in ready,
        "modules/ready.formula.txt",
        "Ready export does not read source-table Internal Ready",
        "source-table override absent",
        "Compute Internal Ready Final in Ready.InternalJobs_Export instead of reading a manual override column.",
    )
    check_required_regex(
        results,
        "modules/ready.formula.txt",
        ready,
        "Ready export emits computed Internal Ready Final",
        r"InternalReadyFinal,\s*InternalReadyComputed.*\"Internal Ready Final\"",
        "Keep readiness output formula-owned and avoid a mixed manual/computed source field.",
    )

    get = read_text(ROOT / "modules" / "get.formula.txt")
    check_required_regex(
        results,
        "modules/get.formula.txt",
        get,
        "get module reads canonical budget input table",
        r"GetBudgetRows\s*=\s*LAMBDA\(tblBudgetInput\[#All\]\).*GetBudgetHeaders\s*=\s*LAMBDA\(TAKE\(GetBudgetRows\(\),\s*1\)\).*GetBudgetBodyRaw\s*=\s*LAMBDA\(.*DROP\(GetBudgetRows\(\),\s*1\).*BYROW",
        "Keep get pointed at tblBudgetInput so formulas consume the canonical import layer.",
    )
    add(
        results,
        "'Planning Table'!$A$2:$BL$2" not in get and "'Planning Table'!$A$3:$BL$234" not in get,
        "modules/get.formula.txt",
        "get module does not read Planning Table ranges directly",
        "direct Planning Table ranges absent",
        "Route source rows through tblBudgetInput instead of fixed Planning Table coordinates.",
    )

    source = read_text(ROOT / "modules" / "source.formula.txt")
    check_required_regex(
        results,
        "modules/source.formula.txt",
        source,
        "Source module reads canonical import trust tables",
        r"tblDataSourceProfile.*tblBudgetImportContract.*tblBudgetInput.*tblBudgetImportStatus.*tblBudgetImportIssues",
        "Keep Source formulas aligned to the v0.5 import trust surfaces.",
    )
    check_required_regex(
        results,
        "modules/source.formula.txt",
        source,
        "Source status uses direct spill-safe status rows",
        r"SOURCE_STATUS = LAMBDA.*VSTACK.*Canonical Table.*Budget Input Rows.*Source Mode.*Last Refresh.*Import Status",
        "Keep Source.SOURCE_STATUS simple enough to spill reliably in the generated starter workbook.",
    )
    check_required_regex(
        results,
        "modules/source.formula.txt",
        source,
        "Source module exposes schema and reconciliation outputs",
        r"SOURCE_SCHEMA_STATUS.*MissingRequired.*ExtraInput.*SOURCE_RECONCILIATION_QUEUE",
        "Keep source review outputs visible in workbook formulas.",
    )

    search = read_text(ROOT / "modules" / "search.formula.txt")
    check_required_regex(
        results,
        "modules/search.formula.txt",
        search,
        "Search helper uses header-driven columns",
        r"HeaderPos.*ColOrBlank.*RowValue.*RowValue\(\"Job ID\"\).*RowValue\(\"Chargeable\"\).*RowValue\(\"Stage\"\).*RowValue\(\"Planning Maturity\"\)",
        "Keep Search resilient to starter-column movement.",
    )
    add(
        results,
        "INDEX(row, 13)" not in search and "INDEX(row, 14)" not in search and "INDEX(row, 54)" not in search,
        "modules/search.formula.txt",
        "Search helper avoids stale row ordinals",
        "stale ordinals absent",
        "Use header lookup for Projects_Health inputs.",
    )


def audit_docs(results: list[Result]) -> None:
    readme = read_text(ROOT / "README.md")
    readme_first = read_text(ROOT / "README_FIRST.md")
    operating = read_text(ROOT / "docs" / "operating_contract.md")
    planning = read_text(ROOT / "docs" / "planning_plugins.md")
    scenarios = read_text(ROOT / "docs" / "scenario_matrix.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")
    release = read_text(ROOT / "docs" / "public_release_checklist.md")
    starter = read_text(ROOT / "docs" / "starter_workbook.md")
    review = read_text(ROOT / "docs" / "technical_review_guide.md")
    notes_workflow = read_text(ROOT / "docs" / "notes_apply_workflow.md")
    database_import = read_text(ROOT / "docs" / "database_import_contract.md")
    optional_adapters = read_text(ROOT / "docs" / "optional_platform_adapters.md")
    copilot_playbook = read_text(ROOT / "docs" / "copilot_review_playbook.md")
    notes_formula = read_text(ROOT / "modules" / "notes.formula.txt")
    asset_workflow = read_text(ROOT / "docs" / "asset_setup_workflow.md")
    asset_evidence_pq = read_text(ROOT / "docs" / "asset_evidence_power_query.md")
    asset_next_steps = read_text(ROOT / "docs" / "asset_tracker_next_steps.md")
    v020_release = read_text(ROOT / "docs" / "v0.2.0_release_notes.md")
    durable_contract = read_text(ROOT / "docs" / "codex_chatgpt_durable_contract.md")
    import_map = read_text(ROOT / "docs" / "workbook_import_map.md")
    structure_map = read_text(ROOT / "docs" / "planning_worksheet_structure_map.md")
    workbook_map = read_text(ROOT / "docs" / "workbook_left_to_right_map.md")
    worktree_doc = read_text(ROOT / "docs" / "git_worktree_workflow.md")
    office_scripts_readme = read_text(ROOT / "office-scripts" / "README.md")
    apply_notes_script = read_text(ROOT / "office-scripts" / "apply_notes.ts")
    apply_assets_script = read_text(ROOT / "office-scripts" / "apply_asset_mappings.ts")
    asset_evidence_seed_builder = read_text(ROOT / "tools" / "build_asset_evidence_pq_seed.ps1")
    asset_evidence_workbook_installer = read_text(ROOT / "tools" / "install_asset_evidence_pq_workbook.ps1")
    asset_evidence_button_launcher = read_text(ROOT / "tools" / "start_asset_evidence_pq_installer.ps1")
    push_helper = read_text(ROOT / "tools" / "push_public.ps1")
    worktree_helper = read_text(ROOT / "tools" / "new_worktree.ps1")
    start_addin = read_text(ROOT / "Start-AddIn.ps1")
    addin_smoke = read_text(ROOT / "tools" / "start_addin_smoke_test.ps1")
    addin_server = read_text(ROOT / "tools" / "start_addin_dev_server.ps1")
    addin_stop = read_text(ROOT / "tools" / "stop_addin_smoke_test.ps1")
    gitignore = read_text(ROOT / ".gitignore")
    package_json = read_text(ROOT / "package.json")
    starter_table = read_text(ROOT / "samples" / "planning_table_starter.tsv")
    cap_starter = read_text(ROOT / "samples" / "cap_setup_starter.tsv")
    budget_contract_starter = read_text(ROOT / "samples" / "budget_import_contract_starter.tsv")
    copilot_cards = read_text(ROOT / "samples" / "copilot_prompt_cards.tsv")
    decision_starter = read_text(ROOT / "samples" / "decision_staging_starter.tsv")
    asset_setup_starter = read_text(ROOT / "samples" / "asset_setup_starter.tsv")
    semantic_assets_starter = read_text(ROOT / "samples" / "semantic_assets_starter.tsv")
    project_asset_map_starter = read_text(ROOT / "samples" / "project_asset_map_starter.tsv")
    asset_changes_starter = read_text(ROOT / "samples" / "asset_changes_starter.tsv")
    asset_state_history_starter = read_text(ROOT / "samples" / "asset_state_history_starter.tsv")
    starter_rows = [line.split("\t") for line in starter_table.splitlines() if line.strip()]
    budget_contract_rows = read_tsv_rows("samples/budget_import_contract_starter.tsv")
    budget_contract_columns = [row[0] for row in budget_contract_rows[1:]]
    planning_columns = starter_rows[0] if starter_rows else []

    add(
        results,
        len(budget_contract_columns) == 64,
        "samples/budget_import_contract_starter.tsv",
        "budget import contract preserves 64-column wide planning contract",
        f"{len(budget_contract_columns)} contract columns",
        "Keep tblBudgetInput aligned to the existing 64-column planning table schema.",
    )
    add(
        results,
        budget_contract_columns == planning_columns,
        "samples/budget_import_contract_starter.tsv",
        "budget import contract matches planning starter headers",
        "contract headers match planning starter" if budget_contract_columns == planning_columns else "contract headers drift from planning starter",
        "Keep tblBudgetImportContract column order identical to samples/planning_table_starter.tsv.",
    )

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
    add(
        results,
        not re.search(
            r"\b(Dataverse|Fabric|Power Platform|Azure Digital Twins|SemanticTwin)\b|digital twin readiness",
            readme,
            re.IGNORECASE,
        ),
        "README.md",
        "README has no platform-roadmap terms",
        "Dataverse, Fabric, Power Platform, Azure Digital Twins, SemanticTwin, and digital twin readiness are absent from README",
        "Keep the README focused on the current Excel workbook, CSV handoff, and validation workflow.",
    )
    public_doc_bundle = "\n".join(
        [
            readme,
            readme_first,
            starter,
            database_import,
            asset_workflow,
            asset_evidence_pq,
            import_map,
            structure_map,
            workbook_map,
        ]
    )
    add(
        results,
        not re.search(
            r"platform path|recommended path|later workflow path|future implementation|digital-twin readiness|enterprise data platform|Fabric-ready|Dataverse workflow|Power Platform path",
            public_doc_bundle,
            re.IGNORECASE,
        ),
        "docs",
        "public docs avoid roadmap language",
        "speculative roadmap phrases absent",
        "Use current workflow, operator package, CSV handoff, workbook input, review queue, approved output, and placeholder adapter language.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents operator launcher",
        r"Start-AddIn\.ps1.*workbook copy.*npm dependencies.*launches Excel.*README_FIRST\.md",
        "Surface the safer operator launcher before the developer smoke command.",
    )
    check_required_regex(
        results,
        "README_FIRST.md",
        readme_first,
        "README first gives minimum operator path",
        r"Right-click `Start-AddIn\.ps1`.*Run with PowerShell.*workbook copy.*Setup \+ Install \+ Validate \+ Outputs.*Copy ApplyNotes Script.*Automate -> New Script",
        "Keep the first-read operator path short and concrete.",
    )
    check_required_regex(
        results,
        "README_FIRST.md",
        readme_first,
        "README first states production workbook safety rule",
        r"Do not click setup or apply buttons in a production workbook.*Use a workbook copy",
        "Keep the launcher safety boundary visible to non-developer users.",
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
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents worktree workflow",
        r"Worktree Workflow.*pristine public template state.*formula/add-in tasks.*workbook-contract review.*automated smoke/lint runs.*workbook-reference analysis.*new_worktree\.ps1.*ready-fix.*docs/git_worktree_workflow\.md",
        "Surface the starter Git worktree workflow from the README.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents v0.2.0 notes and asset workflows",
        r"Notes And Asset Workflows.*Setup Notes Workflow.*tblDecisionStaging.*office-scripts/apply_notes\.ts.*Setup Asset Workflow.*modules/assets\.formula\.txt.*docs/notes_apply_workflow\.md.*docs/asset_setup_workflow\.md.*office-scripts/README\.md",
        "Surface the controlled notes apply and optional asset setup workflows from the README.",
    )
    for check, pattern in [
        ("defines operator start script", r'"start:addin"\s*:\s*"powershell .*Start-AddIn\.ps1"'),
        ("defines operator Excel alias", r'"excel:addin"\s*:\s*"powershell .*Start-AddIn\.ps1"'),
        ("defines test smoke alias", r'"test:smoke"\s*:\s*"powershell .*start_addin_smoke_test\.ps1"'),
        ("defines smoke script", r'"addin:smoke"\s*:\s*"powershell .*start_addin_smoke_test\.ps1"'),
        ("defines dev-server script", r'"dev-server"\s*:\s*"powershell .*start_addin_dev_server\.ps1"'),
        ("defines asset evidence PQ seed builder", r'"build:asset-evidence-pq-seed"\s*:\s*"powershell .*build_asset_evidence_pq_seed\.ps1"'),
        ("defines asset evidence PQ workbook installer", r'"install:asset-evidence-pq"\s*:\s*"powershell .*install_asset_evidence_pq_workbook\.ps1"'),
        ("defines asset evidence PQ button launcher", r'"asset-evidence:pq"\s*:\s*"powershell .*start_asset_evidence_pq_installer\.ps1"'),
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
        ".gitignore",
        gitignore,
        "gitignore excludes generated release artifacts",
        r"release_artifacts/",
        "Keep generated seed workbooks out of source control.",
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
        "README.md",
        readme,
        "README documents v0.5 data import bridge",
        r"Data Import Setup.*PQ Budget Input.*PQ Budget QA.*tblBudgetInput.*Planning Table.*remains the manual starter source.*docs/database_import_contract\.md",
        "Surface the canonical budget import layer from the README.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import,
        "database import contract documents canonical tables",
        r"tblDataSourceProfile.*tblBudgetImportParameters.*tblBudgetImportContract.*tblBudgetInput.*tblBudgetImportStatus.*tblBudgetImportIssues",
        "Document every v0.5 canonical import table.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import,
        "database import contract documents 64-column bridge",
        r"tblBudgetInput.*64-column.*Planning Table.*manual/starter.*tblBudgetInput\[#All\]",
        "Keep docs clear that formulas read the canonical table while Planning Table remains a starter source.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import,
        "database import contract lists Power Query templates",
        r"qBudget_Source_CurrentWorkbook.*qBudget_Source_AzureSql.*qBudget_Source_Dataverse.*qBudget_Source_FabricSqlEndpoint.*qBudget_Source_Selected.*qBudget_Normalized.*qBudget_WideContract.*qBudget_Input.*qBudget_Status.*qBudget_Issues",
        "Document the complete budget-input M template set.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import,
        "database import contract documents selected adapter",
        r"qBudget_Source_Selected.*tblBudgetImportParameters.*ActiveAdapter.*CurrentWorkbook.*AzureSQL.*Dataverse.*FabricSqlEndpoint.*qBudget_Normalized.*qBudget_Issues.*qBudget_Status",
        "Document the active-adapter selector path for the budget input bridge.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import,
        "database import contract states public-safe source profile rules",
        r"Do not commit real server names.*tenant names.*workspace names.*connection strings.*credentials.*tokens.*private URLs.*local workbook paths",
        "Keep public-safe adapter rules explicit.",
    )
    check_required_regex(
        results,
        "docs/optional_platform_adapters.md",
        optional_adapters,
        "optional adapter doc stays placeholder-only",
        r"Excel-first.*tblBudgetInput.*placeholder adapters.*not part of the current operator package.*not a recommended next step.*do not create.*direct database writeback",
        "Keep non-current-workbook adapters documented as placeholders, not as a roadmap.",
    )
    check_required_regex(
        results,
        "docs/copilot_review_playbook.md",
        copilot_playbook,
        "Copilot playbook includes deterministic calculation warning",
        r"Copilot may explain.*must not be the source of governed numeric totals.*Use native Excel formulas for numerical tasks requiring accuracy or reproducibility",
        "Keep Copilot positioned as a review aid, not a calculation engine.",
    )
    check_required_regex(
        results,
        "samples/copilot_prompt_cards.tsv",
        copilot_cards,
        "Copilot prompt cards include deterministic calculation guardrail",
        r"Use Copilot for explanation only; governed numeric calculations stay in Excel formulas.*Use native Excel formulas for totals that require accuracy or reproducibility",
        "Keep prompt cards aligned with deterministic calculation rules.",
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
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents Planning Review note columns",
        r"Planning Review.*`O`.*ExistingMeetingNotes.*`P`.*NewPlanningNotes.*`Q`.*NewTimeline.*`R`.*NewStatus",
        "Document the visible notes columns beside the report.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents visible ApplyNotes control area",
        r"ApplyNotes Control.*Planning Review!O1:R3.*Planning Review!P:R.*run `ApplyNotes` once.*Decision Staging.*run `ApplyNotes` again",
        "Document the in-workbook cue that tells operators to run ApplyNotes twice.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents live ApplyNotes control status",
        r"`ApplyNotes` updates the same `Planning Review!O1:R3` control area.*current phase.*last-run timestamp.*result counts.*next action",
        "Document that ApplyNotes updates the in-workbook control area after it runs.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents Decision Staging table",
        r"Decision Staging.*tblDecisionStaging",
        "Document the controlled staging table for ApplyNotes.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents two-pass ApplyNotes behavior",
        r"Run 1 is prepare.*Prepared.*BudgetRowFound.*Run 2 is apply",
        "Document the prepare/apply behavior of office-scripts/apply_notes.ts.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents ApplyNotes statuses and reset",
        r"Status meanings:.*Prepared.*Applied.*Blocked.*Skipped.*Error.*resets stale non-prepared staging rows",
        "Document operator-visible ApplyNotes statuses and reset behavior.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents writeback targets",
        r"Planning Notes.*Timeline.*Comments.*Status",
        "Document the controlled Planning Table writeback targets.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow states no manual copy paste",
        r"without manually copying values|without manual copy/paste",
        "Keep the notes setup described as workbook-usable from the add-in setup.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents ReviewRow staging identity",
        r"records the source worksheet row in `ReviewRow`.*formulas keyed by `ReviewRow`.*blocks duplicate staged rows",
        "Document that multi-row staging is keyed by source Planning Review row and duplicate targets are blocked.",
    )
    check_required_regex(
        results,
        "docs/notes_apply_workflow.md",
        notes_workflow,
        "notes workflow documents Planning Review to script staging",
        r"Planning Review!P5:R5.*ApplyNotes.*tblDecisionStaging",
        "Document that ApplyNotes smoke input originates on Planning Review and is staged by the script.",
    )
    check_required_regex(
        results,
        "modules/notes.formula.txt",
        notes_formula,
        "Notes.FromArrayv carries Planning Review O:R into staging source",
        r"reviewRows, SEQUENCE\(n, 1, 5, 1\).*\"ReviewRow\".*meetHdrs, 'Planning Review'!\$o\$4:\$r\$4.*allHdrs, SUBSTITUTE\(meetHdrs, \" \", \"\"\).*allData, TAKE\(IF\(meetData = \"\", \"\", meetData\), n\).*HSTACK\(baseHdrs, allHdrs\).*FILTER\(HSTACK\(baseData, allData\), keep",
        "Keep the staging source formula connected to Planning Review O:R and carrying ReviewRow identity.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents optional setup",
        r"optional.*not part of the default `Setup \+ Install \+ Validate \+ Outputs` path",
        "Keep asset setup opt-in.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents created sheets",
        r"Asset Register.*Asset Setup.*Project Asset Map.*Semantic Assets.*Asset Changes.*Asset State History",
        "Document the optional asset setup sheets.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents created tables",
        r"tblAssets.*tblSemanticAssets.*tblAssetPromotionQueue.*tblAssetMappingStaging.*tblProjectAssetMap.*tblAssetChanges.*tblAssetStateHistory",
        "Document the optional asset setup tables.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents asset table map",
        r"Asset Table Map.*tblAssets.*durable asset records.*tblProjectAssetMap.*relationship table.*tblAssetStateHistory.*event trail",
        "Document what each asset table owns.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents relationship dropdowns",
        r"Dropdowns And Relationships.*Asset ID.*Project Key.*advisory dropdowns.*allow new IDs",
        "Document the dropdown-backed relationship contract.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents setup reset boundary",
        r"Rerunning this setup recreates the asset workflow tables.*starter/reset action",
        "Make the destructive reset behavior visible before workbook use.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents review-only formulas and controlled writes",
        r"dynamic-array review formulas.*do not write rows.*office-scripts/apply_asset_mappings\.ts.*controlled-write action",
        "Keep asset formulas review-only and Office Scripts as the write layer.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow documents asset evidence Power Query assistant",
        r"Asset Evidence Power Query.*seed-workbook.*tblAssetEvidenceSource.*tblAssetEvidenceRules.*tblAssetEvidenceOverrides.*qAssetEvidence_Normalized.*qQA_AssetEvidence_MappingQueue.*PresentWithClassifiedEvidence.*classifier metadata",
        "Document the optional asset evidence Power Query assistant and evidence distinction.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow does not own asset evidence PQ tables",
        r"`Setup Asset Workflow` does not create `tblAssetEvidenceSource`.*`tblAssetEvidenceRules`.*`tblAssetEvidenceOverrides`.*asset register, mapping, change, and state-history workflow tables only",
        "Keep the task-pane asset workflow separate from the Power Query evidence surface.",
    )
    check_required_regex(
        results,
        "docs/asset_setup_workflow.md",
        asset_workflow,
        "asset workflow still defers export and finished reports",
        r"does not include external graph export.*validation against ontology files.*finished asset reports.*Power Query seed provides setup tables",
        "Keep graph export and finished reports out of this asset setup slice.",
    )
    check_required_regex(
        results,
        "docs/asset_evidence_power_query.md",
        asset_evidence_pq,
        "asset evidence Power Query doc lists operator flow",
        r"Operator Flow.*build_asset_evidence_pq_seed\.ps1.*install_asset_evidence_pq_workbook\.ps1.*\.asset-evidence-pq\.xlsx.*Refresh Power Query",
        "Document the generated seed workbook workflow.",
    )
    check_required_regex(
        results,
        "docs/asset_evidence_power_query.md",
        asset_evidence_pq,
        "asset evidence Power Query doc lists button launcher",
        r"start_asset_evidence_pq_installer\.ps1.*browse for a workbook copy.*Install Asset Evidence PQ",
        "Document the local button-driven installer.",
    )
    check_required_regex(
        results,
        "docs/asset_evidence_power_query.md",
        asset_evidence_pq,
        "asset evidence Power Query doc lists setup tables and outputs",
        r"tblAssetEvidenceSource.*tblAssetEvidenceRules.*tblAssetEvidenceOverrides.*qAssetEvidence_Normalized.*qAssetEvidence_Classified.*qAssetEvidence_Linked.*qAssetEvidence_Status.*qAssetEvidence_ModelInputs.*qQA_AssetEvidence_MappingQueue",
        "Document setup tables and expected query outputs.",
    )
    check_required_regex(
        results,
        "docs/asset_evidence_power_query.md",
        asset_evidence_pq,
        "asset evidence Power Query doc preserves mapped/classified distinction",
        r"PresentWithMappedEvidence.*ContextCategoryId.*ContextCategoryName.*AssetId.*ProjectKey.*PresentWithClassifiedEvidence.*classified category.*classifier metadata.*Structural hints by themselves are not true classified evidence",
        "Keep structural mapping separate from true classified evidence.",
    )
    check_required_regex(
        results,
        "docs/asset_evidence_power_query.md",
        asset_evidence_pq,
        "asset evidence Power Query doc states asset workflow boundary",
        r"`Setup Asset Workflow` is not a prerequisite.*Power Query should not load into those add-in-created workflow tables",
        "Keep the Power Query evidence bridge from depending on add-in-created asset workflow tables.",
    )
    check_required_regex(
        results,
        "tools/build_asset_evidence_pq_seed.ps1",
        asset_evidence_seed_builder,
        "asset evidence seed builder creates loaded query sheets",
        r"(?=.*Asset_Evidence_PQ_Seed\.xlsx)(?=.*samples\\power-query\\asset-evidence)(?=.*tblAssetEvidenceSource)(?=.*qAssetEvidence_Normalized)(?=.*qQA_AssetEvidence_MappingQueue)(?=.*Add-LoadedQueryTable)(?=.*Queries\.Add)",
        "Keep the seed workbook reproducible from source-controlled M templates.",
    )
    check_required_regex(
        results,
        "tools/install_asset_evidence_pq_workbook.ps1",
        asset_evidence_workbook_installer,
        "asset evidence workbook installer writes output copy",
        r"(?=.*TargetWorkbookPath)(?=.*OutputPath)(?=.*OutputPath must be a workbook copy)(?=.*Copy-Item)(?=.*ReplaceExisting)(?=.*Add-LoadedQueryTable)(?=.*Queries\.Add)",
        "Keep the PowerShell installer focused on installing seed-owned sheets into a target workbook copy.",
    )
    check_required_regex(
        results,
        "tools/start_asset_evidence_pq_installer.ps1",
        asset_evidence_button_launcher,
        "asset evidence button launcher wraps build and install scripts",
        r"(?=.*System\.Windows\.Forms)(?=.*Browse)(?=.*Build Seed)(?=.*Install Asset Evidence PQ)(?=.*build_asset_evidence_pq_seed\.ps1)(?=.*install_asset_evidence_pq_workbook\.ps1)",
        "Keep a local button-driven path for users who should not run raw PowerShell commands.",
    )
    check_required_regex(
        results,
        "docs/asset_tracker_next_steps.md",
        asset_next_steps,
        "asset tracker next steps preserve branch boundary",
        r"reference implementation.*stay separate from the active Capital Planning workbook logic",
        "Keep asset-tracker follow-up isolated from the active workbook logic.",
    )
    check_required_regex(
        results,
        "docs/asset_tracker_next_steps.md",
        asset_next_steps,
        "asset tracker next steps map table ownership",
        r"Table Ownership.*tblAssets.*Durable asset records.*tblProjectAssetMap.*Current project-to-asset relationships",
        "Keep a concise table map for the asset tracker path.",
    )
    check_required_regex(
        results,
        "docs/asset_tracker_next_steps.md",
        asset_next_steps,
        "asset tracker next steps list verification path",
        r"Immediate Verification.*workbook copy.*Setup Asset Workflow.*Asset Register.*dropdowns",
        "Keep the next operator validation steps concrete.",
    )
    check_required_regex(
        results,
        "docs/v0.2.0_release_notes.md",
        v020_release,
        "v0.2.0 release notes name release",
        r"v0\.2\.0-notes-apply-asset-setup",
        "Name the release artifact explicitly.",
    )
    check_required_regex(
        results,
        "docs/v0.2.0_release_notes.md",
        v020_release,
        "v0.2.0 release notes summarize additions",
        r"Setup Notes Workflow.*apply_notes\.ts.*apply_asset_mappings\.ts.*Setup Asset Workflow.*starter TSVs.*modules/assets\.formula\.txt",
        "Summarize the notes/apply and optional asset setup additions.",
    )
    check_required_regex(
        results,
        "docs/v0.2.0_release_notes.md",
        v020_release,
        "v0.2.0 release notes defer export bridge work",
        r"external graph export.*external validation.*Power Query bridge",
        "Explicitly defer external graph export, external validation, and Power Query bridge work.",
    )
    check_required_regex(
        results,
        "docs/codex_chatgpt_durable_contract.md",
        durable_contract,
        "durable contract records release handoff format",
        r"Required Release-Handoff Report.*Feature status:.*Built:.*Scaffolded:.*Missing:.*Validation:.*Known limitations:",
        "Keep the handoff note current for release readiness and reviewer packet work.",
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
        "docs/operating_contract.md",
        operating,
        "operating contract states validation-list sheet",
        r"`Validation Lists` contains dropdown source values",
        "Keep the starter workbook layout contract visible.",
    )
    check_required_regex(
        results,
        "docs/operating_contract.md",
        operating,
        "operating contract states visible controls",
        r"visible controls on `Planning Review`.*unqualified control names",
        "Document that public starter controls are worksheet-visible.",
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
        "docs/planning_plugins.md",
        planning,
        "planning plugins document PM spend report",
        r"PM Spend Report.*`Analysis\.PM_SPEND_REPORT\(\[groupBy\]\)`",
        "Document the PM spend report in the public plugin menu.",
    )
    check_required_regex(
        results,
        "docs/planning_plugins.md",
        planning,
        "planning plugins document working budget screen",
        r"Working Budget Screen.*`Analysis\.WORKING_BUDGET_SCREEN\(\)`",
        "Document the working budget screen in the public plugin menu.",
    )
    check_required_regex(
        results,
        "docs/planning_plugins.md",
        planning,
        "planning plugins document burndown screen",
        r"Burndown Screen.*`Analysis\.BURNDOWN_SCREEN\(\[groupBy\]\)`",
        "Document the burndown screen in the public plugin menu.",
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
        "docs/scenario_matrix.md",
        scenarios,
        "scenario matrix covers implemented planning screens",
        r"PM Spend Report.*Working Budget Screen.*Burndown Screen",
        "Keep scenario coverage aligned with implemented Analysis screens.",
    )
    check_required_regex(
        results,
        "docs/scenario_matrix.md",
        scenarios,
        "scenario matrix covers dropdown application data",
        r"Dropdown Application Data.*Planning Review!B2:D2.*Chargeable.*Y,N",
        "Cover model-driven dropdown behavior in the scenario matrix.",
    )
    check_required_regex(
        results,
        "docs/scenario_matrix.md",
        scenarios,
        "scenario matrix covers ready row flags",
        r"Ready And Row Flags.*Ready\.ChargeableFlag.*Chargeable.*Ready\.InternalEligible.*Internal Eligible.*no separate visible `Eligible` fallback column.*No `JobFlag` column and no source-table `Internal Ready` column are present.*Ready\.InternalJobs_Export.*computed `Internal Ready Final`.*Data > Subtotal.*`Composite Cat`",
        "Cover the no-JobFlag, no-source-Internal-Ready, and manual Composite Cat boundaries.",
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
        "docs/change_log.md",
        changelog,
        "change log records planning screen inventory",
        r"Document implemented planning-screen inventory",
        "Record the planning-screen documentation and audit inventory update.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records starter UX setup",
        r"Add starter workbook UX setup",
        "Record visible workbook controls, dropdowns, formatting, and validation.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records stale server guard",
        r"Guard stale add-in dev server reuse",
        "Record the smoke-test stale-server guard.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records validation summary",
        r"Add task-pane validation summary",
        "Record the task-pane validation summary.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records demo output insertion",
        r"Add demo output insertion action",
        "Record the optional demo output insertion action.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records spill-safe control layout",
        r"Move visible controls above report spill",
        "Record the spill-safe Planning Review control layout.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records main report spill guard",
        r"Guard main report demo spill range",
        "Record the demo-output spill guard.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records dropdown application data",
        r"Centralize dropdown application data",
        "Record the model-driven dropdown setup change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Chargeable JobFlag contract",
        r"Clarify Chargeable and JobFlag readiness contract",
        "Record the Chargeable versus JobFlag documentation guardrail.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Ready chargeability helper",
        r"Stabilize Ready chargeability helper",
        "Record the Ready formula snapshot stabilization.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records JobFlag removal",
        r"Remove JobFlag starter column",
        "Record the starter contract width change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Internal Jobs demo sheet",
        r"Add Internal Jobs demo sheet",
        "Record the Ready demo-output sheet.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records combined setup outputs action",
        r"Fold demo outputs into setup action",
        "Record the primary task-pane action creating demo outputs.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Yes/No dependency map",
        r"Document Yes/No planning worksheet dependencies",
        "Record the public-safe structure map update.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Eligible fallback removal",
        r"Remove Eligible fallback column",
        "Record the starter contract simplification.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records computed readiness output",
        r"Keep Composite Cat manual and compute readiness output",
        "Record the Composite Cat and Internal Ready source-column decision.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records full Planning Table structure map",
        r"Expand Planning Table structure map",
        "Record the full-column structure map update.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Excel worktree workflow",
        r"Tailor worktree roles to Excel workflow",
        "Record the Excel-specific Git worktree workflow.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records worktree concurrency roles",
        r"Clarify Git worktree concurrency roles",
        "Record the named Git worktree concurrency roles.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records worktree workflow starter",
        r"Add Git worktree workflow starter",
        "Record the Git worktree workflow starter.",
    )
    for check, pattern in [
        ("states stable main and task worktrees", r"`main` as the stable product branch.*`codex/<task>`"),
        ("documents concurrency role model", r"`main`.*`work`.*`review`.*`fuzz`.*`scratch`"),
        ("maps work role to formula and add-in work", r"`work`.*formula module edits.*Office\.js setup/validation"),
        ("maps review role to workbook contract review", r"`review`.*workbook contract docs"),
        ("assigns fuzz to automated checks", r"`fuzz`.*automated checks"),
        ("assigns scratch to workbook reference analysis", r"`scratch`.*uploaded workbook"),
        ("distinguishes worktrees from branches", r"not a replacement for branches"),
        ("creates feature worktree with helper", r"new_worktree\.ps1 -Name ready-fix"),
        ("documents role-specific prefixes", r"-BranchPrefix review.*-BranchPrefix scratch"),
        ("checks fast-forward divergence", r"git rev-list --left-right --count origin/main\.\.\.origin/codex/ready-fix"),
        ("removes and prunes worktrees", r"git worktree remove.*git worktree prune"),
        ("keeps workbook binaries out of Git", r"Keep workbook binaries out of Git"),
        ("keeps workbook copies local", r"Treat workbook copies as local operator artifacts"),
        ("promotes scratch only as sanitized text", r"Promote scratch findings only by writing sanitized docs, TSV samples, formula modules, add-in source, or audit checks"),
        ("keeps public safety explicit", r"Do not use linked worktrees to hide private data"),
    ]:
        check_required_regex(
            results,
            "docs/git_worktree_workflow.md",
            worktree_doc,
            f"worktree workflow {check}",
            pattern,
            "Keep the Git worktree workflow aligned with the public-template branch discipline.",
        )
    for check, pattern in [
        ("defaults branch prefix to codex", r'\$BranchPrefix = "codex"'),
        ("defaults base branch to main", r'\$BaseBranch = "main"'),
        ("fetches origin before creating worktree", r"git fetch origin"),
        ("verifies origin base ref", r'git rev-parse --verify \$baseRef'),
        ("blocks existing target path", r"Test-Path \$targetPath.*Target path already exists"),
        ("creates branch-backed worktree", r"git worktree add -b \$branchName \$targetPath \$baseRef"),
    ]:
        check_required_regex(
            results,
            "tools/new_worktree.ps1",
            worktree_helper,
            f"worktree helper {check}",
            pattern,
            "Keep the worktree helper conservative and based on origin/main by default.",
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
            "Start-AddIn.ps1",
            start_addin,
            [
                ("requires workbook-copy confirmation", r"Use a workbook copy.*Read-Host.*Confirm you will use a workbook copy"),
                ("supports noninteractive confirmation", r"\[switch\]\$Yes.*-not \$Yes"),
                ("recovers installed Node path", r"Use-InstalledNodePath.*ProgramFiles.*nodejs"),
                ("installs dependencies when missing", r"node_modules.*npm install"),
                ("delegates to smoke helper", r"start_addin_smoke_test\.ps1.*-Port \$Port.*-SkipStaticChecks:\$SkipStaticChecks"),
                ("warns when npm is missing", r"npm was not found.*Install Node\.js LTS"),
                ("describes workbook-safe behavior", r"It does not edit a workbook by itself"),
            ],
        ),
        (
            "tools/start_addin_smoke_test.ps1",
            addin_smoke,
            [
                ("runs static audit", r"python tools\\audit_capex_module\.py"),
                ("runs formula lint", r"python tools\\lint_formulas\.py modules\\\*\.formula\.txt"),
                ("uses Excel desktop sideload", r"office-addin-debugging start.*--app excel"),
                ("recovers installed Node path", r"Use-InstalledNodePath.*ProgramFiles.*nodejs"),
                ("guards stale dev server reuse", r"Stop-StaleDevServer.*probeUrl.*taskpane\.js.*Invoke-WebRequest.*Stop-Process"),
                ("falls back without npm", r"npm is not on PATH.*sideload addin\\manifest\.xml manually"),
                ("starts server helper", r"start_addin_dev_server\.ps1"),
                ("documents demo output smoke step", r"Setup \+ Install \+ Validate \+ Outputs"),
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
        "docs/starter_workbook.md",
        starter,
        "starter guide documents validation lists",
        r"`Validation Lists`.*Dropdown source values",
        "Document the dropdown-source sheet in the starter workbook guide.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents visible controls",
        r"`B2`.*`PM_Filter_Dropdowns`.*`E2`.*Burndown_Cut_Target",
        "Document the visible Planning Review control panel.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide preserves output ranges",
        r"`Planning Review!A4:N200`.*`Planning Review!O1:R3`.*`Planning Review!O4:R200`.*do not block the report spill",
        "Keep spill, control, and notes ranges reserved in the starter guide.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents demo outputs",
        r"Insert Demo Outputs.*Analysis Hub.*BU Cap Scorecard.*Reforecast Queue.*PM Spend Report.*Working Budget.*Burndown.*Internal Jobs.*Ready\.InternalJobs_Export",
        "Document the optional demo output hubs.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents demo spill guard",
        r"Insert Demo Outputs.*Planning Review!A4:N200.*block the spill",
        "Document the pre-insert main report spill guard.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents application data dropdown contract",
        r"`applicationData`.*dropdown lists.*control bindings.*row-validation rules.*`Chargeable`.*`Y,N`.*row `3` through row `2000`",
        "Document the centralized dropdown contract and Chargeable validation.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents no JobFlag column",
        r"`Chargeable`.*canonical internal-labor chargeability flag.*`Internal Eligible`.*canonical readiness eligibility flag.*`Ready\.ChargeableFlag`.*`Ready\.InternalEligible`.*`Ready\.InternalJobs_Export` computes `Internal Ready Final`.*source table does not carry a separate `Internal Ready` override column.*The starter no longer carries a `JobFlag` column or a separate visible `Eligible` fallback column",
        "Document the current Chargeable/Internal Eligible Ready inputs.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents manual Composite Cat",
        r"`Composite Cat` remains a manual pre-formula planning-table helper.*Excel's built-in sort.*remove-duplicates.*Data > Subtotal",
        "Document Composite Cat as an operator worksheet helper, not formula output.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide links planning worksheet structure map",
        r"planning_worksheet_structure_map\.md.*Yes/No columns.*formula dependencies",
        "Link the Yes/No dependency structure map from the starter guide.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents notes workflow setup",
        r"Setup Notes Workflow.*ApplyNotes Control.*ExistingMeetingNotes.*NewPlanningNotes.*NewTimeline.*NewStatus.*tblDecisionStaging.*apply_notes\.ts",
        "Document the notes workflow setup in the starter workbook guide.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents optional asset setup",
        r"Setup Asset Workflow.*optional.*apply_asset_mappings\.ts.*not part of the default setup path.*docs/asset_setup_workflow\.md",
        "Document that asset setup remains opt-in.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter,
        "starter guide documents asset register setup",
        r"Setup Asset Workflow.*Asset Register.*tblAssets.*advisory `LinkedProjectID` dropdown.*Rerunning it recreates",
        "Document tblAssets and reset behavior in the starter guide.",
    )
    check_required_regex(
        results,
        "docs/workbook_import_map.md",
        import_map,
        "import map documents Ready chargeability input",
        r"`Ready`.*Ready\.ChargeableFlag.*`Chargeable`.*Ready\.InternalEligible.*`Internal Eligible`.*Ready\.InternalJobs_Export.*`Internal Ready Final` is computed in the export.*no source-table `Internal Ready`.*no `JobFlag` starter column.*no visible `Eligible` fallback column",
        "Document Ready's current Chargeable/Internal Eligible input boundary.",
    )
    check_required_regex(
        results,
        "docs/workbook_import_map.md",
        import_map,
        "import map documents validation lists",
        r"`Validation Lists`.*dropdown values",
        "Document the dropdown-source sheet in the import map.",
    )
    check_required_regex(
        results,
        "docs/workbook_import_map.md",
        import_map,
        "import map documents visible control bindings",
        r"`PM_Filter_Dropdowns`.*\$B\$2.*`Burndown_Cut_Target`.*\$E\$2",
        "Document worksheet-visible control name bindings.",
    )
    check_required_regex(
        results,
        "docs/workbook_import_map.md",
        import_map,
        "import map links planning worksheet structure map",
        r"planning_worksheet_structure_map\.md.*Yes/No inputs.*formula dependencies",
        "Link the Yes/No dependency structure map from the import map.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents reference parse shape",
        r"Used range.*`A1:BM98`.*Explicit list validation.*`O3:O98`.*`Y,N`",
        "Keep the public-safe reference parse facts visible.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents full Planning Table span",
        r"64-column table from `A:BL`.*`A`.*`Composite Cat`.*`O`.*`Annual Projected`.*`AZ`.*`December`.*`BA`.*`Comments`.*`BL`.*`Canceled`",
        "Document the whole public Planning Table, not only the Yes/No subset.",
    )
    if starter_rows:
        missing_headers = [header for header in starter_rows[0] if f"`{header}`" not in structure_map]
        add(
            results,
            not missing_headers,
            "docs/planning_worksheet_structure_map.md",
            "structure map covers all starter headers",
            "all starter headers present" if not missing_headers else "missing: " + ", ".join(missing_headers),
            "Keep the structure map aligned to samples/planning_table_starter.tsv.",
        )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map lists all Yes/No columns",
        r"`M`.*`Chargeable`.*`BE`.*`Internal Eligible`.*`BL`.*`Canceled`",
        "List the complete public starter Yes/No field set.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents Chargeable dependencies",
        r"`Chargeable`.*Ready\.ChargeableFlag.*Ready\.InternalReady3.*Ready\.InternalJobs_Export.*Search\.Projects_Health",
        "Document the Chargeable formula dependency chain.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents readiness dependencies",
        r"`Internal Eligible`.*Ready\.InternalEligible.*Ready\.InternalJobs_Export.*computed `Internal Ready Final`.*no source-table `Internal Ready` override column",
        "Document the Ready helper dependencies for eligibility and readiness fields.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents manual Composite Cat",
        r"`Composite Cat` is kept as a manual pre-formula planning-table helper.*Excel's built-in sort.*remove-duplicates.*Data > Subtotal",
        "Keep Composite Cat outside formula-owned output.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents no visible Eligible fallback",
        r"`Internal Eligible` is the canonical readiness eligibility field.*Do not add a separate visible `Eligible` fallback column",
        "Prevent reintroducing duplicate eligibility flags.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents Canceled dependencies",
        r"`Canceled`.*kind\.CapConsumeMask.*main report.*analysis screens.*defer helpers.*Ready\.InternalJobs_Export",
        "Document the cancellation dependency chain.",
    )
    check_required_regex(
        results,
        "docs/planning_worksheet_structure_map.md",
        structure_map,
        "structure map documents non-Yes/No exceptions",
        r"`Carry-over` and `Future Tier Override` are not Yes/No columns",
        "Prevent confusing adjacent planning fields with Yes/No flags.",
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
        bool(starter_rows) and all(len(row) == 64 for row in starter_rows),
        "samples/planning_table_starter.tsv",
        "starter table row width",
        "all rows have 64 tab-delimited columns" if starter_rows else "starter table is empty",
        "Keep every starter row aligned to the 64-column Planning Table contract.",
    )
    add(
        results,
        bool(starter_rows) and "Eligible" not in starter_rows[0],
        "samples/planning_table_starter.tsv",
        "starter table omits visible Eligible fallback",
        "Eligible header absent" if starter_rows else "starter table is empty",
        "Use Internal Eligible as the sole readiness eligibility input.",
    )
    add(
        results,
        bool(starter_rows) and "Internal Ready" not in starter_rows[0],
        "samples/planning_table_starter.tsv",
        "starter table omits source-table Internal Ready",
        "Internal Ready header absent" if starter_rows else "starter table is empty",
        "Keep Internal Ready Final computed in Ready.InternalJobs_Export.",
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
    workflow_starters = [
        ("samples/decision_staging_starter.tsv", decision_starter, 23, r"ReviewRow\tGroupType\tGroupValue\tCategory\tProjDesc.*BudgetRowFound"),
        ("samples/asset_setup_starter.tsv", asset_setup_starter, 15, r"ProjectKey\tCandidateAssetId\tProjectDescription.*ApplyMessage.*ProjectKey\tProjectDescription\tChangeType.*ApplyMessage"),
        ("samples/project_asset_map_starter.tsv", project_asset_map_starter, 11, r"ProjectKey\tProjectDescription\tAssetId\tAssetLabel\tAssetType\tAssetState\tEvidenceId\tMappingStatus\tApplyStatus\tAppliedOn\tApplyMessage"),
        ("samples/semantic_assets_starter.tsv", semantic_assets_starter, 15, r"ProjectKey\tProjectDescription\tCandidateAssetId\tAssetLabel\tAssetType\tProposedChangeType.*ApplyMessage"),
        ("samples/asset_changes_starter.tsv", asset_changes_starter, 10, r"ChangeId\tProjectKey\tChangeType\tSourceAssetId\tTargetAssetId\tInstalledState\tEvidenceId\tChangeStatus\tAppliedOn\tApplyMessage"),
        ("samples/asset_state_history_starter.tsv", asset_state_history_starter, 8, r"EventId\tAssetId\tProjectKey\tAssetState\tEvidenceId\tEventSource\tEventOn\tApplyMessage"),
    ]
    for file_name, text, expected_width, header_pattern in workflow_starters:
        rows = [line.split("\t") for line in text.splitlines() if line.strip()]
        add(
            results,
            bool(rows) and all(len(row) == expected_width for row in rows),
            file_name,
            "workflow starter row width",
            f"all rows have {expected_width} tab-delimited columns" if rows else "starter file is empty",
            "Keep workflow starter TSV rows aligned to the add-in-created table headers.",
        )
        check_required_regex(
            results,
            file_name,
            text,
            "workflow starter headers match setup contract",
            header_pattern,
            "Keep starter TSV headers aligned to the add-in-created table headers.",
        )
    check_required_regex(
        results,
        "samples/decision_staging_starter.tsv",
        decision_starter,
        "decision staging starter includes runnable ApplyNotes smoke row",
        r"Sample over-projected work.*NOTE_TIMELINE_STATUS.*Review forecast against latest meeting note.*OK\t1\tTRUE\t\t\tStarter row; run ApplyNotes once to prepare and again to apply",
        "Keep the starter decision row ready to prepare and apply against the starter Planning Table.",
    )
    for file_name, text, checks in [
        (
            "office-scripts/README.md",
            office_scripts_readme,
            [
                ("documents apply notes script", r"apply_notes\.ts.*Planning Review!P:R.*Decision Staging.*Planning Table"),
                ("documents live ApplyNotes control status", r"apply_notes\.ts.*Planning Review!O1:R3.*last phase/result/next action"),
                ("documents asset mapping script", r"apply_asset_mappings\.ts.*accepted asset setup rows.*asset mapping.*state-history"),
                ("states formulas review and scripts write", r"Formula modules create review queues.*Office Scripts perform controlled writes"),
                ("excludes graph export", r"External graph export is not part of this release"),
            ],
        ),
        (
            "office-scripts/apply_notes.ts",
            apply_notes_script,
            [
                ("uses Decision Staging table", r"Decision Staging.*tblDecisionStaging"),
                ("documents two-pass behavior", r"Run 1 reads Planning Review P:R.*refreshes formula-backed tblDecisionStaging.*Run 2 applies"),
                ("writes expected fields", r"Planning Notes.*Timeline.*Comments.*Status"),
                ("stages Planning Review source inputs", r"buildReviewPrepareRows.*reviewValues.*reviewRow\[15\].*refreshFormulaBackedApplyTableRows.*finish\(\s*\"prepare\""),
                ("preserves Decision Staging formula columns", r"indexedNotesFormula.*DROP\(Notes\.FromArrayv,1\).*ReviewRow.*refreshFormulaBackedApplyTableRows.*setColumnFormulas.*BudgetMatchCount"),
                ("blocks duplicate staged Planning Table targets", r"duplicateTargetMessage.*Planning Review rows target Planning Table row.*preparedTargetCounts"),
                ("updates Planning Review control area", r"CONTROL_RANGE_ADDRESS\s*=\s*\"O1:R3\".*writeApplyNotesControl.*ApplyNotes Control.*Last Run.*Next Action.*finish.*writeApplyNotesControl"),
                ("caps visible comment row height after archive", r"COMMENTS_ROW_HEIGHT_POINTS\s*=\s*45.*commentsRowsToFormat.*setWrapText\(true\).*setRowHeight\(COMMENTS_ROW_HEIGHT_POINTS\)"),
                ("uses clear operator statuses and reset", r"STATUS_BLOCKED.*STATUS_SKIPPED.*resetFormulaBackedApplyTable.*finish\(\s*\"reset\""),
                ("records operator-readable apply messages", r"Blocked: expected exactly 1 Planning Table match.*Prepared: matched Planning Table row.*Applied: updated"),
                ("clears Planning Review source inputs", r"REVIEW_SHEET_NAME\s*=\s*\"Planning Review\".*REVIEW_INPUT_COL0\s*=\s*15.*clearReviewInputs.*cleared Planning Review P:R"),
                ("does not overwrite staged input columns on apply", r"flushApplyTable.*applyStatusRange\.setValues.*msgRange\.setValues(?!.*newNoteRange\.setValues)"),
            ],
        ),
        (
            "office-scripts/apply_asset_mappings.ts",
            apply_assets_script,
            [
                ("uses expected asset tables", r"tblSemanticAssets.*tblAssetPromotionQueue.*tblAssetMappingStaging.*tblProjectAssetMap.*tblAssetChanges.*tblAssetStateHistory"),
                ("states asset register boundary", r"Asset Register / tblAssets.*does not create, overwrite, or enrich tblAssets"),
                ("validates new asset rule", r"new_asset requires target_asset_id"),
                ("validates replacement rule", r"replace_asset requires source_asset_id and target_asset_id"),
                ("excludes graph export", r"does not export graph data|External graph export was not run"),
            ],
        ),
    ]:
        for check, pattern in checks:
            check_required_regex(
                results,
                file_name,
                text,
                f"v0.2.0 Office Script {check}",
                pattern,
                "Keep Office Scripts aligned to the controlled-write release boundary.",
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
    taskpane_css = read_text(ROOT / "addin" / "taskpane.css")
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
        ("creates starter sheets", r"Start Here.*Planning Table.*Cap Setup.*Planning Review.*Source Status.*Analysis Hub.*Asset Hub.*Asset Finance Hub.*Validation Lists"),
        ("creates data import bridge sheets and tables", r"Data Import Setup.*PQ Budget Input.*PQ Budget QA.*tblDataSourceProfile.*tblBudgetImportParameters.*tblBudgetImportContract.*tblBudgetInput.*tblBudgetImportStatus.*tblBudgetImportIssues"),
        ("creates Planning Table as table", r"sheet:\s*\"Planning Table\".*tableName:\s*\"tblPlanningTable\".*planning_table_starter\.tsv"),
        ("creates workbook manifest table", r"Workbook Manifest.*tblWorkbookManifest.*workbook_manifest\.tsv"),
        ("defines workbook visibility rules", r"sheetVisibilityRules.*Start Here.*visible.*PQ Budget Input.*hidden.*Validation Lists.*hidden.*Decision Staging.*hidden"),
        ("uses Office.js sheet visibility enum", r"Excel\.SheetVisibility\.visible.*Excel\.SheetVisibility\.hidden"),
        ("keeps asset setup opt-in", r"Asset Hub\", visibility: \"hidden\".*Asset Finance Hub\", visibility: \"hidden\".*setupAssetWorkflow.*formatAssetRegisterSheet.*formatAssetHubOnboarding.*showOptionalAssetWorkflowSheets"),
        ("applies native asset register validation", r"(?=.*assetTypes: \[\"Equipment\", \"Building\", \"Vehicle\", \"System\", \"Space\", \"Other\"\])(?=.*AssetID.*Enter a stable asset identifier, e\.g\. AHU-001\.)(?=.*AssetName.*Enter a plain-English asset name\.)(?=.*AssetType.*Choose a simple type such as Equipment, Building, Vehicle, System, Space, or Other\.)(?=.*Status.*Choose the current lifecycle status\.)(?=.*ReplacementCost\", nonNegative: true)(?=.*UsefulLifeYears\", nonNegative: true)(?=.*LinkedProjectID.*allowUnknown: true.*Optional project/job key from the current workbook planning data\. This does not imply external refresh or sync\.)"),
        ("orders visible sheets by workbook flow", r"sheet\.position\s*=\s*position"),
        ("styles page and section headers", r"formatPageHeader.*#1F4E79.*#2F75B5.*formatSectionHeader.*#E2EFDA.*#F2F2F2"),
        ("defines application data", r"applicationData\s*=\s*\{.*starterTables.*dropdownLists.*visibleControls.*rowValidationRules"),
        ("defines validation lists", r"dropdownLists\s*:\s*\{.*months.*groupFields.*futureFilters.*closedRows.*statuses.*yesNo.*booleanFlags"),
        ("includes review status and boolean flags", r"statuses:\s*\[\"Active\",\s*\"Review\".*booleanFlags:\s*\[\"TRUE\",\s*\"FALSE\"\]"),
        ("defines asset dropdown lists", r"assetStatuses.*assetConditions.*assetCriticalities.*assetChangeTypes.*assetStates.*assetPromotionStatuses.*assetMappingStatuses.*assetChangeStatuses"),
        ("validates asset ApplyReady as boolean flags", r"tblSemanticAssets:.*ApplyReady\",\s*listKey:\s*\"booleanFlags\".*tblAssetPromotionQueue:.*ApplyReady\",\s*listKey:\s*\"booleanFlags\".*tblAssetMappingStaging:.*ApplyReady\",\s*listKey:\s*\"booleanFlags\""),
        ("defines visible controls", r"visibleControls.*PM_Filter_Dropdowns.*B2.*Burndown_Cut_Target.*E2"),
        ("defines row validation max row", r"maxValidationRow:\s*2000"),
        ("defines header-driven Chargeable validation", r"rowValidationRules.*header:\s*\"Chargeable\".*listKey:\s*\"yesNo\""),
        ("uses 64-column starter layout", r"headerRange:\s*\"A2:BL2\".*requiredHeaderFill:\s*\[\"F2\",\s*\"G2\",\s*\"O2\",\s*\"P2\",\s*\"BE2\"\].*address:\s*\"O3:AZ234\".*address:\s*\"BJ3:BJ234\""),
        ("loads workbook controls", r"../modules/controls\.formula\.txt"),
        ("loads formula modules", r"../modules/kind\.formula\.txt.*../modules/analysis\.formula\.txt"),
        ("loads Source formula module", r"../modules/source\.formula\.txt"),
        ("installs workbook names", r"context\.workbook\.names\.add"),
        ("binds visible control names", r"bindVisibleControlNames\(context\).*Governed formula visible control"),
        ("installs qualified module names", r"name:\s*`\$\{moduleFile\.prefix\}\.\$\{item\.name\}`"),
        ("handles unqualified alias collisions", r"unqualifiedAliases"),
        ("validates required names", r"requiredNames"),
        ("validates workbook control names", r"PM_Filter_Dropdowns.*Future_Filter_Mode.*HideClosed_Status.*Burndown_Cut_Target"),
        ("validates implemented analysis screens", r"Analysis\.PM_SPEND_REPORT.*Analysis\.WORKING_BUDGET_SCREEN.*Analysis\.BURNDOWN_SCREEN"),
        ("validates Ready helpers", r"Ready\.ColumnOrBlank.*Ready\.InternalEligible.*Ready\.ChargeableFlag.*Ready\.InternalReady3.*Ready\.InternalJobs_Export"),
        ("validates workbook-local compatibility helpers", r"TRIMRANGE_KEEPBLANKS.*RBYROW"),
        ("formats starter workbook", r"formatDataImportSetup.*formatBudgetInput.*formatBudgetQa.*formatPlanningTable.*formatCapSetup.*formatPlanningReview.*formatStartHere.*formatSourceStatus.*formatHubShell"),
        ("enriches Start Here navigation", r"formatStartHere.*Workbook flow.*Manual workbook source / optional placeholder adapter.*Go to.*Backend/admin sheets.*setMergedPanel"),
        ("adds clickable hub table of contents", r"formatHubToc.*Go to section.*setInternalSheetLink.*documentReference.*#1F4E79"),
        ("normalizes row heights", r"normalizeSheetRows.*rowHeight"),
        ("widens Start Here flow text columns", r"formatStartHere.*C:C.*columnWidth.*D:D.*columnWidth"),
        ("widens import contract description column", r"formatDataImportSetup.*D:D.*columnWidth"),
        ("labels Planning Review control months", r"formatPlanningReview.*Report As Of Month.*Defer As Of Month.*applyHubColumnWidths\(sheet,\s*\"PlanningReview\"\)"),
        ("labels and widens Planning Review notes flow", r"formatPlanningReviewNotes.*O4:R4.*notesWorkflow\.noteHeaders.*applyHubColumnWidths\(sheet,\s*\"PlanningReview\"\)"),
        ("applies dropdown validation", r"applyListValidation.*dataValidation\.rule"),
        ("sizes validation list headers dynamically", r"getResizedRange\(0,\s*validationListColumns\.length - 1\)"),
        ("applies row validation by header", r"applyRowValidationRules.*dataRangeForHeader.*validationSourceForList"),
        ("applies non-negative validation", r"applyNonNegativeValidation.*greaterThanOrEqualTo"),
        ("validates starter header order", r"assertHeaderOrder\(planningHeaders\.values\[0\], expectedPlanningHeaders"),
        ("validates canonical budget input table", r"getItemOrNullObject\(\"tblBudgetInput\"\).*budgetInputHeaders.*assertHeaderOrder\(budgetInputHeaders\.values\[0\], expectedBudgetHeaders, \"tblBudgetInput\"\)"),
        ("validates row validation rules", r"assertRowValidationRulesConfigured.*headerIndex"),
        ("validates spill-safe control band", r'review\.getRange\("B2:E2"\)'),
        ("validates visible controls", r"assertVisibleControls\(reviewControls\.values, reviewMonths\.values\)"),
        ("validates bound control names", r"assertControlNamesBound\(controlNameItems\)"),
        ("validates cap setup rows", r"assertCapRowsAreValid\(capRows\.values\)"),
        ("renders validation summary", r"renderValidationSummary.*Sheets present.*Workbook names installed.*Dropdown lists ready"),
        ("clears stale spill blockers", r'getRange\("J2:K6"\)\.clear'),
        ("defines demo hub outputs", r"(?=.*demoOutputs\s*:\s*\[)(?=.*Planning Review)(?=.*Source Status)(?=.*Analysis Hub)(?=.*CapitalPlanning\.CAPITAL_PLANNING_REPORT)(?=.*Analysis\.BU_CAP_SCORECARD)(?=.*Analysis\.REFORECAST_QUEUE)(?=.*Analysis\.PM_SPEND_REPORT)(?=.*Analysis\.WORKING_BUDGET_SCREEN)(?=.*Analysis\.BURNDOWN_SCREEN)(?=.*Ready\.InternalJobs_Export)"),
        ("defines Source demo output", r"Source Status.*Source\.SOURCE_STATUS\(\)"),
        ("runs demo outputs from combined setup", r"runAll\(\).*setupWorkbook\(\).*installModules\(\).*validateWorkbook\(\).*insertDemoOutputs\(\{\s*validateFirst:\s*false\s*\}\)"),
        ("binds demo output action", r"bind\(\"insertDemoOutputs\",\s*insertDemoOutputs\)"),
        ("inserts demo output formulas", r"insertDemoOutputs.*validateWorkbook\(\).*placeDemoOutput"),
        ("checks main report spill range", r'getRange\("A4:N200"\).*load\(\["values", "formulas"\]\).*assertMainReportSpillReady'),
        ("reports demo spill blockers", r"assertMainReportSpillReady.*blocks the main report spill"),
        ("renders demo output summary", r"renderDemoOutputSummary.*Demo hub outputs inserted"),
        ("defines ApplyNotes template path", r'applyNotesScriptPath\s*=\s*"../office-scripts/apply_notes\.ts"'),
        ("binds ApplyNotes copy action", r'bind\("copyApplyNotesScript",\s*copyApplyNotesScript\)'),
        ("loads and displays ApplyNotes script text", r"copyApplyNotesScript.*fetchText\(applyNotesScriptPath\).*showApplyNotesScript\(applyNotesText\)"),
        ("copies ApplyNotes script text", r"copyApplyNotesScript.*copyTextToClipboard\(applyNotesText\)"),
        ("explains two-pass ApplyNotes use", r"Use ApplyNotes in two passes.*run once to prepare Decision Staging.*run again to apply"),
        ("logs ApplyNotes script import step", r"Automate > New Script"),
        ("uses clipboard fallback", r"copyTextToClipboard.*navigator\.clipboard.*catch.*legacyCopyText"),
        ("selects ApplyNotes text when clipboard is blocked", r"selectApplyNotesScript.*focus\(\).*select\(\)"),
        ("logs standard setup completion", r"Standard setup complete\. Asset workflow remains optional"),
        ("defines notes workflow setup", r"notesWorkflow.*tblDecisionStaging.*ReviewRow.*ExistingMeetingNotes.*NewPlanningNotes.*NewTimeline.*NewStatus"),
        ("creates visible ApplyNotes control area", r"formatPlanningReviewNotes.*O1:R3.*ApplyNotes Control.*Run 1: Prepare.*Run 2: Apply.*Check Decision Staging"),
        ("keys Decision Staging formulas by ReviewRow", r"indexedNotesFormula.*ReviewRow.*XMATCH.*CHOOSECOLS"),
        ("seeds Planning Review notes smoke input and script-driven staging", r"smokeInputRange:\s*\"P5:R5\".*setupNotesWorkflow.*allCellsBlank\(smokeRange\.values\).*Notes workflow ready: Planning Review P:R inputs will be staged by ApplyNotes run 1"),
        ("uses one smoke row and scalar BudgetMatchCount", r"stagingRowCount:\s*1.*BudgetMatchCount.*SUMPRODUCT.*INDEX.*Planning Table.*XMATCH\(\"Project Description\""),
        ("defines asset workflow setup", r"assetWorkflow.*tblAssets.*tblSemanticAssets.*tblAssetPromotionQueue.*tblAssetMappingStaging.*tblProjectAssetMap.*tblAssetChanges.*tblAssetStateHistory"),
        ("defines asset relationship lists", r"relationshipLists.*assetIds.*tblAssets\[AssetID\].*projectKeys.*tblAssets\[LinkedProjectID\]"),
        ("defines asset table validation rules", r"tableValidationRules.*tblAssets.*assetStatuses.*relationshipListKey:\s*\"projectKeys\".*tblProjectAssetMap.*relationshipListKey:\s*\"assetIds\""),
        ("applies asset table validation", r"setupAssetWorkflow.*buildAssetRelationshipLists.*applyTableValidationRules.*validationSourceForRelationshipList"),
        ("keeps relationship validation advisory", r"allowUnknown:\s*Boolean\(rule\.allowUnknown \|\| rule\.relationshipListKey\).*errorAlert.*showAlert:\s*false"),
        ("sets workflow table header text black", r"getHeaderRowRange\(\).*format\.font\.color\s*=\s*\"#000000\""),
        ("binds notes and asset setup buttons", r'bind\("setupNotesWorkflow",\s*setupNotesWorkflow\).*bind\("setupAssetWorkflow",\s*setupAssetWorkflow\)'),
        ("strips module comments", r"stripBlockComments"),
        ("compacts installed formula bodies", r"compactFormulaBody.*stripBlockComments.*inQuotedSheet"),
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
    add(
        results,
        "HYPERLINK(" not in taskpane.upper(),
        "addin/taskpane.js",
        "task pane avoids formula hyperlink navigation",
        "formula hyperlink navigation absent",
        "Use plain labels or real worksheet hyperlinks instead of cached HYPERLINK formulas.",
    )
    check_required_regex(
        results,
        "addin/taskpane.js",
        taskpane,
        "task pane uses real internal worksheet hyperlinks",
        r"(?=.*setInternalSheetLink)(?=.*range\.hyperlink)(?=.*documentReference)(?=.*formatHubToc)(?=.*setInternalSheetLink\(sheet\.getRange)",
        "Use real worksheet hyperlinks for Start Here and hub section navigation when the host supports them.",
    )
    run_all_match = re.search(r"async function runAll\(\)\s*\{(?P<body>.*?)\n\s*\}", taskpane, flags=re.S)
    run_all_body = run_all_match.group("body") if run_all_match else ""
    add(
        results,
        "setupNotesWorkflow" in run_all_body,
        "addin/taskpane.js",
        "task pane includes notes setup in default runAll",
        "setupNotesWorkflow is in runAll" if "setupNotesWorkflow" in run_all_body else "setupNotesWorkflow missing from runAll",
        "Keep notes workflow setup in the normal combined setup path.",
    )
    add(
        results,
        "setupAssetWorkflow" not in run_all_body,
        "addin/taskpane.js",
        "task pane keeps asset setup opt-in",
        "setupAssetWorkflow absent from runAll" if "setupAssetWorkflow" not in run_all_body else "setupAssetWorkflow is in runAll",
        "Do not run optional asset setup from the default combined setup path.",
    )
    add(
        results,
        "setupAssetEvidencePowerQuery" not in run_all_body,
        "addin/taskpane.js",
        "task pane keeps asset evidence Power Query setup opt-in",
        "setupAssetEvidencePowerQuery absent from runAll" if "setupAssetEvidencePowerQuery" not in run_all_body else "setupAssetEvidencePowerQuery is in runAll",
        "Do not run optional asset evidence Power Query setup from the default combined setup path.",
    )
    for forbidden_symbol in [
        "setupAssetEvidencePowerQuery",
        "copyAssetEvidencePowerQueryTemplates",
        "copyAssetEvidencePowerQueryVbaInstaller",
        "validateAssetEvidencePowerQuery",
        "tblAssetEvidenceSource",
        "tblAssetEvidenceRules",
        "tblAssetEvidenceOverrides",
    ]:
        add(
            results,
            forbidden_symbol not in taskpane,
            "addin/taskpane.js",
            f"task pane omits duplicate asset evidence action {forbidden_symbol}",
            "duplicate action absent",
            "Keep asset evidence Power Query install in the generated seed workbook and PowerShell installer.",
        )

    add(
        results,
        "JobFlag" not in taskpane,
        "addin/taskpane.js",
        "task pane does not configure JobFlag",
        "JobFlag absent",
        "Do not reintroduce JobFlag to the starter setup contract.",
    )
    add(
        results,
        'header: "Eligible"' not in taskpane,
        "addin/taskpane.js",
        "task pane does not configure visible Eligible fallback",
        "Eligible row validation absent",
        "Use Internal Eligible as the starter readiness eligibility field.",
    )
    add(
        results,
        'header: "Internal Ready"' not in taskpane,
        "addin/taskpane.js",
        "task pane does not configure source-table Internal Ready",
        "Internal Ready row validation absent",
        "Keep Internal Ready Final computed in Ready.InternalJobs_Export.",
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
        "addin/taskpane.html",
        taskpane_html,
        "task pane has combined setup output button",
        r'id="runAll".*Setup \+ Install \+ Validate \+ Outputs',
        "Make the primary setup path create the demo output sheets.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane has demo output button",
        r'id="insertDemoOutputs".*Insert Demo Outputs',
        "Expose the output insertion rerun action.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane has ApplyNotes copy button",
        r'id="copyApplyNotesScript".*Copy ApplyNotes Script',
        "Expose a one-click ApplyNotes script handoff for sideload users.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane has notes and asset workflow buttons",
        r'id="setupNotesWorkflow".*Setup Notes Workflow.*id="setupAssetWorkflow".*Setup Asset Workflow',
        "Expose notes workflow setup and optional asset workflow setup.",
    )
    for forbidden_markup in [
        "setupAssetEvidencePowerQuery",
        "copyAssetEvidencePowerQueryTemplates",
        "copyAssetEvidencePowerQueryVbaInstaller",
        "validateAssetEvidencePowerQuery",
        "assetEvidencePowerQueryText",
    ]:
        add(
            results,
            forbidden_markup not in taskpane_html,
            "addin/taskpane.html",
            f"task pane markup omits duplicate asset evidence control {forbidden_markup}",
            "duplicate control absent",
            "Keep asset evidence Power Query install in the generated seed workbook and PowerShell installer.",
        )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane has ApplyNotes import instruction",
        r"Automate</code> -> <code>New Script</code>.*save as <code>ApplyNotes</code>",
        "Show the exact script import step in the task pane.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane explains ApplyNotes two-pass workflow",
        r"Planning Review!P:R.*Run <code>ApplyNotes</code> once to prepare.*run it again to apply",
        "Keep the two-pass ApplyNotes operator flow visible in the task pane.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane has ApplyNotes script text fallback",
        r'id="applyNotesScriptText".*readonly.*hidden',
        "Keep a visible in-pane fallback for blocked clipboard access.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane shows ApplyNotes source path",
        r"Template source:\s*<code>\.\./office-scripts/apply_notes\.ts</code>",
        "Keep the source template path visible in the task pane.",
    )
    check_required_regex(
        results,
        "addin/taskpane.html",
        taskpane_html,
        "task pane marks asset workflow optional",
        r'id="setupAssetWorkflow".*class="optional-action".*Setup Asset Workflow.*Optional',
        "Visually mark asset setup as optional.",
    )
    check_required_regex(
        results,
        "addin/taskpane.css",
        taskpane_css,
        "task pane color-codes optional asset action",
        r"button\.optional-action.*background:\s*#fff4ce.*border-color:\s*#c19c00.*\.badge",
        "Keep the optional asset action visually distinct.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document operator launcher",
        r"operator-style local use.*Start-AddIn\.ps1.*workbook copy.*npm dependencies.*does not edit a workbook by itself",
        "Show the safer launcher before the developer smoke helper.",
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
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document validation lists",
        r"Creates the `Validation Lists` sheet",
        "Document that the add-in creates dropdown source lists.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document application data contract",
        r"`applicationData` model.*starter sheets.*dropdown source lists.*control bindings.*row-validation rules",
        "Document the centralized add-in data model.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document Chargeable validation",
        r"`Chargeable` dropdown.*header.*row `2`.*rows `3:2000`.*`Y,N`",
        "Document header-driven Chargeable dropdown validation.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document no JobFlag starter column",
        r"`Chargeable`.*chargeability input.*`Search`.*`Ready`.*`Internal Eligible`.*readiness eligibility input.*Ready\.InternalEligible.*Ready\.InternalJobs_Export.*computes `Internal Ready Final`.*no source-table `Internal Ready`.*no `JobFlag` starter column.*no separate visible `Eligible` fallback column",
        "Document the current starter chargeability and eligibility inputs.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document manual Composite Cat",
        r"`Composite Cat` remains a manual pre-formula helper.*Excel Data > Subtotal",
        "Document Composite Cat as a worksheet-layer helper.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document visible controls",
        r"`Planning Review` controls in `B2:E2`.*`M2:N2`",
        "Document the visible starter workbook controls.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document control rebinding",
        r"PM_Filter_Dropdowns -> 'Planning Review'!\$B\$2.*Burndown_Cut_Target -> 'Planning Review'!\$E\$2",
        "Document that unqualified controls point to visible cells.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs preserve output ranges",
        r"leaves `A4:N200` open.*leaves `O4:R200` open",
        "Document preserved spill and note ranges.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document validation summary",
        r"Validation summary:.*Sheets present.*Workbook names installed.*Dropdown lists ready",
        "Document the operator-facing validation summary.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document combined setup output action",
        r"Setup \+ Install \+ Validate \+ Outputs.*creates the starter sheets.*installs formulas.*validates the workbook contract.*inserts the demo hub formulas",
        "Document that the primary add-in action now creates outputs.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document demo outputs",
        r"Insert Demo Outputs.*Planning Review.*CapitalPlanning\.CAPITAL_PLANNING_REPORT.*Analysis Hub.*Burndown.*Analysis\.BURNDOWN_SCREEN.*Internal Jobs.*Ready\.InternalJobs_Export",
        "Document the output insertion rerun action.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document demo spill guard",
        r"checks `Planning Review!A4:N200`.*block the main report spill.*safe to rerun",
        "Document the pre-insert main report spill guard.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document ApplyNotes helper",
        r"`ApplyNotes` setup helper.*loads the script template from `\.\./office-scripts/apply_notes\.ts`.*displays the script text when clipboard access is blocked",
        "Document the ApplyNotes script handoff and blocked-clipboard fallback inside the add-in.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document notes setup in normal path",
        r"Setup Notes Workflow.*normal `Setup \+ Install \+ Validate \+ Outputs` path.*Planning Review!O1:R3.*ApplyNotes Control.*Planning Review!O:R.*tblDecisionStaging",
        "Document that notes setup is included in the normal setup path.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document live ApplyNotes control status",
        r"`ApplyNotes` updates that control area after each normal run.*last phase.*result.*next action",
        "Document that ApplyNotes updates the worksheet control area after it runs.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document ApplyNotes import step",
        r"Copy ApplyNotes Script.*clipboard access is blocked.*Automate -> New Script.*save the script as `ApplyNotes`",
        "Document the operator path for importing the ApplyNotes script.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document optional asset setup",
        r"Setup Asset Workflow.*optional.*not run from the default path.*Asset Register.*tblAssets.*Asset Setup.*Project Asset Map.*Semantic Assets.*Asset Changes.*Asset State History",
        "Document that asset setup remains opt-in.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document asset evidence Power Query helper",
        r"outside the Office\.js task pane.*start_asset_evidence_pq_installer\.ps1.*install_asset_evidence_pq_workbook\.ps1.*tblAssetEvidenceSource.*qAssetEvidence_Normalized.*qQA_AssetEvidence_MappingQueue",
        "Document the generated asset evidence Power Query seed workflow.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document asset relationship dropdowns",
        r"native validation and input messages.*Rerunning it recreates.*advisory relationship lists for `Asset ID` and `Project Key`.*LinkedProjectID.*allows blank or manually typed IDs",
        "Document the optional asset setup dropdown and reset behavior.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document optional asset color cue",
        r"color-codes the asset setup button as optional.*standard setup completion message",
        "Document the UI cue that asset setup is separate.",
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
        "README.md",
        readme,
        "README documents asset evidence Power Query assistant",
        r"Asset Evidence Power Query.*seed-workbook.*samples/power-query/asset-evidence/.*start_asset_evidence_pq_installer\.ps1.*install_asset_evidence_pq_workbook\.ps1.*docs/asset_evidence_power_query\.md",
        "Surface the optional asset evidence Power Query seed path from the README.",
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
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records ApplyNotes add-in handoff",
        r"Add in-add-in ApplyNotes script handoff",
        "Record the task-pane ApplyNotes script setup path.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records asset tracker starter",
        r"Promote asset workflow to tracker starter.*tblAssets.*relationship dropdowns.*starter/reset",
        "Record the asset tracker starter setup change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records asset evidence Power Query assistant",
        r"(?=.*Add asset evidence Power Query seed workbook)(?=.*tblAssetEvidenceSource)(?=.*tblAssetEvidenceRules)(?=.*tblAssetEvidenceOverrides)(?=.*start_asset_evidence_pq_installer\.ps1)(?=.*install_asset_evidence_pq_workbook\.ps1)(?=.*Setup Asset Workflow.*does not create asset-evidence Power Query setup or output tables)(?=.*PresentWithClassifiedEvidence.*classifier metadata)",
        "Record the optional asset evidence Power Query seed path.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records optional asset setup UI",
        r"Clarify optional asset setup UI.*Color-coded.*Setup Asset Workflow.*black text",
        "Record the optional asset setup UI correction.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records operator add-in launcher",
        r"Add operator add-in launcher.*Start-AddIn\.ps1.*README_FIRST\.md.*start:addin.*workbook copy",
        "Record the safer non-developer add-in launch path.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records live ApplyNotes control area",
        r"Make ApplyNotes control area live.*Planning Review!O1:R3.*last phase.*result counts.*next action",
        "Record that ApplyNotes updates the worksheet control area after it runs.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Planning Review ApplyNotes control area",
        r"Add Planning Review ApplyNotes control area.*Planning Review!O1:R3.*ApplyNotes Control.*run `ApplyNotes` once to prepare.*run `ApplyNotes` again to apply",
        "Record the visible workbook control area that explains the two-pass ApplyNotes flow.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records ApplyNotes comments row-height cap",
        r"Cap visible ApplyNotes comment row height.*45-point height.*full `Comments` text remains stored",
        "Record the visual row-height cap for ApplyNotes comment archives.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records ReviewRow ApplyNotes staging fix",
        r"Key ApplyNotes staging by Planning Review row.*ReviewRow.*duplicating the first staged source row.*duplicate staged writes",
        "Record the ApplyNotes multi-row staging identity fix.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records ApplyNotes message and reset tightening",
        r"Tighten ApplyNotes messages and staging reset.*Blocked.*Skipped.*reset path.*two-pass operator flow",
        "Record ApplyNotes message and reset behavior changes.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Planning Review ApplyNotes smoke input",
        r"Seed Planning Review ApplyNotes smoke input.*Planning Review!P:R.*Setup Notes Workflow.*fresh workbook can test ApplyNotes.*ApplyNotes` run 1.*formula-backed `tblDecisionStaging`",
        "Record the seeded ApplyNotes smoke input behavior.",
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
    burndown_detail = extract_named_formula(analysis, "BURNDOWN_SCREEN_DETAIL_FROM_AXIS")
    check_required_regex(
        results,
        "modules/analysis.formula.txt",
        burndown_detail,
        "BURNDOWN_SCREEN_DETAIL_FROM_AXIS guards Cut Candidate display errors",
        r"BlankCut,\s*MAKEARRAY\(ROWS\(DetailSorted\),\s*1,\s*LAMBDA\(rw,\s*c,\s*\"\"\)\).*CutMask,\s*IFERROR\(\(CutTarget > 0\) \* \(CumRem <= CutTarget\),\s*0\).*CutDisp,\s*IFERROR\(IF\(CutTarget <= 0,\s*BlankCut,\s*IF\(CutMask,\s*\"Yes\",\s*BlankCut\)\),\s*BlankCut\)",
        "Keep the generated starter free of visible #N/A values in the burndown Cut Candidate column.",
    )


def audit_asset_evidence_power_query_contract(results: list[Result]) -> None:
    query_dir = ROOT / "samples" / "power-query" / "asset-evidence"
    seed_builder = read_text(ROOT / "tools" / "build_asset_evidence_pq_seed.ps1")
    workbook_installer = read_text(ROOT / "tools" / "install_asset_evidence_pq_workbook.ps1")
    expected_templates = {
        "qAssetEvidence_Normalized.m": [
            r"tblAssetEvidenceSource",
            r"EvidenceId.*SourceSystem.*ProjectKey.*AssetId.*FundingSource.*DepreciationClass",
        ],
        "qAssetEvidence_Classified.m": [
            r"tblAssetEvidenceRules.*tblAssetEvidenceOverrides",
            r"ClassifiedCategoryId.*ClassifiedCategoryName.*ClassifierSourceType.*ClassifierSourceLabel.*ClassifierRuleId",
        ],
        "qAssetEvidence_Linked.m": [
            r"qAssetEvidence_Classified",
            r"ContextCategoryId.*ContextCategoryName.*HasMappedEvidence.*MappedCategoryId.*MappedCategoryName.*AssetId.*ProjectKey",
        ],
        "qAssetEvidence_Status.m": [
            r"qAssetEvidence_Linked",
            r"HasClassifiedEvidence.*ClassifiedCategoryId.*ClassifiedCategoryName.*ClassifierSourceType.*ClassifierSourceLabel.*ClassifierRuleId",
            r"PresentWithSourceEvidence.*PresentWithMappedEvidence.*PresentWithClassifiedEvidence",
            r"Mapped context requires classifier review",
        ],
        "qAssetEvidence_ModelInputs.m": [
            r"qAssetEvidence_Status",
            r"FundingSource.*DepreciationClass.*PresentWithClassifiedEvidence",
        ],
        "qQA_AssetEvidence_MappingQueue.m": [
            r"qAssetEvidence_Status",
            r"PresentWithMappedEvidence.*PresentWithClassifiedEvidence.*ReviewIssue",
        ],
    }

    for file_name, patterns in expected_templates.items():
        path = query_dir / file_name
        label = rel(path)
        text = read_text(path)
        add(
            results,
            bool(text),
            label,
            "asset evidence Power Query template exists",
            "template present",
            "Keep every expected asset evidence M template in samples/power-query/asset-evidence/.",
        )
        for pattern in patterns:
            check_required_regex(
                results,
                label,
                text,
                "asset evidence Power Query contract",
                pattern,
                "Keep the public asset evidence query contract intact.",
            )

    check_required_regex(
        results,
        "tools/build_asset_evidence_pq_seed.ps1",
        seed_builder,
        "asset evidence seed builder creates loaded workbook artifact",
        r"(?=.*Asset_Evidence_PQ_Seed\.xlsx)(?=.*tblAssetEvidenceSource)(?=.*tblAssetEvidenceRules)(?=.*tblAssetEvidenceOverrides)(?=.*Add-LoadedQueryTable)(?=.*SaveAs\(\$resolvedOutputPath,\s*51\))",
        "Keep the seed workbook build reproducible from text sources.",
    )
    check_required_regex(
        results,
        "tools/build_asset_evidence_pq_seed.ps1",
        seed_builder,
        "asset evidence seed builder names all expected queries",
        r"qAssetEvidence_Normalized.*qAssetEvidence_Classified.*qAssetEvidence_Linked.*qAssetEvidence_Status.*qAssetEvidence_ModelInputs.*qQA_AssetEvidence_MappingQueue",
        "Keep the seed workbook aligned with the M template set.",
    )
    check_required_regex(
        results,
        "tools/install_asset_evidence_pq_workbook.ps1",
        workbook_installer,
        "asset evidence workbook installer protects original target",
        r"(?=.*TargetWorkbookPath)(?=.*OutputPath)(?=.*OutputPath must be a workbook copy)(?=.*Copy-Item)(?=.*ReplaceExisting)(?=.*Add-AssetEvidenceSetup)",
        "Keep the install path non-destructive by writing a separate output workbook.",
    )
    check_required_regex(
        results,
        "tools/start_asset_evidence_pq_installer.ps1",
        read_text(ROOT / "tools" / "start_asset_evidence_pq_installer.ps1"),
        "asset evidence button launcher exposes install button",
        r"(?=.*System\.Windows\.Forms)(?=.*OpenFileDialog)(?=.*Build Seed)(?=.*Install Asset Evidence PQ)(?=.*install_asset_evidence_pq_workbook\.ps1)",
        "Keep a local button path for the workbook installer.",
    )
    check_required_regex(
        results,
        "tools/install_asset_evidence_pq_workbook.ps1",
        workbook_installer,
        "asset evidence workbook installer names all expected sheets",
        r"Asset Evidence Setup.*PQ Asset Evidence Normalized.*PQ Asset Evidence Classified.*PQ Asset Evidence Linked.*PQ Asset Evidence Status.*PQ Asset Evidence Model Inputs.*PQ Asset Evidence Mapping Queue",
        "Keep the install script aligned with the seed workbook sheet set.",
    )
    check_required_regex(
        results,
        "tools/install_asset_evidence_pq_workbook.ps1",
        workbook_installer,
        "asset evidence workbook installer names all expected queries",
        r"qAssetEvidence_Normalized.*qAssetEvidence_Classified.*qAssetEvidence_Linked.*qAssetEvidence_Status.*qAssetEvidence_ModelInputs.*qQA_AssetEvidence_MappingQueue",
        "Keep the install script aligned with the M template set.",
    )


def audit_budget_input_power_query_contract(results: list[Result]) -> None:
    query_dir = ROOT / "samples" / "power-query" / "budget-input"
    expected_templates = {
        "qBudget_Source_CurrentWorkbook.m": [
            r"Excel\.CurrentWorkbook\(\)\{\[Name = \"tblPlanningTable\"\]\}\[Content\]",
        ],
        "qBudget_Source_AzureSql.m": [
            r"SERVER_OR_ENDPOINT_PLACEHOLDER",
            r"DATABASE_OR_WORKSPACE_PLACEHOLDER",
            r"vBudgetPlanningWorkbookContract",
            r"Sql\.Database",
        ],
        "qBudget_Source_Dataverse.m": [
            r"DATAVERSE_ENVIRONMENT_PLACEHOLDER",
            r"gef_budgetplanningworkbookcontract",
            r"CommonDataService\.Database",
        ],
        "qBudget_Source_FabricSqlEndpoint.m": [
            r"FABRIC_SQL_ENDPOINT_PLACEHOLDER",
            r"DATABASE_OR_WORKSPACE_PLACEHOLDER",
            r"vBudgetPlanningWorkbookContract",
            r"Sql\.Database",
        ],
        "qBudget_Source_Selected.m": [
            r"tblBudgetImportParameters",
            r"ActiveAdapter",
            r"qBudget_Source_AzureSql",
            r"qBudget_Source_Dataverse",
            r"qBudget_Source_FabricSqlEndpoint",
            r"qBudget_Source_CurrentWorkbook",
        ],
        "qBudget_Normalized.m": [
            r"qBudget_Source_Selected",
            r"tblBudgetImportContract",
            r"Table\.SelectColumns\(Source, ContractColumns, MissingField\.UseNull\)",
        ],
        "qBudget_WideContract.m": [
            r"qBudget_Normalized",
            r"Table\.ReorderColumns\(Source, ContractColumns, MissingField\.UseNull\)",
        ],
        "qBudget_Input.m": [
            r"qBudget_WideContract",
        ],
        "qBudget_Status.m": [
            r"tblBudgetImportParameters",
            r"ActiveAdapter",
            r"qBudget_Input",
            r"QueryName.*SourceMode.*LastRefreshUtc.*RowCount.*Status.*Message",
        ],
        "qBudget_Issues.m": [
            r"qBudget_Source_Selected",
            r"tblBudgetImportContract",
            r"MissingColumn.*ExtraColumn.*SchemaOK",
        ],
    }

    for file_name, patterns in expected_templates.items():
        path = query_dir / file_name
        label = rel(path)
        text = read_text(path)
        add(
            results,
            bool(text),
            label,
            "budget input Power Query template exists",
            "template present",
            "Keep every expected budget-input M template in samples/power-query/budget-input/.",
        )
        for pattern in patterns:
            check_required_regex(
                results,
                label,
                text,
                "budget input Power Query contract",
                pattern,
                "Keep the public budget input query contract intact.",
            )

    placeholders = [
        "SERVER_OR_ENDPOINT_PLACEHOLDER",
        "DATABASE_OR_WORKSPACE_PLACEHOLDER",
        "DATAVERSE_ENVIRONMENT_PLACEHOLDER",
        "FABRIC_SQL_ENDPOINT_PLACEHOLDER",
    ]
    combined = "\n".join(read_text(query_dir / file_name) for file_name in expected_templates)
    for placeholder in placeholders:
        add(
            results,
            placeholder in combined,
            "samples/power-query/budget-input",
            f"budget input adapters use {placeholder}",
            "placeholder present",
            "Keep public adapter templates placeholder-only.",
        )


def audit_integration_bridge_contract(results: list[Result]) -> None:
    builder = read_text(ROOT / "tools" / "build_governance_starter_workbook.ps1")
    taskpane = read_text(ROOT / "addin" / "taskpane.js")
    manifest = read_text(ROOT / "samples" / "workbook_manifest.tsv")
    contract_doc = read_text(ROOT / "docs" / "integration_bridge_contract.md")
    database_import_doc = read_text(ROOT / "docs" / "database_import_contract.md")
    starter_doc = read_text(ROOT / "docs" / "starter_workbook.md")
    addin_doc = read_text(ROOT / "docs" / "office_addin.md")
    readme = read_text(ROOT / "README.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")
    financial_query = read_text(ROOT / "samples" / "power-query" / "integration-bridge" / "qBridge_FinancialProjectRegister.m")
    approved_query = read_text(ROOT / "samples" / "power-query" / "integration-bridge" / "qBridge_ApprovedProjectEvidence.m")

    expected_financial_headers = [
        "Source ID",
        "Job ID",
        "ProjectKey",
        "Project Description",
        "Status",
        "BU",
        "Category",
        "Site",
        "PM",
    ]
    expected_approved_headers = [
        "ProjectKey",
        "EvidenceId",
        "EvidenceType",
        "EvidencePath",
        "EvidenceName",
        "Extension",
        "DocumentAreaID",
        "DocumentAreaName",
        "CategoryID",
        "CategoryName",
        "DateModified",
        "ReviewStatus",
        "ApprovedOn",
        "ReviewerNotes",
        "StatusSignal",
    ]

    financial_rows = read_tsv_rows("samples/financial_project_register_export_starter.tsv")
    approved_rows = read_tsv_rows("samples/approved_project_evidence_starter.tsv")
    add(
        results,
        bool(financial_rows) and financial_rows[0] == expected_financial_headers,
        "samples/financial_project_register_export_starter.tsv",
        "integration bridge project register headers",
        "headers match bridge export contract",
        "Keep the financial project register export shape aligned to the integration bridge contract.",
    )
    add(
        results,
        bool(approved_rows) and approved_rows[0] == expected_approved_headers,
        "samples/approved_project_evidence_starter.tsv",
        "integration bridge approved evidence headers",
        "headers match approved evidence import contract",
        "Keep the approved evidence import shape aligned to the integration bridge contract.",
    )

    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        manifest,
        "workbook manifest exposes Integration Bridge",
        r"Integration Bridge\ttblFinancialProjectRegisterExport; tblApprovedProjectEvidence\tControl\tReviewed evidence handoff\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin",
        "Keep the operator bridge visible in generated workbook editions.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates Integration Bridge",
        r"Build-IntegrationBridge.*tblFinancialProjectRegisterExport.*financial_project_register_export_starter\.tsv.*tblApprovedProjectEvidence.*approved_project_evidence_starter\.tsv",
        "Build the bridge tables from tracked starter TSV files.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder keeps bridge advisory",
        r"Integration Bridge.*advisory only.*does not create official financial projects.*must not overwrite planning status, create projects, or turn documentation signals into official finance status",
        "Keep reviewed evidence as context rather than workbook writeback logic.",
    )
    check_required_regex(
        results,
        "addin/taskpane.js",
        taskpane,
        "task pane creates Integration Bridge starter tables",
        r"integrationBridge: \"Integration Bridge\".*tblFinancialProjectRegisterExport.*financial_project_register_export_starter\.tsv.*tblApprovedProjectEvidence.*approved_project_evidence_starter\.tsv.*formatIntegrationBridge",
        "Keep the Office.js blank-workbook setup aligned with the generated starter.",
    )
    check_required_regex(
        results,
        "addin/taskpane.js",
        taskpane,
        "task pane keeps Integration Bridge visible and advisory",
        r"Integration Bridge\", visibility: \"visible\".*ProjectKey is Source ID & \"-\" & Job ID.*Approved evidence is advisory only.*must not overwrite planning status, create projects",
        "Keep the add-in bridge visible but non-authoritative.",
    )
    check_required_regex(
        results,
        "samples/power-query/integration-bridge/qBridge_FinancialProjectRegister.m",
        financial_query,
        "financial project register query derives ProjectKey",
        r"tblBudgetInput.*\"Source ID\".*\"Job ID\".*\"ProjectKey\".*SourceId & \"-\" & JobId.*Table\.ReorderColumns",
        "Keep ProjectKey derived from Source ID and Job ID only.",
    )
    check_required_regex(
        results,
        "samples/power-query/integration-bridge/qBridge_ApprovedProjectEvidence.m",
        approved_query,
        "approved evidence query filters approved rows",
        r"tblApprovedProjectEvidence.*RequiredColumns.*\"ReviewStatus\".*Text\.Upper.*\"APPROVED\"",
        "Only approved evidence rows should flow into the advisory import query.",
    )
    check_required_regex(
        results,
        "docs/integration_bridge_contract.md",
        contract_doc,
        "integration bridge contract states refresh safety",
        r"ProjectKey.*Source ID & \"-\" & Job ID.*ReviewKey = EvidenceId & \"\|\" & CandidateProjectKey.*must not erase manual approval history",
        "Document the bridge key rules and refresh-safety boundary.",
    )
    check_required_regex(
        results,
        "docs/integration_bridge_contract.md",
        contract_doc,
        "integration bridge contract forbids authoritative side effects",
        r"does not.*create official financial projects.*update official project status.*use raw file paths as project keys.*treat candidate evidence mappings as approved rows.*overwrite manual review decisions",
        "Keep the reviewed evidence bridge advisory-only.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import_doc,
        "database import doc includes Integration Bridge boundary",
        r"Integration Bridge Boundary.*tblFinancialProjectRegisterExport.*tblApprovedProjectEvidence.*Source ID & \"-\" & Job ID.*does not auto-create financial projects.*does not update official project status",
        "Document the optional reviewed-evidence bridge beside the data import contract.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter_doc,
        "starter workbook doc includes Integration Bridge",
        r"Integration Bridge.*tblFinancialProjectRegisterExport.*tblApprovedProjectEvidence.*Source ID & \"-\" & Job ID.*Candidate mappings and review decisions stay outside the generated workbook",
        "Keep starter workbook docs aligned with the bridge tables.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "office add-in doc includes Integration Bridge",
        r"Integration Bridge.*tblFinancialProjectRegisterExport.*tblApprovedProjectEvidence.*Source ID & \"-\" & Job ID.*advisory context only",
        "Keep add-in docs aligned with the bridge setup path.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README mentions optional Integration Bridge",
        r"Integration Bridge.*Source ID.*Job ID.*ProjectKey.*does not create projects, update official status, or use raw file paths as project keys",
        "Keep the README clear about the optional bridge boundary.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records optional integration bridge",
        r"Add optional integration bridge handoff.*tblFinancialProjectRegisterExport.*ProjectKey = Source ID & \"-\" & Job ID.*tblApprovedProjectEvidence.*advisory-only",
        "Record the visible bridge sheet and advisory behavior change.",
    )


def audit_release_accelerator_contract(results: list[Result]) -> None:
    feature_status = read_text(ROOT / "docs" / "feature_status.tsv")
    feature_reporter = read_text(ROOT / "tools" / "report_feature_status.py")
    review_packet = read_text(ROOT / "tools" / "build_review_packet.py")
    artifact_scanner = read_text(ROOT / "tools" / "check_release_artifact.py")
    workflow = read_text(ROOT / ".github" / "workflows" / "validate.yml")
    package = read_text(ROOT / "package.json")
    agents = read_text(ROOT / "AGENTS.md")
    codex_contract = read_text(ROOT / "docs" / "codex_chatgpt_durable_contract.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")

    check_required_regex(
        results,
        "docs/feature_status.tsv",
        feature_status,
        "feature status file has required columns",
        r"FeatureId\tFeatureName\tExpectedStatus\tCategory\tEvidenceType\tEvidencePattern\tNotes",
        "Keep feature status source-controlled and machine-readable.",
    )
    for feature_id in [
        "canonical_budget_input",
        "integration_bridge",
        "pq_selected_adapter",
        "source_module",
        "start_here_hubs",
        "workbook_manifest",
        "artifact_sanitization",
        "release_artifact_scan",
        "ci_validation",
        "review_packet",
        "feature_status_reporter",
        "asset_editions",
        "asset_guided_start",
        "asset_finance_empty_state",
        "semantic_reference_edition",
        "semantic_crosswalk_lite",
        "semantic_starter_contract",
        "semantic_export_reference",
        "external_semantic_integration",
    ]:
        add(
            results,
            feature_id in feature_status,
            "docs/feature_status.tsv",
            f"feature status tracks {feature_id}",
            "feature row present",
            f"Track {feature_id} in docs/feature_status.tsv.",
        )
    check_required_regex(
        results,
        "tools/report_feature_status.py",
        feature_reporter,
        "feature status reporter supports expected statuses",
        r"VALID_STATUSES.*Built.*Scaffolded.*Missing.*Mismatch.*ExpectedStatus.*EvidencePattern",
        "Keep feature status output explicit about built/scaffolded/missing/mismatch state.",
    )
    check_required_regex(
        results,
        "tools/report_feature_status.py",
        feature_reporter,
        "feature status reporter fails only missing built evidence",
        r"expected_status == \"Built\".*actual_status == \"Missing\"",
        "Only block CI when a feature expected to be Built lacks evidence.",
    )
    check_required_regex(
        results,
        "tools/build_review_packet.py",
        review_packet,
        "review packet generator writes ignored packet",
        r"release_artifacts.*review_packet.*review_packet\.md.*Feature Status.*Workbook Manifest Summary.*Power Query Adapter Summary.*Asset Workflow Status",
        "Keep review packets useful for branch handoff.",
    )
    check_required_regex(
        results,
        "tools/check_release_artifact.py",
        artifact_scanner,
        "release artifact scanner checks package hazards",
        r"WORKBOOK_SUFFIXES.*\.xlsx.*\.xltx.*forbidden_needles.*#REF!.*#VALUE!.*#N/A.*#NAME\?.*#DIV/0!.*HYPERLINK\(",
        "Keep generated workbook artifacts independently scannable without Excel COM.",
    )
    check_required_regex(
        results,
        ".github/workflows/validate.yml",
        workflow,
        "GitHub Actions runs source validation",
        r"pull_request.*codex/\*\*.*python -S tools/audit_capex_module\.py.*python -S tools/lint_formulas\.py modules/\*\.formula\.txt.*python -S tools/report_feature_status\.py.*git diff --check",
        "Keep PR and codex branch validation aligned with local checks.",
    )
    for script_name in ["validate", "review:packet", "feature:status", "check:release-artifact"]:
        add(
            results,
            f'"{script_name}"' in package,
            "package.json",
            f"package script {script_name} exists",
            "script present",
            f"Expose npm script {script_name} for local reviewers.",
        )
    check_required_regex(
        results,
        "AGENTS.md",
        agents,
        "agent contract requires built scaffolded missing report",
        r"Feature status:.*Built:.*Scaffolded:.*Missing:.*Validation:.*Not changed:",
        "Keep implementation reports reviewer-friendly.",
    )
    check_required_regex(
        results,
        "docs/codex_chatgpt_durable_contract.md",
        codex_contract,
        "durable contract documents review packet and feature status",
        r"Required Release-Handoff Report.*tools/build_review_packet\.py.*tools/report_feature_status\.py.*Built.*Scaffolded.*Missing",
        "Keep the Codex/ChatGPT handoff convention documented.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records release accelerator tooling",
        r"Add v0\.5 release accelerator tooling.*feature status.*review-packet.*CI validation.*release-artifact package scanner",
        "Record the release accelerator and reviewer handoff tooling.",
    )


def audit_governance_starter_template_contract(results: list[Result]) -> None:
    builder = read_text(ROOT / "tools" / "build_governance_starter_workbook.ps1")
    readme = read_text(ROOT / "README.md")
    starter_doc = read_text(ROOT / "docs" / "starter_workbook.md")
    addin_doc = read_text(ROOT / "docs" / "office_addin.md")
    database_import = read_text(ROOT / "docs" / "database_import_contract.md")
    workbook_map_doc = read_text(ROOT / "docs" / "workbook_left_to_right_map.md")
    finance_doc = read_text(ROOT / "docs" / "asset_finance_model_modules.md")
    asset_quick_start = read_text(ROOT / "docs" / "asset_quick_start.md")
    setup_doc = read_text(ROOT / "docs" / "asset_setup_workflow.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")
    package = read_text(ROOT / "package.json")
    gitignore = read_text(ROOT / ".gitignore")
    workbook_manifest = read_text(ROOT / "samples" / "workbook_manifest.tsv")

    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        workbook_manifest,
        "workbook manifest has required columns",
        r"SheetName\tTableName\tZone\tRole\tVisibility\tPresence\tEdition\tFriendlyName\tDataFlowFrom\tDataFlowTo\tOperatorAction\tNotes",
        "Keep the workbook manifest machine-readable for sheet visibility and navigation.",
    )
    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        workbook_manifest,
        "workbook manifest lists planning edition visible surface",
        r"Start Here.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Source Status.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Data Import Setup.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Planning Table.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Cap Setup.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Planning Review.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Analysis Hub.*\tvisible\tGenerated\tPlanning;AssetsLite;AssetsFull;SemanticTwin",
        "Keep the default Planning edition focused on planning surfaces.",
    )
    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        workbook_manifest,
        "workbook manifest makes asset hubs opt-in by edition",
        r"Asset Hub.*\tvisible\tGenerated\tAssetsLite;AssetsFull;SemanticTwin\tAsset Hub.*Asset Finance Hub.*\tvisible\tGenerated\tAssetsFull;SemanticTwin\tAsset Finance Hub.*Asset Register.*\tvisible\tGenerated\tAssetsLite;AssetsFull;SemanticTwin\tAsset Register",
        "Keep asset surfaces out of the default Planning edition and make Asset Register visible only in asset-enabled editions.",
    )
    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        workbook_manifest,
        "workbook manifest hides backend and admin sheets",
        r"PQ Budget Input.*\thidden\t.*PQ Budget QA.*\thidden\t.*Validation Lists.*\thidden\t.*Decision Staging.*\thidden\t.*Automation Setup.*\thidden\t.*Asset Evidence Setup.*\thidden\t.*PQ Asset Evidence Model Inputs.*\thidden\t",
        "Keep governed backend sheets present but hidden by default.",
    )
    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        workbook_manifest,
        "workbook manifest keeps advanced asset sheets hidden",
        r"Semantic Map Setup.*\thidden\tGenerated\tAssetsFull;SemanticTwin.*Asset Setup.*\thidden\tGenerated\tAssetsLite;AssetsFull;SemanticTwin.*Project Asset Map.*\thidden\tGenerated\tAssetsLite;AssetsFull;SemanticTwin.*Asset Changes.*\thidden\tGenerated\tAssetsLite;AssetsFull;SemanticTwin.*Asset State History.*\thidden\tGenerated\tAssetsLite;AssetsFull;SemanticTwin.*Asset Evidence Setup.*\thidden\tGenerated\tAssetsFull;SemanticTwin",
        "Keep simple asset entry visible while advanced asset, evidence, finance, and semantic setup sheets stay admin-scoped.",
    )
    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        workbook_manifest,
        "workbook manifest hides legacy separate output sheets",
        r"BU Cap Scorecard.*\thidden\tOptionalLegacy\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Reforecast Queue.*\thidden\tOptionalLegacy\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*PM Spend Report.*\thidden\tOptionalLegacy\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Working Budget.*\thidden\tOptionalLegacy\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Burndown.*\thidden\tOptionalLegacy\tPlanning;AssetsLite;AssetsFull;SemanticTwin.*Asset Depreciation.*\thidden\tOptionalLegacy\tAssetsFull;SemanticTwin.*Asset Finance Charts.*\thidden\tOptionalLegacy\tAssetsFull;SemanticTwin",
        "Keep legacy output sheet names out of the default visible workbook surface.",
    )

    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder keeps dropdown values aligned",
        r"statuses = @\(\"Active\", \"Review\".*booleanFlags = @\(\"TRUE\", \"FALSE\"\).*@{ Key = \"booleanFlags\"; Header = \"Boolean Flag\" }.*Apply-TableListValidation -Table \$table -Header \"ApplyReady\" -ListKey \"booleanFlags\"",
        "Keep generated starter dropdown lists aligned with seeded notes and asset values.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder configures simple asset register validation",
        r"assetTypes = @\(\"Equipment\", \"Building\", \"Vehicle\", \"System\", \"Space\", \"Other\"\).*Apply-AssetRegisterValidation.*AssetID.*Enter a stable asset identifier, e\.g\. AHU-001\..*AssetName.*Enter a plain-English asset name\..*AssetType.*Choose a simple type such as Equipment, Building, Vehicle, System, Space, or Other\..*Status.*Choose the current lifecycle status\..*Apply-TableNonNegativeValidation.*ReplacementCost.*Apply-TableNonNegativeValidation.*UsefulLifeYears.*LinkedProjectID.*AllowUnknown.*Optional project/job key from the current workbook planning data\. This does not imply external refresh or sync\.",
        "Keep tblAssets as the native Excel entry table for simple manual assets.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates xlsx and xltx artifacts",
        r"(?=.*ValidateSet\(\"Planning\", \"AssetsLite\", \"AssetsFull\", \"SemanticTwin\"\))(?=.*Governance_Starter\$artifactSuffix\.xlsx)(?=.*Governance_Starter\$artifactSuffix\.xltx)(?=.*SaveAs\(\$templateWorkbookPath,\s*54\))",
        "Keep the generated starter template build reproducible.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder uses source-controlled inputs",
        r"(?=.*modules\\controls\.formula\.txt)(?=.*modules\\analysis\.formula\.txt)(?=.*modules\\source\.formula\.txt)(?=.*modules\\asset_finance\.formula\.txt)(?=.*samples\\planning_table_starter\.tsv)(?=.*samples\\budget_import_contract_starter\.tsv)(?=.*samples\\workbook_manifest\.tsv)(?=.*samples\\asset_register_starter\.tsv)(?=.*samples\\asset_finance_assumptions_starter\.tsv)(?=.*install_asset_evidence_pq_workbook\.ps1)",
        "Keep the workbook artifact generated from tracked text sources.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates workbook UX simplification surfaces",
        r"(?=.*Build-StartHere)(?=.*Build-AnalysisHub)(?=.*Build-AssetHub)(?=.*Build-AssetFinanceHub)(?=.*Build-WorkbookManifest)(?=.*tblWorkbookManifest)(?=.*Start Here)(?=.*Analysis Hub)(?=.*Asset Hub)(?=.*Asset Finance Hub)(?=.*Format-PageHeader)(?=.*Format-SectionHeader)",
        "Keep workbook navigation and output hubs generated from source-controlled builder code.",
    )
    legacy_output_creators = [
        '@{ Sheet = "BU Cap Scorecard"',
        '@{ Sheet = "Reforecast Queue"',
        '@{ Sheet = "PM Spend Report"',
        '@{ Sheet = "Working Budget"',
        '@{ Sheet = "Burndown"',
        '@{ Sheet = "Internal Jobs"',
        '@{ Sheet = "Asset Review"',
        '@{ Sheet = "Asset Depreciation"',
        '@{ Sheet = "Asset Funding Requirements"',
        '@{ Sheet = "Asset Finance Totals"',
        '@{ Sheet = "Asset Finance Charts"',
    ]
    add(
        results,
        not any(needle in builder for needle in legacy_output_creators),
        "tools/build_governance_starter_workbook.ps1",
        "governance starter no longer creates separate demo output sheets",
        "legacy output sheet creators absent",
        "Route demo outputs through Analysis Hub, Asset Hub, and Asset Finance Hub.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder applies manifest visibility and opens Start Here",
        r"Apply-WorkbookManifestVisibility.*Edition.*samples\\workbook_manifest\.tsv.*editionTokens.*includedInEdition.*visibleSheetOrder.*Move\(\$Workbook\.Worksheets\.Item\(1\)\).*Worksheets\.Item\(\"Start Here\"\)\.Activate\(\)",
        "Keep backend sheets hidden by default and make Start Here the generated workbook front door.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder normalizes table header text",
        r"Format-TableHeader.*HeaderRowRange\.Font\.Color\s*=\s*0.*tblPlanningTable.*A2:BL2.*Font\.Color\s*=\s*0.*tblCapSetup.*A2:B2.*Font\.Color\s*=\s*0",
        "Keep generated table headers readable after Excel table styles are applied.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates Automation Setup sheet",
        r"(?=.*Build-AutomationSetup)(?=.*Automation Setup)(?=.*tblAutomationSetup)(?=.*ApplyNotes\.ts)(?=.*Automate > New Script)(?=.*Add-Worksheet.*Automation Setup)",
        "Keep the generated template explicit about optional Office Script import.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates data import bridge",
        r"(?=.*Build-DataImportSetup)(?=.*Build-BudgetInput)(?=.*Build-BudgetQA)(?=.*Data Import Setup)(?=.*PQ Budget Input)(?=.*PQ Budget QA)(?=.*tblDataSourceProfile)(?=.*tblBudgetImportParameters)(?=.*tblBudgetImportContract)(?=.*tblBudgetInput)(?=.*tblBudgetImportStatus)(?=.*tblBudgetImportIssues)(?=.*Source Status)",
        "Keep the generated starter aligned to the v0.5 canonical budget import bridge.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder keeps Start Here flow table shaped",
        r"New-Object 'object\[\]\[\]' 7.*Manual workbook source / optional placeholder adapter.*tblStartHereFlow",
        "Keep tblStartHereFlow as a four-column workbook table instead of a flattened one-column array.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder enriches Start Here navigation",
        r"(?=.*tblStartHereNavigation)(?=.*Set-InternalWorksheetLink)(?=.*Set-MergedPanel)(?=.*Backend/admin sheets)(?=.*refresh or re-sync before relying on formula outputs)",
        "Keep Start Here useful as a front-door navigation and source-boundary sheet.",
    )
    add(
        results,
        "HYPERLINK(" not in builder.upper(),
        "tools/build_governance_starter_workbook.ps1",
        "governance starter builder avoids formula hyperlink navigation",
        "formula hyperlink navigation absent",
        "Use real worksheet hyperlinks or plain labels instead of cached HYPERLINK formulas.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder adds hub tables of contents",
        r"Add-HubTableOfContents.*Go to section.*tblAnalysisHubSections.*tblAssetHubSections.*TopLeft \"A4\".*tblAssetFinanceHubSections.*TopLeft \"A4\".*tblSemanticMapHubSections",
        "Keep stacked hub sections navigable.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder keeps hub contents readable after column templates",
        r"Set-HubTableOfContentsColumnWidths.*ColumnWidth = 28.*ColumnWidth = 64.*Apply-HubColumnWidthTemplate.*Set-HubTableOfContentsColumnWidths",
        "Keep generated hub contents tables readable after fixed hub column widths are applied.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder sanitizes release artifacts",
        r"(?=.*Set-PublicWorkbookProperties)(?=.*Sanitize-WorkbookPackage)(?=.*Assert-WorkbookPackagePublic)(?=.*x15ac:absPath)(?=.*Governed Excel Formula Modules)(?=.*iCloud)(?=.*One)(?=.*release.*_artifacts)",
        "Strip local path metadata and public-release document properties from generated workbook packages.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder scans visible sheets for release errors",
        r"Assert-NoVisibleWorkbookErrors.*#N/A.*#REF!.*#VALUE!.*#NAME\?.*#DIV/0!.*CalculateFull",
        "Block public release artifacts with visible formula errors on visible sheets.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder normalizes generated row heights",
        r"Normalize-GeneratedSheetRows.*row-height normalization.*Build-AnalysisHub.*Build-AssetHub.*Build-AssetFinanceHub",
        "Keep generated starter sheets readable without over-expanded wrapped rows.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder applies fixed column-width templates",
        r"Apply-HubColumnWidthTemplate.*SourceStatus.*PlanningReview.*AssetFinance.*Template \"Analysis\".*Template \"Asset\".*Template \"AssetFinance\"",
        "Keep Source Status, Planning Review, and hub sheets readable after formula spills.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder widens import contract description",
        r"tblBudgetImportContract.*Columns\.Item\(3\)\.ColumnWidth\s*=\s*58.*Columns\.Item\(4\)\.ColumnWidth\s*=\s*64.*Columns\.Item\(6\)\.ColumnWidth\s*=\s*34.*Columns\.Item\(7\)\.ColumnWidth\s*=\s*58",
        "Keep tblBudgetImportContract[Description] and parameter Value/Description columns readable in the generated workbook.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder labels Planning Review control months",
        r"Report As Of Month.*Defer As Of Month.*Apply-HubColumnWidthTemplate -Worksheet \$Worksheet -Template \"PlanningReview\"",
        "Keep Planning Review M2:N2 month controls self-explanatory.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder labels Planning Review notes flow",
        r"notesHeaderRows.*ExistingMeetingNotes.*TopLeft \"O4\".*O4:R4.*Interior\.Color.*Apply-HubColumnWidthTemplate -Worksheet \$Worksheet -Template \"PlanningReview\"",
        "Keep Planning Review O1:R4 readable as the ApplyNotes flow and notes input header block.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates Asset Finance bridge",
        r"(?=.*Build-AssetFinanceSetup)(?=.*Build-AssetFinanceHub)(?=.*Asset Finance Setup)(?=.*tblAssetFinanceAssumptions)(?=.*AssetFinance)(?=.*Asset Finance Hub)(?=.*FINANCE_START_HERE)(?=.*FINANCE_READINESS_STATUS)(?=.*Asset Depreciation)(?=.*Asset Funding Requirements)(?=.*Asset Finance Totals)(?=.*Asset Finance Charts)(?=.*tblAssetEvidence_ModelInputs)",
        "Keep generated asset finance setup and output sheets aligned with the v0.4 bridge.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder creates guided asset hub",
        r"tblAssetWorkflowSettings.*assetWorkflowModes.*To enter one asset, go to Asset Register\..*Open Asset Register.*ASSET_REGISTER_FIELD_GUIDE.*ASSET_REGISTER_START_HERE.*ASSET_REGISTER_STATUS.*ASSET_REGISTER_ISSUES.*ASSET_NEXT_ACTIONS.*ASSET_REVIEW_QUEUE.*ASSET_GLOSSARY.*ASSET_TABLE_MAP",
        "Keep Asset Hub as simple asset entry first, technical queues second.",
    )
    check_required_regex(
        results,
        "samples/asset_finance_assumptions_starter.tsv",
        read_text(ROOT / "samples" / "asset_finance_assumptions_starter.tsv"),
        "asset finance assumptions starter has minimum columns",
        r"DepreciationClass\tUsefulLifeYears\tDepreciationMethod\tFundingSource\tFundingRequirementRule\tChartGroup",
        "Keep the generated assumption table shape stable.",
    )
    starter_asset_files = [
        "samples/asset_register_starter.tsv",
        "samples/asset_setup_starter.tsv",
        "samples/semantic_assets_starter.tsv",
        "samples/project_asset_map_starter.tsv",
        "samples/asset_changes_starter.tsv",
        "samples/asset_state_history_starter.tsv",
    ]
    for file_name in starter_asset_files:
        text = read_text(ROOT / file_name)
        add(
            results,
            all(token not in text for token in ["ASSET-001", "CHG-001", "EVT-001"]),
            file_name,
            "public asset starter has no fake demo rows",
            "demo identifiers absent",
            "Keep public asset starter TSVs headers-only or blank-row only; put fake rows under samples/demo/asset_workflow/.",
        )
    for demo_file in [
        "samples/demo/asset_workflow/asset_register_demo.tsv",
        "samples/demo/asset_workflow/semantic_assets_demo.tsv",
        "samples/demo/asset_workflow/asset_setup_demo.tsv",
        "samples/demo/asset_workflow/project_asset_map_demo.tsv",
        "samples/demo/asset_workflow/asset_changes_demo.tsv",
        "samples/demo/asset_workflow/asset_state_history_demo.tsv",
    ]:
        add(
            results,
            bool(read_text(ROOT / demo_file)),
            demo_file,
            "asset demo row file exists",
            "demo file present",
            "Keep optional demo asset rows separate from public starter TSVs.",
        )
    check_required_regex(
        results,
        "package.json",
        package,
        "npm exposes governance starter build",
        r"build:governance-starter.*build_governance_starter_workbook\.ps1",
        "Expose the generated starter build from npm scripts.",
    )
    check_required_regex(
        results,
        ".gitignore",
        gitignore,
        "workbook templates remain ignored",
        r"\*\.xltx.*\*\.xltm",
        "Keep generated Excel templates out of source control.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents generated governance starter template",
        r"Generated Governance Starter Template.*build_governance_starter_workbook\.ps1.*Governance_Starter\.xltx.*formula modules.*starter TSVs.*M templates",
        "Surface the generated template path in the README.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents generated data import bridge",
        r"v0\.5 data import bridge.*Data Import Setup.*PQ Budget Input.*PQ Budget QA.*tblBudgetInput.*Planning Table.*remains the manual starter source",
        "Surface the v0.5 canonical import layer in the generated starter docs.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents simplified workbook UX",
        r"Starter Workbook Editions.*Governance_Starter\.xltx.*planning-only.*Start Here -> Source Status -> Data Import Setup -> Integration Bridge -> Planning Table -> Cap Setup -> Planning Review -> Analysis Hub.*AssetsLite.*Asset Hub.*AssetsFull.*Asset Finance Hub",
        "Document the default visible workbook surface.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents Automation Setup sheet",
        r"Automation Setup.*ApplyNotes\.ts.*Automate -> New Script.*does not embed or auto-install Office Scripts",
        "Surface the Office Script release-asset boundary in the README.",
    )
    check_required_regex(
        results,
        "README.md",
        readme,
        "README documents v0.4 asset finance bridge",
        r"Asset Finance Setup.*tblAssetFinanceAssumptions.*AssetFinance.*Asset Finance Hub.*depreciation.*funding requirements.*totals.*chart-ready feeds.*tblAssetEvidence_ModelInputs.*PresentWithClassifiedEvidence = TRUE",
        "Surface the asset finance bridge in the README.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter_doc,
        "starter docs document generated template contents",
        r"Generated Template.*Governance_Starter\.xltx.*Data Import Setup.*tblBudgetInput.*Automation Setup.*formula-module workbook names.*optional asset workflow starter tables.*asset evidence Power Query setup",
        "Document what the generated starter template contains.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter_doc,
        "starter docs document simplified workbook UX",
        r"(?=.*default build is the `Planning` edition)(?=.*Start Here -> Source Status -> Data Import Setup -> Integration Bridge -> Planning Table -> Cap Setup -> Planning Review -> Analysis Hub)(?=.*AssetsLite.*Asset Hub)(?=.*AssetsFull.*Asset Hub.*Asset Finance Hub)(?=.*reference-only semantic crosswalk edition)(?=.*PQ Budget Input)(?=.*Validation Lists)(?=.*Workbook Manifest)(?=.*hidden by default)",
        "Document the generated workbook front door and hidden backend sheets.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter_doc,
        "starter docs document data import bridge",
        r"Data Import Bridge.*tblBudgetInput.*Planning Table.*Power Query adapter.*tblBudgetInput\[#All\].*tblBudgetImportStatus.*tblBudgetImportIssues",
        "Document the v0.5 canonical import bridge and table surfaces.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter_doc,
        "starter docs document Automation Setup sheet",
        r"(?=.*Automation Setup Sheet)(?=.*ApplyNotes\.ts)(?=.*Automate > New Script)(?=.*does not embed)(?=.*operator chooses whether to import and run)",
        "Document that Office Scripts are optional release assets, not embedded workbook automation.",
    )
    check_required_regex(
        results,
        "docs/starter_workbook.md",
        starter_doc,
        "starter docs document asset finance bridge",
        r"Asset Finance Bridge.*tblAssetFinanceAssumptions.*DepreciationClass.*UsefulLifeYears.*AssetFinance\.DEPRECIATION_SCHEDULE.*AssetFinance\.FUNDING_REQUIREMENTS.*AssetFinance\.FINANCE_TOTALS.*AssetFinance\.CHART_FEEDS.*PresentWithClassifiedEvidence = TRUE",
        "Document the generated asset finance setup and output sheets.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs point new workbook starts to generated template",
        r"preferred path is now the generated starter template.*build_governance_starter_workbook\.ps1.*default generated `.xltx` is planning-only.*-Edition AssetsLite.*-Edition AssetsFull",
        "Keep the add-in boundary aligned with the generated starter path.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document data import bridge",
        r"Data Import Setup.*PQ Budget Input.*PQ Budget QA.*tblBudgetInput.*formula modules consume `tblBudgetInput`.*refresh or re-sync.*current-workbook budget adapter",
        "Keep blank-workbook add-in setup aligned to the v0.5 canonical import bridge.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document simplified workbook UX",
        r"Start Here.*hub sheets.*Excel\.SheetVisibility\.visible.*Excel\.SheetVisibility\.hidden.*tblPlanningTable",
        "Keep blank-workbook add-in setup aligned with the simplified workbook surface.",
    )
    check_required_regex(
        results,
        "docs/database_import_contract.md",
        database_import,
        "database import contract documents canonical source boundary",
        r"(?=.*tblBudgetInput.*canonical formula source)(?=.*Planning Table.*manual.*local writeback)(?=.*ApplyNotes)(?=.*refresh or re-sync)",
        "Keep the source boundary explicit for workbook operators.",
    )
    check_required_regex(
        results,
        "docs/workbook_left_to_right_map.md",
        workbook_map_doc,
        "workbook map documents simplified visible flow",
        r"Start Here.*Source Status.*Data Import Setup.*Planning Table.*tblBudgetInput.*Planning Review.*Analysis Hub.*AssetsLite.*Asset Hub.*AssetsFull.*Asset Finance Hub.*reference-only semantic crosswalk.*hidden backend",
        "Keep the workbook map aligned to the simplified generated template.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document Automation Setup sheet",
        r"Automation Setup.*ApplyNotes\.ts.*optional Office Script release asset.*Automate -> New Script",
        "Keep the generated starter and add-in docs aligned on Office Script import.",
    )
    check_required_regex(
        results,
        "docs/office_addin.md",
        addin_doc,
        "add-in docs document generated asset finance bridge",
        r"Asset Finance Setup.*tblAssetFinanceAssumptions.*AssetFinance.*depreciation.*funding requirements.*totals.*chart-ready feeds.*Asset Finance Hub.*tblAssetEvidence_ModelInputs",
        "Keep add-in docs clear that the generated starter owns the asset finance bridge.",
    )
    check_required_regex(
        results,
        "docs/asset_finance_model_modules.md",
        finance_doc,
        "asset finance v0.4 docs start module slice",
        r"v0\.4.*codex/asset-finance-model-modules.*Automation Setup.*depreciation.*funding requirements.*totals.*chart-ready.*qAssetEvidence_ModelInputs",
        "Keep the v0.4 model-module branch scoped and documented.",
    )
    check_required_regex(
        results,
        "docs/asset_finance_model_modules.md",
        finance_doc,
        "asset finance docs document implemented bridge",
        r"Governance_Starter_AssetsFull\.xltx -> Asset Evidence Setup -> qAssetEvidence_ModelInputs -> PQ Asset Evidence Model Inputs / tblAssetEvidence_ModelInputs -> AssetFinance outputs.*AssetFinance\.FINANCE_START_HERE.*AssetFinance\.FINANCE_READINESS_STATUS.*AssetFinance\.DEPRECIATION_SCHEDULE.*AssetFinance\.FUNDING_REQUIREMENTS.*AssetFinance\.FINANCE_TOTALS.*AssetFinance\.CHART_FEEDS.*PresentWithClassifiedEvidence = TRUE",
        "Document the exact workbook bridge and classified-only finance rule.",
    )
    check_required_regex(
        results,
        "docs/asset_quick_start.md",
        asset_quick_start,
        "asset quick start explains optional asset path",
        r"Asset workflow is optional.*Start with Asset Hub to decide whether assets are needed.*Start with Asset Register to enter a simple asset.*Do not start with Asset Evidence, Asset State History, or PQ asset sheets.*LinkedProjectID`? is optional and advisory.*Asset Finance is advanced and requires classified evidence",
        "Keep first-time asset guidance explicit.",
    )
    for file_name, doc_text in [
        ("README.md", readme),
        ("README_FIRST.md", read_text(ROOT / "README_FIRST.md")),
        ("docs/starter_workbook.md", starter_doc),
        ("docs/asset_setup_workflow.md", setup_doc),
        ("docs/workbook_left_to_right_map.md", workbook_map_doc),
        ("docs/office_addin.md", addin_doc),
    ]:
        check_required_regex(
            results,
            file_name,
            doc_text,
            "docs explain simple asset entry and unchanged budget boundary",
            r"(?=.*Start with Asset Hub to decide whether assets are needed)(?=.*Start with Asset Register to enter a simple asset)(?=.*Do not start with Asset Evidence, Asset State History, or PQ asset sheets)(?=.*LinkedProjectID`? is optional and advisory)(?=.*tblBudgetInput remains the manual/canonical planning input table for this release because refresh is not surfaced)",
            "Keep docs aligned to the simple asset-entry path and unchanged budget-source boundary.",
        )
    check_required_regex(
        results,
        "docs/asset_finance_model_modules.md",
        finance_doc,
        "asset finance docs constrain v0.4 assumption semantics",
        r"v0\.4 Assumption Semantics.*tblAssetEvidence_ModelInputs.*PresentWithClassifiedEvidence = TRUE.*straight-line behavior only.*DepreciationMethod.*blank `AnnualDepreciation`.*`DepreciationIssue`.*FundingRequirementRule.*blank `FundingRequirementAmount`.*`FundingIssue`.*Chart feeds exclude unsupported rows",
        "Document the v0.4 straight-line, full-amount, issue-field, and chart-exclusion assumption semantics.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records governance starter template",
        r"Add generated governance starter template.*Governance_Starter\.xltx.*build_governance_starter_workbook\.ps1.*asset_register_starter\.tsv.*workbook binaries out of tracked source",
        "Record the generated starter template change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records generated table header normalization",
        r"Normalize generated starter table headers.*tblPlanningTable.*tblCapSetup.*black text",
        "Record generated starter table header readability fixes.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records v0.4 Automation Setup start",
        r"Start v0\.4 asset finance model branch with Automation Setup.*codex/asset-finance-model-modules.*ApplyNotes\.ts.*Automate -> New Script.*depreciation.*funding requirements.*totals.*chart-ready feeds",
        "Record the v0.4 branch start and Automation Setup worksheet.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records v0.4 asset finance bridge",
        r"Add v0\.4 asset finance bridge outputs.*modules/asset_finance\.formula\.txt.*tblAssetFinanceAssumptions.*Asset Depreciation.*Asset Funding Requirements.*Asset Finance Totals.*Asset Finance Charts.*PresentWithClassifiedEvidence = TRUE",
        "Record the v0.4 asset finance bridge implementation.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records v0.4 assumption semantics constraint",
        r"Constrain v0\.4 AssetFinance assumption semantics.*PresentWithClassifiedEvidence = TRUE.*straight-line behavior only.*full grouped classified amounts.*DepreciationMethod.*FundingRequirementRule.*ChartGroup.*DepreciationClass",
        "Record the v0.4 AssetFinance assumption semantics documentation and audit clarification.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records unsupported assumption surfacing",
        r"Surface unsupported AssetFinance assumptions.*DepreciationIssue.*FundingIssue.*FundingRequirementAmount.*AnnualDepreciation",
        "Record the formula-visible unsupported assumption surfacing behavior.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records v0.5 data import bridge",
        r"Add v0\.5 data import bridge.*tblBudgetInput.*Data Import Setup.*PQ Budget Input.*PQ Budget QA.*Source.*Power Query templates.*Copilot.*native Excel formulas",
        "Record the canonical budget input source-boundary change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records workbook UX simplification",
        r"Simplify generated workbook UX.*workbook manifest.*Start Here.*Analysis Hub.*Asset Hub.*Asset Finance Hub.*hidden by default.*tblBudgetInput.*Planning Table.*page-header.*section-header",
        "Record the simplified workbook front door and visibility behavior.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records Defer audit and dropdown drift fix",
        r"Fix Defer audit reference and dropdown drift.*defer\.Audit.*get\.GetBudgetDetailRows\(\).*Review.*ApplyReady.*qualified formula references",
        "Record the Defer audit reference fix and validation-list drift cleanup.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records optional guided asset workflow",
        r"Make asset workflow optional and guided.*Planning.*AssetsLite.*AssetsFull.*Asset Hub.*Asset Finance Hub.*ASSET_START_HERE.*FINANCE_START_HERE.*samples/demo/asset_workflow",
        "Record the editioned asset onboarding change.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records simple asset entry path",
        r"(?=.*Make simple asset entry obvious)(?=.*Asset Register)(?=.*AssetsLite)(?=.*AssetsFull)(?=.*Reference-only semantic crosswalk edition)(?=.*To enter one asset, go to Asset Register\.)(?=.*ASSET_REGISTER_START_HERE)(?=.*ASSET_REGISTER_STATUS)(?=.*ASSET_REGISTER_ISSUES)(?=.*ASSET_REGISTER_FIELD_GUIDE)(?=.*LinkedProjectID)(?=.*tblBudgetInput remains the manual/canonical planning input table)",
        "Record the asset-register entry path without changing the budget input boundary.",
    )


def audit_semantic_crosswalk_contract(results: list[Result]) -> None:
    builder = read_text(ROOT / "tools" / "build_governance_starter_workbook.ps1")
    manifest = read_text(ROOT / "samples" / "workbook_manifest.tsv")
    semantic_doc = read_text(ROOT / "docs" / "semantic_standards_strategy.md")
    readme = read_text(ROOT / "README.md")
    starter_doc = read_text(ROOT / "docs" / "starter_workbook.md")
    workbook_map_doc = read_text(ROOT / "docs" / "workbook_left_to_right_map.md")
    asset_quick_start = read_text(ROOT / "docs" / "asset_quick_start.md")
    architecture_doc = read_text(ROOT / "docs" / "reference_architecture_tree.md")
    changelog = read_text(ROOT / "docs" / "change_log.md")
    ontology_formula = read_text(ROOT / "modules" / "ontology.formula.txt")

    starter_contracts = {
        "samples/ontology_namespaces_starter.tsv": r"Prefix\tNamespace\tStandard\tNotes.*rec\thttps://w3id\.org/rec#.*brick\thttps://brickschema\.org/schema/Brick#.*gef\turn:governed-excel-formula-modules:semantic:",
        "samples/ontology_class_map_starter.tsv": r"UserFacingType\tStandard\tClassIri\tUseWhen\tNotes.*Building\tREC\trec:Building.*Room\tREC\trec:Room.*Equipment\tBrick\tbrick:Equipment.*Sensor\tBrick\tbrick:Sensor",
        "samples/ontology_relationship_map_starter.tsv": r"UserFacingRelationship\tPredicate\tUseWhen\tNotes.*Located in\trec:locatedIn.*Has point\trec:hasPoint.*Is affected by project\tgef:affectedByProject",
        "samples/project_semantic_map_starter.tsv": r"ProjectKey\tSubjectId\tSubjectClass\tRelationship\tObjectId\tObjectClass\tSourceTable\tSourceRowKey\tConfidence\tMappingStatus\tNotes",
        "samples/asset_semantic_map_starter.tsv": r"AssetId\tSubjectId\tSubjectClass\tRelationship\tObjectId\tObjectClass\tSourceTable\tSourceRowKey\tConfidence\tMappingStatus\tNotes",
        "samples/ontology_export_queue_starter.tsv": r"SubjectId\tPredicate\tObjectId\tSubjectClass\tObjectClass\tSourceTable\tSourceRowKey\tConfidence\tIssueFlag",
        "samples/ontology_issues_starter.tsv": r"IssueType\tSourceTable\tSourceRowKey\tSubjectId\tRelationship\tObjectId\tDetail",
    }
    for file_name, pattern in starter_contracts.items():
        check_required_regex(
            results,
            file_name,
            read_text(ROOT / file_name),
            "semantic crosswalk starter has required contract",
            pattern,
            "Keep semantic starter tables public-safe and machine-readable.",
        )

    check_required_regex(
        results,
        "modules/ontology.formula.txt",
        ontology_formula,
        "Ontology module exposes semantic crosswalk formulas",
        r"ONTOLOGY_START_HERE.*CLASS_MAP.*RELATIONSHIP_MAP.*SEMANTIC_MAPPING_STATUS.*ONTOLOGY_ISSUES.*TRIPLE_EXPORT_QUEUE.*JSONLD_EXPORT_HELP",
        "Keep the optional semantic crosswalk importable as workbook names.",
    )
    check_required_regex(
        results,
        "modules/ontology.formula.txt",
        ontology_formula,
        "Ontology export queue is simple triple-shaped table",
        r"SubjectId.*Predicate.*ObjectId.*SubjectClass.*ObjectClass.*SourceTable.*SourceRowKey.*Confidence.*IssueFlag",
        "Keep semantic export reviewable as a flat Subject-Predicate-Object queue.",
    )
    check_required_regex(
        results,
        "tools/build_governance_starter_workbook.ps1",
        builder,
        "governance starter builder supports reference semantic edition",
        r"(?=.*ValidateSet\(\"Planning\", \"AssetsLite\", \"AssetsFull\", \"SemanticTwin\"\))(?=.*Build-SemanticMapSetup)(?=.*Build-SemanticMapHub)(?=.*modules\\ontology\.formula\.txt)",
        "Keep the reference semantic edition isolated from the default operator flow.",
    )
    check_required_regex(
        results,
        "samples/workbook_manifest.tsv",
        manifest,
        "workbook manifest makes Semantic Map Hub reference-edition only",
        r"Semantic Map Hub.*\tvisible\tGenerated\tSemanticTwin\tSemantic Map Hub.*Semantic Map Setup.*\thidden\tGenerated\tAssetsFull;SemanticTwin",
        "Keep semantic mapping out of the default Planning and AssetsLite visible surfaces.",
    )
    check_required_regex(
        results,
        "docs/semantic_standards_strategy.md",
        semantic_doc,
        "semantic strategy marks crosswalk reference-only",
        r"archived/reference note.*not part of the current operator workflow.*not a recommended next step.*does not import full REC or Brick ontology dumps.*does not implement Azure Digital Twins.*complete integration",
        "Document the semantic crosswalk as reference-only without claiming a deployed integration.",
    )
    for file_name, text in [
        ("docs/starter_workbook.md", starter_doc),
        ("docs/workbook_left_to_right_map.md", workbook_map_doc),
        ("docs/asset_quick_start.md", asset_quick_start),
        ("docs/reference_architecture_tree.md", architecture_doc),
    ]:
        check_required_regex(
            results,
            file_name,
            text,
            "docs mention reference-only semantic crosswalk",
            r"semantic crosswalk.*reference-only|reference-only semantic crosswalk|reference crosswalk",
            "Keep public docs clear that semantic mapping is reference-only and limited.",
        )

    for suffix in [".ttl", ".owl", ".rdf", ".nt", ".jsonld"]:
        offenders = [rel(path) for path in tracked_files() if path.suffix.lower() == suffix]
        add(
            results,
            not offenders,
            "semantic crosswalk",
            f"no full ontology dump files tracked ({suffix})",
            "no ontology dump files tracked" if not offenders else f"tracked files: {', '.join(offenders[:5])}",
            "Do not commit full REC/Brick ontology dumps to this public template.",
        )

    check_required_regex(
        results,
        "tools/audit_capex_module.py",
        read_text(ROOT / "tools" / "audit_capex_module.py"),
        "audit includes Ontology qualified-reference module",
        r"\"Ontology\": \"modules/ontology\.formula\.txt\"",
        "Keep qualified formula reference checks covering the Ontology module.",
    )
    check_required_regex(
        results,
        "docs/change_log.md",
        changelog,
        "change log records semantic crosswalk reference",
        r"Add optional semantic crosswalk reference.*Semantic Map Hub.*modules/ontology\.formula\.txt.*TRIPLE_EXPORT_QUEUE",
        "Record the optional semantic crosswalk reference change.",
    )


def main() -> int:
    results: list[Result] = []
    files = tracked_files()

    audit_public_safety(results, files)
    audit_formula_files(results)
    audit_docs(results)
    audit_cap_setup_contract(results)
    audit_addin_contract(results)
    audit_governance_starter_template_contract(results)
    audit_asset_evidence_power_query_contract(results)
    audit_budget_input_power_query_contract(results)
    audit_integration_bridge_contract(results)
    audit_release_accelerator_contract(results)
    audit_semantic_crosswalk_contract(results)
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
