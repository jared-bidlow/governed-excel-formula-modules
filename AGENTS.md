# Governed Excel Formula Modules - Codex Instructions

## Hard Boundaries

Do not edit Excel workbooks, `.xlsx` files, binary workbook packages, or generated workbook artifacts.

Work only on plain-text formula modules, documentation, and validation tools.

This public-template branch must not contain:

- real company names,
- real workbook paths,
- real project/job data,
- real employee, vendor, or customer data,
- internal-only process names,
- generated workbook binaries.

## Canonical Logic Artifacts

The example formula modules are:

- `modules/get.formula.txt`
- `modules/kind.formula.txt`
- `modules/capital_planning_report.formula.txt`
- `modules/analysis.formula.txt`

Keep the formula modules importable as workbook named formulas. If a formula gets too large, split it into helper formulas rather than widening one named formula past Excel's practical limits.

## Template Contract

Preserve the source-control pattern:

1. Workbook logic lives in text modules.
2. Workbook binaries remain out of scope.
3. Public docs use generic workbook and capital-planning vocabulary.
4. Static audit checks enforce public-safety strings and core formula contracts.
5. Planning screens must be documented with scenario coverage.

## Change Discipline

For each behavior change:

1. Explain the intended semantic change.
2. Show a minimal diff.
3. State whether visible report totals, subtotal flags, or cap remaining values can change.
4. Update `docs/change_log.md`.
5. Add or update static checks in `tools/audit_capex_module.py`.

Do not add private data, workbook binaries, or generated artifacts.
