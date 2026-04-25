# Operating Contract

## Purpose

This document defines the public-template contract for maintaining Excel formula modules as source code.

The repo demonstrates a governed workbook pattern:

- formula modules are edited in plain text,
- workbook binaries are excluded,
- important formulas are checked by static audit,
- planning screens are documented with scenarios,
- workbook-side changes are captured as text before they are applied in Excel.

## Scope

In scope:

- `modules/*.formula.txt`
- documentation under `docs/`
- validation tooling under `tools/`

Out of scope:

- `.xlsx`, `.xlsm`, `.xlsb`, or other workbook binaries,
- generated workbook packages,
- real source data,
- company-specific paths or workbook names.

## Example Module Chain

The sample import order is:

```text
get -> kind -> Capex_Tracker_2026 -> Analysis
```

The modules illustrate these responsibilities:

| Module | Role |
|---|---|
| `get` | Workbook range extraction helpers. |
| `kind` | Shared calculation, grouping, flag, and display helpers. |
| `Capex_Tracker_2026` | Main cap-feasibility report formula. |
| `Analysis` | Optional planning plugins and drilldown screens. |

## Formula Safety Rules

- Keep named formulas within workbook import limits.
- Prefer helper formulas over one giant workbook name.
- Keep hidden workbook dependencies explicit in docs.
- Keep public examples generic and data-free.
- Add audit coverage when a formula contract becomes important.

## Planning-Screen Contract

Planning screens in `Analysis` should:

- read from the same generic workbook inputs as the main report,
- avoid writing back to workbook tables,
- show decision evidence clearly,
- keep ranked detail separate from summary pivots,
- document expected behavior in `docs/scenario_matrix.md`.
