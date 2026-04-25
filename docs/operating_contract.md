# Operating Contract

## Purpose

This document defines the public-template contract for maintaining Excel formula modules as source code.

The repo demonstrates a governed workbook pattern:

- formula modules are edited in plain text,
- workbook binaries are excluded,
- important formulas are checked by static audit,
- planning screens are documented with scenarios,
- workbook-side changes are captured as text before they are applied in Excel.

## Runtime Position

Excel is the runtime because this pattern is meant for planning teams whose review, notes, and decisions already happen in workbooks.

The source-control boundary is what makes that defensible:

- plain-text formula modules can be reviewed and linted,
- workbook binaries and real data stay out of Git,
- complex workbook logic is split into named formulas instead of hidden cell sprawl,
- scenario docs describe expected planning behavior,
- audit tooling catches public-safety issues and contract drift before release.

This is not a claim that every planning process should stay in Excel. If the core requirement is multi-user writeback, permissions, APIs, or durable transactional storage, a database-backed application is the better system boundary.

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
get -> kind -> CapitalPlanning -> Analysis
```

The modules illustrate these responsibilities:

| Module | Role |
|---|---|
| `get` | Workbook range extraction helpers. |
| `kind` | Shared calculation, cap lookup, grouping, flag, and display helpers. |
| `CapitalPlanning` | Main `CAPITAL_PLANNING_REPORT()` formula. |
| `Analysis` | Optional planning plugins and drilldown screens. |

## Workbook Input Contract

The formula modules assume generic workbook inputs, not private workbook state:

- `Planning Table` contains job rows and the finance block.
- `Cap Setup` contains the `BU` and `Cap` columns used by `kind.CapByBU(...)`.
- `Planning Review` contains meeting controls and report spill areas.
- `Validation Lists` contains dropdown source values for the public starter workbook.

BU cap values should be changed in the workbook's `Cap Setup` sheet. They should not be edited inside `modules/kind.formula.txt`.

The public starter keeps visible controls on `Planning Review` and binds the unqualified control names to those cells. Module-qualified `Controls.*` names remain safe defaults.

## Add-In Packaging Contract

The Office.js add-in under `addin/` is a packaging and installation layer.

It may create starter sheets, paste sample table shapes, install workbook defined names, and validate required workbook contracts. It should not move planning logic out of native Excel formulas.

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
