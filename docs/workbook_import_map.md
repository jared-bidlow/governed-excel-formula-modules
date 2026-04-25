# Workbook Import Map

Manual workbook import is the default workflow for this template.

Treat URL imports as snapshots unless the workbook owner has verified live refresh from source.

For a blank-workbook trial, use `docs/starter_workbook.md`, paste `samples/planning_table_starter.tsv` into `Planning Table!A2`, and paste `samples/cap_setup_starter.tsv` into `Cap Setup!A2` before importing formulas.

## Canonical Modules

| Workbook name | Repo source | Manual update note |
|---|---|---|
| `Controls` | `modules/controls.formula.txt` | Import first in a blank workbook so report filters have default workbook names. |
| `get` | `modules/get.formula.txt` | Import when workbook extraction helpers change. |
| `kind` | `modules/kind.formula.txt` | Import before report or analysis formulas when helper logic changes. |
| `CapitalPlanning` | `modules/capital_planning_report.formula.txt` | Import after `kind` so helper references resolve. |
| `Analysis` | `modules/analysis.formula.txt` | Import after `get` and `kind` when optional planning screens change. |

## Workbook Input Sheets

| Sheet | Required setup |
|---|---|
| `Planning Table` | Holds job rows and the finance block. The first annual projection header should be `Annual Projected`. |
| `Cap Setup` | Holds BU cap limits. Paste `samples/cap_setup_starter.tsv` into `Cap Setup!A2`, then replace the fake caps. |
| `Planning Review` | Holds meeting controls such as the as-of month and output spill areas. |

`kind.CapByBU(...)` reads `Cap Setup`, not hardcoded module constants. The BU value in `Planning Table[BU]` can include a description after a colon; only the code before the colon is used for cap lookup.

The `Controls` module defines default workbook-control names used by the public report screens: `PM_Filter_Dropdowns`, `Future_Filter_Mode`, `HideClosed_Status`, and `Burndown_Cut_Target`. Replace those names with worksheet-linked controls later if you want interactive dropdown cells.

The `get` and `kind` modules also include small compatibility helpers that older workbook copies may have held as workbook-local names, including `TRIMRANGE_KEEPBLANKS` and `RBYROW`. Importing the modules, or using the Office.js installer, should create those names in a blank workbook.

## Supporting Modules

| Workbook name | Repo source | Dependency note |
|---|---|---|
| `Notes` | `modules/notes.formula.txt` | Example note-context module. |
| `Phasing` | `modules/phasing.formula.txt` | Example monthly phasing helpers. |
| `Ready` | `modules/ready.formula.txt` | Example readiness export helpers. |
| `Search` | `modules/search.formula.txt` | Example budget search helpers. |
| `defer` | `modules/defer.formula.txt` | Example deferral-selector helpers. |

## Public Template Rule

Do not publish real workbook URLs, local paths, workbook copies, or screenshots in this map.
