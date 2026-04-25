# Workbook Import Map

Manual workbook import is the default workflow for this template.

Treat URL imports as snapshots unless the workbook owner has verified live refresh from source.

For a blank-workbook trial, use `docs/starter_workbook.md` and paste `samples/planning_table_starter.tsv` into `Planning Table!A2` before importing formulas.

## Canonical Modules

| Workbook name | Repo source | Manual update note |
|---|---|---|
| `get` | `modules/get.formula.txt` | Import when workbook extraction helpers change. |
| `kind` | `modules/kind.formula.txt` | Import before report or analysis formulas when helper logic changes. |
| `Capex_Tracker_2026` | `modules/Capex_Tracker_2026.formula.txt` | Import after `kind` so helper references resolve. |
| `Analysis` | `modules/analysis.formula.txt` | Import after `get` and `kind` when optional planning screens change. |

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
