# Name Manager Logic Inventory

This public-template file shows how to inventory workbook defined names without publishing a real workbook inventory.

## Purpose

Excel Name Manager can contain both imported module formulas and workbook-local names. Treating those as separate ownership classes makes workbook logic easier to review.

## Ownership Classes

| Ownership class | Meaning | Change route |
|---|---|---|
| Module-owned names | Period-qualified names that correspond to tracked module files. | Edit the tracked module, then import it into the workbook. |
| Workbook-local names | No-prefix helpers, controls, anchors, and sheet-local support formulas. | Inventory first; promote to text modules only when durable source control is needed. |
| Excel internal names | Internal filter/database names created by Excel. | Usually ignore unless workbook filtering breaks. |

## Example Module Prefixes

| Prefix | Owner |
|---|---|
| `get.*` | `modules/get.formula.txt` |
| `kind.*` | `modules/kind.formula.txt` |
| `CapitalPlanning.*` | `modules/capital_planning_report.formula.txt` |
| `Analysis.*` | `modules/analysis.formula.txt` |

## Public Template Boundary

Do not publish a real workbook's defined-name inventory. Replace real sheet names, table names, paths, and counts with fake examples before sharing.
