# Power Query Actuals Map

This public-template file shows the documentation pattern for a workbook actuals import layer.

It intentionally does not include real source paths, company folders, workbook filenames, sheet names, or M code.

## Generic Flow

```text
source export folder -> newest source file -> raw actuals query -> normalized actuals table -> grouped actuals lookup -> workbook monthly actual columns -> get.GetActuals12 -> kind.GetYTDActuals -> report and analysis screens
```

## What To Document In A Private Implementation

| Area | Example documentation |
|---|---|
| Source selection | Whether the query reads newest file, exact file, or configured path. |
| Raw cleanup | Header promotion, field renames, null handling, and row exclusions. |
| Normalized table | Stable column names and data types used by downstream formulas. |
| Grouped lookup | Lookup key, month/period field, amount field, and aggregation rule. |
| Review surfaces | Large charges, accrual-like rows, exceptions, and refresh metadata. |
| Formula consumption | Which workbook formulas read the normalized or grouped table. |

## Public Template Boundary

Keep real paths and workbook-specific M code out of the public repo. Publish only the documentation pattern and fake examples.
