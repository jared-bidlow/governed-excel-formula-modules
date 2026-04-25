# Starter Workbook

This repo does not ship a workbook. The fastest way to try the formulas is to create a blank workbook with the minimum sheet names and starter table shape.

## Minimum Sheets

Create these worksheets:

| Sheet | Purpose |
|---|---|
| `Planning Table` | Source rows for jobs, forecasts, actuals, budget, status, and grouping. |
| `Planning Review` | Output/control sheet for report formulas and the as-of month cell. |
| `Decision Staging` | Optional sheet for staged writeback examples. |

On `Planning Review`, put an as-of month abbreviation such as `Mar` in cell `M2`.

## Paste The Starter Table

Open `samples/planning_table_starter.tsv`, copy all rows, and paste into `Planning Table!A2`.

The starter includes fake rows only. Delete or replace them after confirming the formulas spill.

The included BU values, such as `BU-A: Sample Unit` and `BU-B: Sample Unit`, are fictional placeholders. Replace them with your own public-safe or private workbook values before using the template for real planning.

## Why The Starter Table Is Wide

The current formula contract expects a finance block with:

- annual projection,
- current authorized amount,
- twelve monthly projected columns,
- twelve monthly actuals columns,
- twelve monthly budget columns.

That is why the starter has three finance columns for each month:

```text
January Projected | January Actuals | January
February Projected | February Actuals | February
...
December Projected | December Actuals | December
```

The columns need to exist because helper formulas select them by position:

- `get.GetFinalProj12(...)` reads the monthly projected columns.
- `get.GetActuals12(...)` reads the monthly actuals columns.
- `get.GetBudget12(...)` reads the monthly budget columns.

Blank values are acceptable. Missing columns are not.

## What Can Be Blank

For a first test, users can leave most monthly projected and monthly budget cells blank. The most important values are:

- `2026 Projected`
- `Current Authorized Amount`
- monthly `Actuals` through the as-of month
- `Status`
- `BU`
- `Project Description`

The scorecard and report become more meaningful when the monthly budget columns are populated. The reforecast queue can still demonstrate useful behavior with blanks in monthly projected and budget columns, as long as the columns are present.

## Import Order

Import formula modules in this order:

```text
get -> kind -> Capex_Tracker_2026 -> Analysis
```

Then try these formulas on `Planning Review`:

```excel
=Analysis.REFORECAST_QUEUE()
=Analysis.BU_CAP_SCORECARD()
```
