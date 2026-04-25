# Starter Workbook

This repo does not ship a workbook. The fastest way to try the formulas is to create a blank workbook with the minimum sheet names and starter table shape.

## Minimum Sheets

Create these worksheets:

| Sheet | Purpose |
|---|---|
| `Planning Table` | Source rows for jobs, forecasts, actuals, budget, status, and grouping. |
| `Cap Setup` | Business-unit cap limits used by `kind.CapByBU` and `kind.PortfolioCap`. |
| `Planning Review` | Output/control sheet for report formulas and the as-of month cell. |
| `Validation Lists` | Dropdown source values used by the starter add-in. |
| `Decision Staging` | Optional sheet for staged writeback examples. |

On `Planning Review`, put an as-of month abbreviation such as `Mar` in cell `M2`. Formulas in `defer` use `N2` as their as-of month.

## Paste The Starter Table

Open `samples/planning_table_starter.tsv`, copy all rows, and paste into `Planning Table!A2`.

Open `samples/cap_setup_starter.tsv`, copy all rows, and paste into `Cap Setup!A2`.

The starter includes fake rows only. Delete or replace them after confirming the formulas spill.

The included BU values, such as `BU-A: Sample Unit` and `BU-B: Sample Unit`, are fictional placeholders. Replace them with your own public-safe or private workbook values before using the template for real planning.

The cap setup values are also fake. Replace `Cap Setup[Cap]` with the limits for your workbook. `kind.CapByBU(...)` reads the BU code before any colon in `Planning Table[BU]`, and `kind.PortfolioCap` is the sum of the cap table.

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

- `Annual Projected`
- `Current Authorized Amount`
- monthly `Actuals` through the as-of month
- `Status`
- `BU`
- `Project Description`

The scorecard and report become more meaningful when the monthly budget columns are populated. The reforecast queue can still demonstrate useful behavior with blanks in monthly projected and budget columns, as long as the columns are present.

## Import Order

Import formula modules in this order:

```text
get -> kind -> CapitalPlanning -> Analysis
```

Then try these formulas on `Planning Review`:

```excel
=Analysis.REFORECAST_QUEUE()
=Analysis.BU_CAP_SCORECARD()
```

After those spill successfully, the other implemented planning screens are:

```excel
=Analysis.PM_SPEND_REPORT()
=Analysis.WORKING_BUDGET_SCREEN()
=Analysis.BURNDOWN_SCREEN()
```

## Starter Layout And Controls

The Office.js starter can create the workbook layout for you. It writes the starter data, creates the `Validation Lists` sheet, formats the source sheets, and adds a visible control panel on `Planning Review`.

The public control cells are:

| Cell | Control | Default | Used by |
|---|---|---|---|
| `B2` | Group selector | `BU` | Main report grouping through `PM_Filter_Dropdowns`. |
| `C2` | Future filter | `All` | Main report, scorecard, and burndown future-scope filters. |
| `D2` | Closed rows | `SHOW` | Main report, scorecard, and burndown closed-row filters. |
| `E2` | Burndown cut target | `0` | Burndown candidate labeling. |
| `M2` | Report as-of month | `Mar` | Main report and `Analysis` screens. |
| `N2` | Defer as-of month | `Mar` | `defer` module examples. |

After formula installation, the unqualified workbook names point to the visible controls:

```text
PM_Filter_Dropdowns -> 'Planning Review'!$B$2
Future_Filter_Mode -> 'Planning Review'!$C$2
HideClosed_Status -> 'Planning Review'!$D$2
Burndown_Cut_Target -> 'Planning Review'!$E$2
```

The module-qualified `Controls.*` names remain defaults and documentation fallbacks.

Keep `Planning Review!A4:N200` clear for the main report spill. Keep `Planning Review!O4:R200` clear for the note-context example formulas. The visible control band stays above row 4 so it does not block the report spill.

Use the task-pane `Insert Demo Outputs` button after setup and validation if you want the public demo sheets created automatically. Before it writes the main report, it checks `Planning Review!A4:N200` and reports the first cell that would block the spill. It inserts the main report at `Planning Review!A4` and places the Analysis screens at `A4` on separate sheets named `BU Cap Scorecard`, `Reforecast Queue`, `PM Spend Report`, `Working Budget`, and `Burndown`.

## Add-In Option

The `addin/` folder provides an Office.js starter that can create the sheets, paste the starter data, install the named formulas, and validate the workbook contract from a task pane.

See `docs/office_addin.md` for the packaging boundary.
