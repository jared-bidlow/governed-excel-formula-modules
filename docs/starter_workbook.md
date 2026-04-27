# Starter Workbook

The repo keeps workbook logic and setup sources as text, then can generate a local starter workbook/template artifact when Excel is available.

## Generated Template

Build the generated governance starter:

```powershell
.\tools\build_governance_starter_workbook.ps1
```

or:

```powershell
npm run build:governance-starter
```

The build writes ignored artifacts under `release_artifacts/governance-starter/`:

```text
Governance_Starter.xlsx
Governance_Starter.xltx
```

Use `Governance_Starter.xltx` as the Excel template. Use `Governance_Starter.xlsx` for inspection and smoke testing. The generator pulls from source-controlled formula modules, starter TSVs, and M templates, so the workbook artifact can be rebuilt instead of reviewed as source.

The generated starter includes:

- planning source/cap setup sheets,
- validation lists and visible controls,
- an `Automation Setup` sheet that explains how to import the optional `ApplyNotes.ts` release asset,
- formula-module workbook names,
- demo planning output sheets,
- notes staging,
- optional asset workflow starter tables,
- asset evidence Power Query setup and loaded output sheets.

The fastest no-build path is still a blank workbook with the minimum sheet names and starter table shape.

## Minimum Sheets

Create these worksheets:

| Sheet | Purpose |
|---|---|
| `Planning Table` | Source rows for jobs, forecasts, actuals, budget, status, and grouping. |
| `Cap Setup` | Business-unit cap limits used by `kind.CapByBU` and `kind.PortfolioCap`. |
| `Planning Review` | Output/control sheet for report formulas and the as-of month cell. |
| `Validation Lists` | Dropdown source values used by the starter add-in. |
| `Decision Staging` | Notes/status/timeline staging table created by the notes workflow. |

Optional asset setup creates additional worksheets only when `Setup Asset Workflow` is selected:

- `Asset Register`
- `Asset Setup`
- `Project Asset Map`
- `Semantic Assets`
- `Asset Changes`
- `Asset State History`

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

The Office.js starter can create the workbook layout for you. It writes the starter data, creates the `Validation Lists` sheet, formats the source sheets, and adds a visible control panel on `Planning Review`. Its setup behavior is driven by the `applicationData` model in `addin/taskpane.js`, which defines dropdown lists, control bindings, and row-validation rules in one place.

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

On `Planning Table`, the add-in finds row-validation targets by header name. The `Chargeable`, `Internal Eligible`, and `Canceled` columns receive `Y,N` dropdowns from row `3` through row `2000`; the same model also drives the current status dropdown.

Treat `Chargeable` as the canonical internal-labor chargeability flag and `Internal Eligible` as the canonical readiness eligibility flag. The `Search` helpers inspect `Chargeable`, and the `Ready` export helpers use both fields when deriving the example internal-ready output. `Ready.ChargeableFlag` and `Ready.InternalEligible` resolve these inputs by header name, not by hardcoded column letters. `Ready.InternalJobs_Export` computes `Internal Ready Final` in its output; the source table does not carry a separate `Internal Ready` override column. The starter no longer carries a `JobFlag` column or a separate visible `Eligible` fallback column.

`Composite Cat` remains a manual pre-formula planning-table helper. It can be used for Excel's built-in sort, remove-duplicates, and Data > Subtotal workflows before the formula reports run; the add-in does not try to compute it.

See `docs/planning_worksheet_structure_map.md` for the public-safe reference map of Yes/No columns and formula dependencies.

Keep `Planning Review!A4:N200` clear for the main report spill. `Setup Notes Workflow` uses `Planning Review!O1:R3` for the `ApplyNotes Control` area and `Planning Review!O4:R200` for note-context and note-input columns. The visible control bands stay above row 4 so they do not block the report spill.

`Setup Notes Workflow` creates the `Planning Review!O1:R3` `ApplyNotes Control` area and the `Planning Review!O:R` note columns:

- `ExistingMeetingNotes`
- `NewPlanningNotes`
- `NewTimeline`
- `NewStatus`

It also creates or refreshes formula-backed `Decision Staging` / `tblDecisionStaging` so `office-scripts/apply_notes.ts` can run its two-pass prepare/apply workflow without manual copy/paste. The worksheet control area states the required sequence: type in `P:R`, run `ApplyNotes` once to prepare, inspect `Decision Staging`, then run `ApplyNotes` again to apply. `ApplyNotes` updates the control area after each normal run with the last phase, result, and next action. A fresh setup seeds `Planning Review!P5:R5` when blank. The seeded smoke input targets `Sample over-projected work` in the starter `Planning Table`; run `ApplyNotes` once to stage it into `tblDecisionStaging` as `Prepared` while preserving `ReviewRow`-keyed helper formulas, inspect `ApplyMessage`, and run it a second time to apply it. For multi-row tests, enter values in more than one `Planning Review!P:R` row; each staged row carries its source `ReviewRow`, and duplicate staged writes to the same `Planning Table` row are blocked. If there are no current `Planning Review!P:R` inputs, a later script run resets stale staging rows to one blank formula-backed row.

## Automation Setup Sheet

The generated template includes an `Automation Setup` sheet because the public `.xltx` does not embed Office Scripts like VBA macros. The sheet points users to the `ApplyNotes.ts` release asset and gives the import sequence:

```text
Download ApplyNotes.ts -> Automate > New Script -> paste -> save as ApplyNotes
```

This keeps script installation explicit and tenant-controlled. The workbook has the staging tables and review surfaces; the operator chooses whether to import and run the optional writeback script.

`Setup Asset Workflow` is optional. It creates `tblAssets` plus the asset setup, mapping, change, and state-history tables used by `office-scripts/apply_asset_mappings.ts`; it is not part of the default setup path. It also applies dropdowns for asset state/status fields and advisory relationship dropdowns for asset IDs and project keys. Rerunning it recreates those workflow tables from headers, so use it as a starter/reset action before entering real asset rows or against a workbook copy.

The task-pane `Setup + Install + Validate + Outputs` button creates the public demo sheets as part of the full starter flow. The standalone `Insert Demo Outputs` button remains available for rerunning only the output insertion. Before either path writes the main report, it checks `Planning Review!A4:N200` and reports the first cell that would block the spill. It inserts the main report at `Planning Review!A4` and places the Analysis screens at `A4` on separate sheets named `BU Cap Scorecard`, `Reforecast Queue`, `PM Spend Report`, `Working Budget`, and `Burndown`. It also creates an `Internal Jobs` sheet at `A4` with `=Ready.InternalJobs_Export()` for readiness smoke testing.

## Add-In Option

The `addin/` folder provides an Office.js starter that can create the sheets, paste the starter data, install the named formulas, and validate the workbook contract from a task pane.

The normal `Setup + Install + Validate + Outputs` path includes notes workflow setup. Asset workflow setup remains opt-in from the standalone `Setup Asset Workflow` button.

See `docs/office_addin.md` for the packaging boundary, `docs/notes_apply_workflow.md` for the notes apply flow, and `docs/asset_setup_workflow.md` for optional asset setup.
