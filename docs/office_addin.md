# Office.js Add-In Starter

This repo includes a minimal Excel Office.js task-pane add-in under `addin/`.

The add-in is an installer and validator. It does not replace the formula modules with JavaScript business logic.

## What It Does

- Creates the starter sheets: `Planning Table`, `Cap Setup`, and `Planning Review`.
- Creates the `Validation Lists` sheet for dropdown sources.
- Pastes the public starter TSV data into `Planning Table!A2` and `Cap Setup!A2`.
- Formats starter headers, freezes top rows, applies currency formats, and adds non-negative cap validation.
- Adds visible `Planning Review` controls in `B2:E2`, with as-of month cells in `M2:N2`.
- Uses one `applicationData` model for starter sheets, dropdown source lists, control bindings, and row-validation rules.
- Applies dropdown validation for month, group, future-filter, closed-row, status, and yes/no fields, including header-driven `Chargeable` validation on `Planning Table`.
- Validates the public `Ready` helpers, including the header-driven `Ready.ChargeableFlag` name used by example readiness exports.
- Reads `modules/*.formula.txt` from the hosted repo root.
- Installs workbook defined names through the Excel JavaScript API.
- Compacts formula whitespace and comments before creating workbook names so saved `.xlsx` files stay under Excel's named-formula length limit.
- Installs default workbook-control names such as `PM_Filter_Dropdowns`, `Future_Filter_Mode`, `HideClosed_Status`, and `Burndown_Cut_Target`.
- Rebinds the unqualified workbook-control names to the visible `Planning Review` cells after module installation.
- Adds module-qualified names such as `kind.CapByBU` and `Analysis.REFORECAST_QUEUE`.
- Adds unqualified compatibility aliases for the first occurrence of each formula name.
- Validates required sheets, names, starter header order, cap setup shape, visible control values, bound control names, row-validation headers, and compatibility helpers such as `TRIMRANGE_KEEPBLANKS` and `RBYROW`.
- Prints a validation summary showing sheets present, workbook names installed, header count, configured cap rows, bound controls, dropdown lists, and row-validation rules.
- Inserts demo output formulas into predictable review sheets so a reviewer can inspect the implemented screens without typing formula names.
- Runs `Setup Notes Workflow` as part of the normal `Setup + Install + Validate + Outputs` path, creating `Planning Review!O:R` notes columns and `Decision Staging` / `tblDecisionStaging`.
- Provides a standalone `Setup Asset Workflow` button for optional asset sheets and tables; asset setup is not run from the default path.

## Local Trial Shape

Run the smoke-test helper from the repo root:

```powershell
.\tools\start_addin_smoke_test.ps1
```

With Node/npm installed, this equivalent npm script is also available:

```powershell
npm run addin:smoke
```

The helper:

- runs the static repo checks,
- creates/reuses a local trusted certificate for the server-only fallback,
- starts the local HTTPS server on `localhost` port `3000`,
- asks Excel desktop to sideload `addin/manifest.xml` when npm is available.

After Excel opens, use the task pane button:

```text
Setup + Install + Validate + Outputs
```

The combined button creates the starter sheets, installs formulas, validates the workbook contract, checks the main report spill range, and inserts the demo output formulas. When validation succeeds, the task pane status area includes a compact validation summary:

```text
Validation summary:
- Sheets present
- Workbook names installed
- Planning Table headers
- Cap Setup rows with BU
- Visible controls bound
- Dropdown lists ready
- Row validations configured
```

The standalone output button remains available when you want to rerun only the output-sheet insertion:

```text
Insert Demo Outputs
```

That button validates the workbook first, checks `Planning Review!A4:N200` for cells that would block the main report spill, then places demo formulas at fixed locations. If `Planning Review!A4` already contains the expected main report formula and is not showing `#SPILL!`, the button is safe to rerun.

The normal setup path also runs:

```text
Setup Notes Workflow
```

That action creates the notes helper/input columns beside the report and refreshes `Decision Staging` / `tblDecisionStaging` for the controlled ApplyNotes script. See `docs/notes_apply_workflow.md`.

The optional asset workflow is separate:

```text
Setup Asset Workflow
```

That action creates the asset setup sheets and tables only when selected. It is intentionally not part of `Setup + Install + Validate + Outputs`. See `docs/asset_setup_workflow.md`.

| Sheet | Cell | Formula |
|---|---|---|
| `Planning Review` | `A4` | `=CapitalPlanning.CAPITAL_PLANNING_REPORT()` |
| `BU Cap Scorecard` | `A4` | `=Analysis.BU_CAP_SCORECARD()` |
| `Reforecast Queue` | `A4` | `=Analysis.REFORECAST_QUEUE()` |
| `PM Spend Report` | `A4` | `=Analysis.PM_SPEND_REPORT()` |
| `Working Budget` | `A4` | `=Analysis.WORKING_BUDGET_SCREEN()` |
| `Burndown` | `A4` | `=Analysis.BURNDOWN_SCREEN()` |
| `Internal Jobs` | `A4` | `=Ready.InternalJobs_Export()` |

When the test session is done, run:

```powershell
.\tools\stop_addin_smoke_test.ps1
```

or:

```powershell
npm run addin:stop
```

If npm is not installed, or the Office debugging tool is blocked on a machine, use the server-only helper and sideload the manifest manually:

```powershell
.\tools\start_addin_dev_server.ps1
```

The manifest points Excel to:

```text
addin/taskpane.html
```

The task pane reads formula modules and samples by relative path, so it needs the full repo content available from the same hosted root.

## Starter Workbook Layout

The setup path is intentionally small and inspectable:

- `Planning Table` starts at `A2`, freezes the top two rows, formats the 64-column starter contract, and adds model-driven dropdowns for common status and yes/no fields.
- The `Chargeable` dropdown is applied by finding the `Chargeable` header on row `2`, then validating rows `3:2000` against `Y,N`.
- `Chargeable` is the chargeability input used by the `Search` and `Ready` helper modules. `Internal Eligible` is the readiness eligibility input used by `Ready.InternalEligible`. `Ready.InternalJobs_Export` computes `Internal Ready Final`; there is no source-table `Internal Ready`, no `JobFlag` starter column, and no separate visible `Eligible` fallback column.
- `Composite Cat` remains a manual pre-formula helper for operator sorting, dedupe, and Excel Data > Subtotal workflows.
- `Cap Setup` starts at `A2`, formats `Cap` as currency, and validates caps as non-negative numbers.
- `Planning Review` uses `B2:E2` for visible controls, `M2:N2` for month controls, leaves `A4:N200` open for the main report spill, and leaves `O4:R200` open for note examples.
- `Planning Review!O:R` is used by the notes workflow: `ExistingMeetingNotes`, `NewPlanningNotes`, `NewTimeline`, and `NewStatus`.
- `Decision Staging` stores `tblDecisionStaging`, the controlled staging table consumed by `office-scripts/apply_notes.ts`.
- `Validation Lists` stores the dropdown values used by the starter workbook.
- Demo output sheets are created by the combined `Setup + Install + Validate + Outputs` action, or by the standalone `Insert Demo Outputs` rerun action.
- Optional asset setup creates `Asset Setup`, `Project Asset Map`, `Semantic Assets`, `Asset Changes`, and `Asset State History` with asset workflow tables.

The unqualified control names are rebound to the visible cells:

```text
PM_Filter_Dropdowns -> 'Planning Review'!$B$2
Future_Filter_Mode -> 'Planning Review'!$C$2
HideClosed_Status -> 'Planning Review'!$D$2
Burndown_Cut_Target -> 'Planning Review'!$E$2
```

## Boundary

The add-in is not a workbook binary, not VBA, and not a calculation engine.

The calculation logic still lives in native Excel named formulas after installation. This keeps the public story aligned with governed formula modules rather than a hidden JavaScript planning engine.

## Production Notes

Before using this as a production add-in:

- replace the local development host in `addin/manifest.xml`,
- decide whether the add-in is internal-only or public Marketplace/AppSource material,
- add real icon assets if required by the deployment channel,
- test sideloading in desktop Excel and Excel for the web,
- keep formula module import validation in `tools/audit_capex_module.py`.
