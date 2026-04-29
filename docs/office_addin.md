# Office.js Add-In Starter

This repo includes a minimal Excel Office.js task-pane add-in under `addin/`.

The add-in is an installer and validator. It does not replace the formula modules with JavaScript business logic.

For a new workbook artifact, the preferred path is now the generated starter template from `tools/build_governance_starter_workbook.ps1`. The default generated `.xltx` is planning-only. Use `-Edition AssetsLite` for a visible `Asset Hub`, `-Edition AssetsFull` for visible `Asset Hub` plus `Asset Finance Hub`, and `-Edition SemanticTwin` for the optional REC/Brick semantic crosswalk.

SemanticTwin is generated through the workbook builder rather than a new default add-in workflow. The add-in remains focused on setup, formula installation, validation, and planning/asset workflow helpers; it does not import full ontologies or complete graph/digital-twin export.

The generated template also includes a hidden `Automation Setup` worksheet. That sheet explains that `ApplyNotes.ts` is an optional Office Script release asset and must be imported through Excel `Automate -> New Script` before the notes writeback automation can run.

## What It Does

- Creates the starter sheets: `Start Here`, `Planning Table`, `Cap Setup`, `Data Import Setup`, `PQ Budget Input`, `PQ Budget QA`, `Planning Review`, `Source Status`, `Analysis Hub`, `Asset Hub`, `Asset Finance Hub`, and `Workbook Manifest`.
- Creates canonical data import tables: `tblDataSourceProfile`, `tblBudgetImportParameters`, `tblBudgetImportContract`, `tblBudgetInput`, `tblBudgetImportStatus`, and `tblBudgetImportIssues`.
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
- Validates required sheets, names, starter header order, `tblBudgetInput` header order, cap setup shape, visible control values, bound control names, row-validation headers, and compatibility helpers such as `TRIMRANGE_KEEPBLANKS` and `RBYROW`.
- Prints a validation summary showing sheets present, workbook names installed, header count, configured cap rows, bound controls, dropdown lists, and row-validation rules.
- Inserts demo output formulas into predictable hub sheets so a reviewer can inspect the implemented screens without typing formula names.
- Applies sheet visibility rules using `Excel.SheetVisibility.visible` and `Excel.SheetVisibility.hidden`, leaving backend sheets hidden but unhideable for troubleshooting. `tblWorkbookManifest[Presence]` marks generated sheets separately from `OptionalLegacy` sheet names.
- Provides an `ApplyNotes` setup helper in the task pane that loads the script template from `../office-scripts/apply_notes.ts`, copies it when clipboard access is available, displays the script text when clipboard access is blocked, and shows the exact Excel `Automate -> New Script` import step.
- Runs `Setup Notes Workflow` as part of the normal `Setup + Install + Validate + Outputs` path, creating a visible `Planning Review!O1:R3` `ApplyNotes Control` area, creating `Planning Review!O:R` notes columns, seeding public-safe `Planning Review!P5:R5` smoke input when blank, and creating formula-backed `Decision Staging` / `tblDecisionStaging` for the ApplyNotes first-pass staging step keyed by `ReviewRow`.
- Provides a standalone `Setup Asset Workflow` button for optional asset sheets, `tblAssets`, mapping/change/history tables, and asset relationship dropdowns; asset setup is not run from the default path.
- Leaves asset evidence Power Query import to the generated seed workbook plus PowerShell installer on this branch; the task pane does not expose duplicate M-template copy buttons.

Asset workflow is optional. Start with `Asset Hub` only when project-to-asset tracking is needed. Do not start with PQ asset evidence sheets. Do not start with `Asset State History`. Asset Finance is advanced and requires classified evidence.

## Local Trial Shape

For operator-style local use after downloading the repo ZIP, run the safer launcher from the repo root:

```powershell
.\Start-AddIn.ps1
```

That launcher confirms you are using a workbook copy, installs npm dependencies when needed, then delegates to the smoke-test helper that starts the add-in and launches Excel. It does not edit a workbook by itself.

For developer smoke testing, run the smoke-test helper directly:

```powershell
.\tools\start_addin_smoke_test.ps1
```

With Node/npm installed, this equivalent npm script is also available:

```powershell
npm run addin:smoke
npm run test:smoke
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

The combined button creates the starter sheets, installs formulas, validates the workbook contract, checks the main report spill range, applies the simplified workbook visibility rules, and inserts the demo hub formulas. When validation succeeds, the task pane status area includes a compact validation summary:

```text
Validation summary:
- Sheets present
- Workbook names installed
- Planning Table headers
- tblBudgetInput headers
- Cap Setup rows with BU
- Visible controls bound
- Dropdown lists ready
- Row validations configured
```

The standalone output button remains available when you want to rerun only the output-sheet insertion:

```text
Insert Demo Outputs
```

That button validates the workbook first, checks `Planning Review!A4:N200` for cells that would block the main report spill, then places demo formulas at fixed hub locations. It writes `Planning Review` with `=CapitalPlanning.CAPITAL_PLANNING_REPORT()` and uses the hub sheets for the stacked review outputs, including `Analysis Hub` sections for Burndown with `=Analysis.BURNDOWN_SCREEN()` and Internal Jobs with `=Ready.InternalJobs_Export()`. Each hub starts with a clickable `Go to section` table. If `Planning Review!A4` already contains the expected main report formula and is not showing `#SPILL!`, the button is safe to rerun.

The task pane also includes this helper:

```text
Copy ApplyNotes Script
```

That helper loads the `ApplyNotes` template from the repo into the task pane, copies it to the clipboard when the host allows clipboard access, and leaves the script text visible in the pane if clipboard access is blocked. In Excel, use `Automate -> New Script`, replace the default code, then save the script as `ApplyNotes`. Operators then use the script in two passes: run once to prepare `Decision Staging` from `Planning Review!P:R`, inspect `ApplyMessage`, and run again to apply prepared rows.

The normal setup path also runs:

```text
Setup Notes Workflow
```

That action creates the notes helper/input columns beside the report and refreshes `Decision Staging` / `tblDecisionStaging` for the controlled ApplyNotes script. It also writes an `ApplyNotes Control` area at `Planning Review!O1:R3` so the worksheet itself states: type in `P:R`, run once to prepare, inspect `Decision Staging`, then run again to apply. `ApplyNotes` updates that control area after each normal run with the last phase, result, and next action. It seeds `Planning Review!P5:R5` on a fresh workbook, so ApplyNotes run 1 can stage a test row without manual typing. Multi-row staging uses `ReviewRow` to keep each staged row tied to the exact `Planning Review` source row, and duplicate staged writes to the same `Planning Table` row are blocked. A later script run with no current `Planning Review!P:R` inputs resets stale staging rows to a blank formula-backed row. See `docs/notes_apply_workflow.md`.

The optional asset workflow is separate:

```text
Setup Asset Workflow
```

That action creates the asset setup sheets and tables only when selected. It is intentionally optional and not run from the default path. The task pane color-codes the asset setup button as optional, and the standard setup completion message states that asset workflow setup is still separate. It creates `Asset Register` / `tblAssets`, `Asset Setup`, `Project Asset Map`, `Semantic Assets`, `Asset Changes`, and `Asset State History`. Rerunning it recreates the asset workflow tables from headers, so treat it as a starter/reset action on a workbook copy or before entering live asset rows. See `docs/asset_setup_workflow.md`.

Asset Evidence Power Query is intentionally outside the Office.js task pane on this branch. For new workbook starts, run `tools/build_governance_starter_workbook.ps1` and use `release_artifacts/governance-starter/Governance_Starter.xltx`. For a button-driven local install into an existing workbook copy, run `tools/start_asset_evidence_pq_installer.ps1`; for automation, run `tools/install_asset_evidence_pq_workbook.ps1` against a workbook copy. The installed sheets include `Asset Evidence Setup` with `tblAssetEvidenceSource`, `tblAssetEvidenceRules`, and `tblAssetEvidenceOverrides`, plus loaded output tables for `qAssetEvidence_Normalized`, `qAssetEvidence_Classified`, `qAssetEvidence_Linked`, `qAssetEvidence_Status`, `qAssetEvidence_ModelInputs`, and `qQA_AssetEvidence_MappingQueue`. See `docs/asset_evidence_power_query.md`.

The generated starter also installs the v0.4 asset finance bridge outside the Office.js task pane. It creates hidden `Asset Finance Setup` / `tblAssetFinanceAssumptions`, installs `AssetFinance` names from `modules/asset_finance.formula.txt`, and places the depreciation, funding requirements, totals, and chart-ready feeds on `Asset Finance Hub`. Those formulas consume `tblAssetEvidence_ModelInputs` only, and mapped-only evidence does not drive final finance outputs.

The v0.5 data import bridge is part of both generated starter setup and blank-workbook add-in setup. `Planning Table` remains the manual starter source and local writeback surface, but formula modules consume `tblBudgetInput` as the canonical formula source. After `ApplyNotes.ts` or manual edits update `Planning Table`, refresh or re-sync the current-workbook budget adapter before relying on formula outputs that read `tblBudgetInput`.

The add-in-created blank workbook path now creates both `tblPlanningTable` and `tblBudgetInput`, so the current-workbook Power Query adapter has the same source table contract as the generated starter.

| Sheet | Cell | Formula |
|---|---|---|
| `Planning Review` | `A4` | `=CapitalPlanning.CAPITAL_PLANNING_REPORT()` |
| `Source Status` | `A4` | `=Source.SOURCE_STATUS()` |
| `Analysis Hub` | `A14` | `=Analysis.BU_CAP_SCORECARD()` |
| `Analysis Hub` | `A52` | `=Analysis.REFORECAST_QUEUE()` |
| `Analysis Hub` | `A114` | `=Analysis.PM_SPEND_REPORT()` |
| `Analysis Hub` | `A176` | `=Analysis.WORKING_BUDGET_SCREEN()` |
| `Analysis Hub` | `A238` | `=Analysis.BURNDOWN_SCREEN()` |
| `Analysis Hub` | `A300` | `=Ready.InternalJobs_Export()` |

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
- `Start Here` is the active sheet after setup and explains the left-to-right operator flow, source rule, visible-sheet navigation, and hidden backend/admin sheets.
- `Data Import Setup` starts the public-safe source profile and 64-column import contract.
- `PQ Budget Input` starts hidden `tblBudgetInput` at `A1`; formula modules read this table.
- `PQ Budget QA` stores hidden `tblBudgetImportStatus` and `tblBudgetImportIssues`.
- The `Chargeable` dropdown is applied by finding the `Chargeable` header on row `2`, then validating rows `3:2000` against `Y,N`.
- `Chargeable` is the chargeability input used by the `Search` and `Ready` helper modules. `Internal Eligible` is the readiness eligibility input used by `Ready.InternalEligible`. `Ready.InternalJobs_Export` computes `Internal Ready Final`; there is no source-table `Internal Ready`, no `JobFlag` starter column, and no separate visible `Eligible` fallback column.
- `Composite Cat` remains a manual pre-formula helper for operator sorting, dedupe, and Excel Data > Subtotal workflows.
- `Cap Setup` starts at `A2`, formats `Cap` as currency, and validates caps as non-negative numbers.
- `Planning Review` uses `B2:E2` for visible controls, `M2:N2` for month controls, leaves `A4:N200` open for the main report spill, uses `O1:R3` for `ApplyNotes Control`, and leaves `O4:R200` open for note examples.
- Hidden `Automation Setup` explains the public release boundary for Office Scripts: download `ApplyNotes.ts`, open `Automate -> New Script`, paste the script, save it as `ApplyNotes`, then run the two-pass workflow when writeback is wanted.
- `Planning Review!O:R` is used by the notes workflow: `ExistingMeetingNotes`, `NewPlanningNotes`, `NewTimeline`, and `NewStatus`.
- Hidden `Decision Staging` stores formula-backed `tblDecisionStaging`, the controlled staging table consumed by `office-scripts/apply_notes.ts`; ApplyNotes run 1 resizes it from `Planning Review!P:R` while preserving `ReviewRow`-keyed review/context/helper formulas.
- Hidden `Validation Lists` stores the dropdown values used by the starter workbook.
- `Source Status`, `Analysis Hub`, `Asset Hub`, and `Asset Finance Hub` replace the old scattered demo output sheets. The stacked hub sheets include clickable section tables at the top.
- Optional asset setup creates `Asset Register`, `Asset Setup`, `Project Asset Map`, `Semantic Assets`, `Asset Changes`, and `Asset State History` with `tblAssets`, asset staging, mapping, change, and state-history tables.
- Asset setup also writes dropdown-backed validation lists for asset status, condition, criticality, change type, asset state, promotion status, mapping status, and advisory relationship lists for `Asset ID` and `Project Key`.
- Optional asset evidence setup is handled by the generated governance starter template, the generated seed workbook, and the PowerShell installer, not by task-pane copy buttons.

The unqualified control names are rebound to the visible cells:

```text
PM_Filter_Dropdowns -> 'Planning Review'!$B$2
Future_Filter_Mode -> 'Planning Review'!$C$2
HideClosed_Status -> 'Planning Review'!$D$2
Burndown_Cut_Target -> 'Planning Review'!$E$2
```

## Boundary

The add-in is not a workbook binary, not VBA, and not a calculation engine. The optional asset evidence Power Query seed is generated separately from source-controlled M templates for workbook-copy installs.

The calculation logic still lives in native Excel named formulas after installation. This keeps the public story aligned with governed formula modules rather than a hidden JavaScript planning engine.

## Production Notes

Before using this as a production add-in:

- replace the local development host in `addin/manifest.xml`,
- decide whether the add-in is internal-only or public Marketplace/AppSource material,
- add real icon assets if required by the deployment channel,
- test sideloading in desktop Excel and Excel for the web,
- keep formula module import validation in `tools/audit_capex_module.py`.
