## 2026-04-26 - Key ApplyNotes staging by Planning Review row

Semantic change:

- Added `ReviewRow` to `tblDecisionStaging` so each prepared row is tied to the exact `Planning Review` row that supplied `P:R` inputs.
- Changed formula-backed Decision Staging helper columns to resolve through `ReviewRow` instead of row-position-specific formulas, which prevents Excel table autofill from duplicating the first staged source row across multiple rows.
- Blocked duplicate staged rows that would write to the same `Planning Table` row in one apply batch.
- Kept the two-pass flow: run once to prepare from `Planning Review!P:R`, inspect `ApplyMessage`, run again to apply.

Minimal diff summary:

- Updated `modules/notes.formula.txt`, `addin/taskpane.js`, and `office-scripts/apply_notes.ts`.
- Updated the decision staging starter header, notes/add-in/starter docs, Office Scripts README, and static audit coverage.

Visible impact:

- Workbook behavior: multiple distinct Planning Review input rows can stage independently; duplicate staged writes to the same Planning Table row are blocked with an operator-readable message.
- Formula logic: `Notes.FromArrayv` now carries `ReviewRow` for staging identity.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Tighten ApplyNotes messages and staging reset

Semantic change:

- Updated `ApplyNotes` messages so eligible rows say `Prepared`, successful rows say exactly what was applied, and unsafe rows use `Blocked` instead of looking prepared.
- Added `Skipped` handling for prepared rows that no longer have non-empty target values.
- Added a safe reset path: when there are no current `Planning Review!P:R` inputs and no prepared rows waiting to apply, `ApplyNotes` resets stale staging rows to one blank formula-backed row.
- Updated task-pane and docs guidance to state the two-pass operator flow: type in `Planning Review!P:R`, run once to prepare, inspect `ApplyMessage`, run again to apply.

Minimal diff summary:

- Updated `office-scripts/apply_notes.ts`.
- Updated add-in instructions in `addin/taskpane.html` and `addin/taskpane.js`.
- Updated notes/add-in/starter docs and static audit coverage.

Visible impact:

- Workbook behavior: blocked or unsafe rows now show clearer statuses and messages, and stale staging rows can be reset without losing the formula-backed staging design.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Seed Planning Review ApplyNotes smoke input

Semantic change:

- Changed the public-safe notes smoke path so test values originate in `Planning Review!P:R`, not by bypassing the review surface.
- Updated `Notes.FromArrayv` to carry `ExistingMeetingNotes`, `NewPlanningNotes`, `NewTimeline`, and `NewStatus` from `Planning Review!O:R`.
- Updated `Setup Notes Workflow` to seed `Planning Review!P5:R5` when blank and wire a single smoke row in `tblDecisionStaging` to `Notes.FromArrayv` formulas, so a fresh workbook can test ApplyNotes without manual row entry.
- Kept `BudgetMatchCount` scalar inside the table by using `SUMPRODUCT` over the matched `Planning Table` project-description column.
- Changed `ApplyNotes` run 1 to actively scan `Planning Review!P5:R200`, resize `tblDecisionStaging`, restore formula-backed review/context/helper columns, and mark rows `Prepared`.
- Updated `ApplyNotes` so successful applies clear the matching `Planning Review!P:R` source inputs instead of overwriting formula-backed staging input columns.
- Kept `samples/decision_staging_starter.tsv` as a public-safe expected staging-shape fixture.

Minimal diff summary:

- Updated `modules/notes.formula.txt`.
- Updated `office-scripts/apply_notes.ts`.
- Updated `samples/decision_staging_starter.tsv`.
- Updated `addin/taskpane.js`.
- Updated `office-scripts/README.md`.
- Updated notes/add-in/starter docs and static audit coverage.

Visible impact:

- Workbook behavior: fresh setup now includes one ready Planning Review smoke input; `ApplyNotes` run 1 stages current `P:R` inputs into formula-backed `tblDecisionStaging`, and run 2 applies prepared rows.
- Formula logic: `Notes.FromArrayv` now includes `ExistingMeetingNotes` alongside the new note/timeline/status inputs.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

2026-04-26 - Clarify optional asset setup UI

Semantic change:

- Clarified the task-pane standard setup path by logging that asset workflow setup remains optional after notes setup completes.
- Color-coded the `Setup Asset Workflow` button as optional.
- Forced workflow table header text to black after table creation so asset table headers remain readable on light header fills.

Minimal diff summary:

- Updated `addin/taskpane.html`, `addin/taskpane.css`, and `addin/taskpane.js`.
- Updated add-in docs and audit coverage.

Visible impact:

- Workbook behavior: optional asset table headers should render with black text; standard setup still does not run asset setup.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Promote asset workflow to tracker starter

Semantic change:

- Added a first-class `Asset Register` / `tblAssets` starter table to the optional asset workflow.
- Added asset-specific dropdown lists and advisory relationship dropdowns for asset IDs and project keys.
- Documented that `Setup Asset Workflow` is a starter/reset action because it recreates the workflow tables from headers.

Minimal diff summary:

- Updated `addin/taskpane.js` asset workflow table setup and validation data.
- Updated `office-scripts/apply_asset_mappings.ts` comments to state the `tblAssets` boundary.
- Updated asset/add-in/starter docs and added `docs/asset_tracker_next_steps.md`.
- Updated static audit coverage for the asset register, relationship dropdowns, and reset boundary.

Visible impact:

- Workbook behavior: optional asset setup now creates `tblAssets` and applies asset relationship dropdowns when selected.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Add in-add-in ApplyNotes script handoff

Semantic change:

- Added a task-pane `ApplyNotes` section so operators no longer need to manually browse the repo to find the script template.
- Added a one-click `Copy ApplyNotes Script` action that reads `../ApplyNotes`, displays the script text in the task pane, and copies it to the clipboard when host permissions allow it.
- Added a blocked-clipboard fallback that selects the visible script text inside the add-in.
- Added explicit in-pane import instructions for `Automate -> New Script` and added static audit coverage for the new handoff path.

Minimal diff summary:

- Updated `addin/taskpane.html`, `addin/taskpane.js`, and `addin/taskpane.css`.
- Updated `docs/office_addin.md`.
- Updated `tools/audit_capex_module.py`.

Visible impact:

- Workbook behavior: sideloaded add-in now includes an explicit ApplyNotes script setup path.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Tailor worktree roles to Excel workflow

Semantic change:

- Refined the Git worktree role model around Excel-formula repo work: formula/add-in implementation, workbook contract review, automated smoke/lint runs, and disposable workbook-reference analysis.
- Clarified that workbook copies remain local operator artifacts and scratch findings must be promoted only as sanitized text artifacts.

Minimal diff summary:

- Updated `docs/git_worktree_workflow.md`.
- Updated the README worktree starter example.
- Updated static audit coverage for the Excel-specific worktree role model.

Visible impact:

- Workbook behavior: no formula logic change.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Clarify Git worktree concurrency roles

Semantic change:

- Refined the Git worktree workflow around named concurrent roles: `main`, `work`, `review`, `fuzz`, and `scratch`.
- Kept branches as the source-control primitive and documented worktrees as task-concurrency folders.

Minimal diff summary:

- Updated `docs/git_worktree_workflow.md`.
- Updated the README worktree summary.
- Updated static audit coverage for the role model.

Visible impact:

- Workbook behavior: no formula logic change.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-26 - Add Git worktree workflow starter

Semantic change:

- Added a small Git worktree workflow guide for managing `main` as the stable product branch and `codex/*` branches as temporary task branches.
- Added a PowerShell helper that creates a sibling linked worktree from `origin/main`.

Minimal diff summary:

- Added `docs/git_worktree_workflow.md`.
- Added `tools/new_worktree.ps1`.
- Updated the README.

Visible impact:

- Workbook behavior: no formula logic change.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Compact installed formula names for workbook save

Semantic change:

- Changed the Office.js installer to compact formula whitespace and block comments before creating workbook defined names.
- Kept the source formula modules readable in the repository while reducing the workbook-installed formula text.
- Added audit coverage that checks compacted installed formula bodies stay within Excel's `8192` character save limit.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated `tools/audit_capex_module.py`.
- Updated add-in docs.

Visible impact:

- Workbook behavior: rerunning formula install should repair workbooks that fail to save because installed names are too long.
- Formula logic: no intended logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Keep Composite Cat manual and compute readiness output

Semantic change:

- Kept `Composite Cat` as a manual pre-formula planning-table helper for Excel sort, remove-duplicates, and Data > Subtotal workflows.
- Removed the source-table `Internal Ready` override column from the public starter contract.
- Made `Ready.InternalJobs_Export` emit computed `Internal Ready Final` directly from eligibility, maturity, stage, and chargeability.
- Shifted the starter workbook width from `A:BM` to `A:BL`.

Minimal diff summary:

- Updated `samples/planning_table_starter.tsv`, `addin/taskpane.js`, `modules/get.formula.txt`, and `modules/ready.formula.txt`.
- Updated starter, add-in, scenario, import-map, and structure-map docs.
- Updated static audit coverage for the 64-column starter contract and computed readiness output.

Visible impact:

- Workbook behavior: new starter workbooks expose `Composite Cat`, `Chargeable`, `Internal Eligible`, and `Canceled`, but no source-table `Internal Ready` override.
- Formula logic: `Ready.InternalJobs_Export` output changes by removing the raw `Internal Ready` column and computing `Internal Ready Final` directly.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Expand Planning Table structure map

Semantic change:

- Expanded the public-safe structure map from Yes/No fields to the full 65-column `Planning Table` contract.
- Documented each column's role, validation or formatting treatment, and primary formula/add-in dependencies.
- Kept the reference parse facts separate from the public starter contract.

Minimal diff summary:

- Updated `docs/planning_worksheet_structure_map.md`.
- Updated static audit coverage for the full-column map.

Visible impact:

- Workbook behavior: no intended change.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Remove Eligible fallback column

Semantic change:

- Removed the visible `Eligible` column from the public starter `Planning Table`.
- Made `Internal Eligible` the only readiness eligibility input in the starter contract.
- Removed the `Ready` fallback path that looked for a legacy `Eligible` column.
- Shifted the starter workbook width from `A:BN` to `A:BM`.

Minimal diff summary:

- Updated `samples/planning_table_starter.tsv`.
- Updated `addin/taskpane.js`, `modules/get.formula.txt`, and `modules/ready.formula.txt`.
- Updated starter, add-in, import-map, and structure-map docs.
- Updated static audit coverage for the 65-column starter contract and no visible `Eligible` fallback column.

Visible impact:

- Workbook behavior: new starter workbooks expose one eligibility flag, `Internal Eligible`, instead of both `Eligible` and `Internal Eligible`.
- Formula logic: `Ready.InternalEligible` now resolves `Internal Eligible` directly; older fallback behavior is intentionally removed.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Document Yes/No planning worksheet dependencies

Semantic change:

- Added a public-safe planning worksheet structure map derived from the reference parse.
- Listed the complete public starter Yes/No field set and their formula dependencies.
- Documented that the old explicit `Y,N` validation position is now handled by header-driven `Chargeable` validation.

Minimal diff summary:

- Added `docs/planning_worksheet_structure_map.md`.
- Linked the structure map from starter and import docs.
- Updated static audit coverage for the Yes/No dependency map.

Visible impact:

- Workbook behavior: no intended change.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Fold demo outputs into setup action

Semantic change:

- Changed the primary task-pane action from `Setup + Install + Validate` to `Setup + Install + Validate + Outputs`.
- The combined action now inserts the public demo output sheets after validation succeeds.
- Kept the standalone `Insert Demo Outputs` action for rerunning only output insertion.

Minimal diff summary:

- Updated `addin/taskpane.html` and `addin/taskpane.js`.
- Updated Office add-in and starter workbook docs.
- Updated smoke helper text and static audit coverage for the four-step action.

Visible impact:

- Workbook behavior: the primary setup button now creates output sheets automatically.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Add Internal Jobs demo sheet

Semantic change:

- Added an `Internal Jobs` demo sheet to the task-pane `Insert Demo Outputs` flow.
- The sheet places `=Ready.InternalJobs_Export()` at `A4` so readiness output can be smoke-tested like the Analysis screens.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated starter workbook, Office add-in, and scenario docs.
- Updated static audit coverage for the new demo sheet.

Visible impact:

- Workbook behavior: `Insert Demo Outputs` now creates an `Internal Jobs` sheet for the Ready export.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Remove JobFlag starter column

Semantic change:

- Removed `JobFlag` from the public starter `Planning Table` contract.
- Shifted the starter workbook width from `A:BO` to `A:BN`.
- Updated `Search` helper logic to find health-check inputs by header name instead of hardcoded row ordinals.

Minimal diff summary:

- Updated `samples/planning_table_starter.tsv`.
- Updated `addin/taskpane.js`, `modules/get.formula.txt`, and `modules/search.formula.txt`.
- Updated starter workbook, Office add-in, workbook import map, and scenario docs.
- Updated static audit coverage for the 66-column starter contract and absence of `JobFlag` from live starter/setup contracts.

Visible impact:

- Workbook behavior: new starter workbooks no longer include a `JobFlag` column, and `Search.Projects_Health` follows headers after the width change.
- Formula logic: `get` range bounds and `Search` helper lookup behavior changed; main report, Analysis, and Ready chargeability logic were not otherwise changed.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Stabilize Ready chargeability helper

Semantic change:

- Replaced the stale `Ready.JobFlag` column helper with `Ready.ChargeableFlag`.
- Made the public `Ready` range helpers find starter inputs by header name instead of hardcoded old-workbook columns.
- Made the example execution-stage list self-contained so `Ready.InternalReady3` no longer depends on an uncreated workbook list sheet.

Minimal diff summary:

- Updated `modules/ready.formula.txt`.
- Updated add-in required-name validation for the public `Ready` helpers.
- Updated starter workbook, Office add-in, workbook import map, and scenario docs.
- Updated static audit coverage for the `Ready.ChargeableFlag` contract and stale `Ready.JobFlag` removal.

Visible impact:

- Workbook behavior: `Ready` example outputs can now resolve the public `Chargeable` column by header and no longer treat public column `O` as chargeability.
- Formula logic: `Ready` helper logic changed; main report and Analysis formulas were not changed.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Clarify Chargeable and JobFlag readiness contract

Semantic change:

- Documented `Chargeable` as the canonical internal-labor chargeability flag for the public starter workbook.
- Documented `JobFlag` as a separate starter yes/no planning flag that formula modules do not currently consume.
- Added audit coverage so future changes keep the add-in data model, starter docs, import map, and scenario matrix aligned.

Minimal diff summary:

- Updated `addin/taskpane.js` row-validation metadata.
- Updated starter workbook, Office add-in, workbook import map, and scenario docs.
- Updated static audit coverage for the `Chargeable` versus `JobFlag` contract.

Visible impact:

- Workbook behavior: no dropdown behavior change; the starter still validates both public yes/no columns by header.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Centralize dropdown application data

Semantic change:

- Refactored add-in setup data into one `applicationData` model for sheets, dropdown lists, visible controls, and row-validation rules.
- Changed starter row dropdown setup to find validation targets by header name, including `Chargeable` rows `3:2000` with `Y,N`.
- Kept the spill-safe `Planning Review!B2:E2` control layout and `M2:N2` month controls.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated add-in, starter workbook, and scenario docs.
- Updated static audit coverage for the centralized dropdown contract.

Visible impact:

- Workbook behavior: dropdown setup becomes model-driven and `Chargeable` validation extends through row `2000`.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Guard main report demo spill range

Semantic change:

- Added a task-pane guard before `Insert Demo Outputs` writes the main report formula to `Planning Review!A4`.
- The guard checks `Planning Review!A4:N200` and reports the first cell that would block the main report spill.
- If `Planning Review!A4` already contains the expected main report formula and is not showing `#SPILL!`, the button remains safe to rerun.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated starter workbook and Office add-in docs.
- Updated static audit coverage for the pre-insert spill guard.

Visible impact:

- Workbook behavior: spill blockers now fail with a specific task-pane error before the main demo output is inserted.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Move visible controls above report spill

Semantic change:

- Moved the editable `Planning Review` control values from `K3:K6` to the top control band at `B2:E2`.
- Kept `M2:N2` as the report/defer as-of month cells because the formula modules read those addresses.
- The starter setup clears the old `J2:K6` panel so rerunning setup removes stale cells that can block the `A4` report spill.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated starter workbook, import-map, and Office add-in docs.
- Updated static audit coverage for the spill-safe control layout.

Visible impact:

- Workbook behavior: the main report formula at `Planning Review!A4` has an unobstructed `A:N` spill path.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Add demo output insertion action

Semantic change:

- Added a task-pane action that validates the starter workbook, creates demo output sheets, and inserts the implemented report formulas at fixed `A4` spill points.
- The action leaves setup and validation separate, so operators can choose when to create the demo output sheets.

Minimal diff summary:

- Updated `addin/taskpane.html` and `addin/taskpane.js`.
- Updated starter workbook and Office add-in docs.
- Updated static audit coverage for the demo-output action.

Visible impact:

- Workbook behavior: optional demo sheets can be created from the task pane.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Add task-pane validation summary

Semantic change:

- Added a compact task-pane status summary after workbook validation succeeds.
- The summary reports sheets present, workbook names installed, Planning Table header count, configured cap rows, visible controls, and dropdown lists.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated add-in docs and audit coverage.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Guard stale add-in dev server reuse

Semantic change:

- Made the Office.js smoke helper verify that port 3000 is serving the current checkout before Excel sideload starts.
- If another checkout is serving stale task-pane files on the smoke-test port, the helper stops that listener so the current repo can start its own dev server.

Minimal diff summary:

- Updated `tools/start_addin_smoke_test.ps1`.
- Updated audit coverage for stale-server detection.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Add starter workbook UX setup

Semantic change:

- Extended the Office.js starter so a blank workbook gets formatted starter sheets, visible planning controls, dropdown-backed validation lists, and stronger workbook validation.
- Rebound unqualified workbook-control names to visible cells on `Planning Review` after formula installation while leaving module-qualified `Controls.*` defaults intact.
- Documented the starter workbook layout, control cells, validation-list sheet, and preserved output ranges.

Minimal diff summary:

- Updated `addin/taskpane.js`.
- Updated starter workbook, Office add-in, import-map, operating-contract, changelog, and audit documentation checks.
- Updated `tools/audit_capex_module.py` to enforce the visible-control setup and validation contract.

Visible impact:

- Workbook behavior: controls become visible and dropdown-backed in the starter workbook.
- Formula logic: no formula module change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Document implemented planning-screen inventory

Semantic change:

- Brought the public planning-plugin menu in line with the `Analysis` formulas already present in the template.
- Extended the static audit and add-in validator so implemented planning screens remain importable and documented.
- Added scenario coverage for PM spend, working-budget, and burndown screens.

Minimal diff summary:

- Updated `docs/planning_plugins.md`, `docs/scenario_matrix.md`, and `docs/starter_workbook.md`.
- Updated `addin/taskpane.js` to validate the implemented Analysis entry points.
- Updated `tools/audit_capex_module.py` to require the Analysis formulas, docs, scenarios, and changelog entry.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Guard starter report subtotal errors

Semantic change:

- Guarded the main report's hidden-burn and BU cap subtotal lookups so public starter data does not surface `#VALUE!` in subtotal flag or remaining columns.
- Empty hidden-burn groupings now behave as zero hidden burn.

Minimal diff summary:

- Updated `modules/capital_planning_report.formula.txt`.
- Added audit coverage for hidden-burn and BU-cap fallback logic.

Visible impact:

- Workbook behavior: starter subtotal rows should render clean flag and remaining values.
- Main report totals: no intended change except replacing error cells with zero-fallback subtotal math.
- Subtotal flags: can change from `#VALUE!` to a blank or cap flag where applicable.
- Cap remaining values: can change from `#VALUE!` to the calculated remaining cap.

## 2026-04-25 - Add workbook-control defaults

Semantic change:

- Added a tracked `Controls` formula module for default workbook-control names that public blank workbooks otherwise reported as `#NAME?`.
- The Office.js installer now creates defaults for `PM_Filter_Dropdowns`, `Future_Filter_Mode`, `HideClosed_Status`, and `Burndown_Cut_Target`.

Minimal diff summary:

- Added `modules/controls.formula.txt`.
- Updated the add-in installer, import map, add-in docs, changelog, and audit coverage.

Visible impact:

- Workbook behavior: report filters now resolve to safe defaults in a clean workbook.
- Main report totals: no intended change when defaults match the previous workbook controls.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Bound local add-in server stalls

Semantic change:

- Added TCP, TLS, and stream timeouts to the local Office.js dev server so one stalled local request cannot block the smoke-test host.
- Extended audit coverage for the dev-server timeout guard.

Minimal diff summary:

- Updated `tools/start_addin_dev_server.ps1`, `tools/audit_capex_module.py`, and this changelog.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Technical review guide added

Semantic change:

- Added a public technical-review guide so the repository communicates the governed-workbook systems pattern clearly to reviewers.
- Surfaced the guide from the README without changing formula modules or workbook behavior.
- Extended audit coverage so the review guide remains present and tied to the public/private boundary.

Minimal diff summary:

- Added `docs/technical_review_guide.md`.
- Updated README reviewer guidance and repository layout.
- Updated `tools/audit_capex_module.py` documentation checks.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Promote workbook-local compatibility helpers

Semantic change:

- Added `TRIMRANGE_KEEPBLANKS` and `RBYROW` to the tracked formula modules so clean workbooks do not depend on hidden workbook-local LAMBDAs.
- Extended the Office.js validator and static audit to require those compatibility helpers.

Minimal diff summary:

- Updated `modules/get.formula.txt`, `modules/kind.formula.txt`, `addin/taskpane.js`, docs, and audit coverage.

Visible impact:

- Workbook behavior: existing formulas should resolve in a clean workbook after module install.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Fix Office manifest validation

Semantic change:

- Updated the Office add-in manifest so Microsoft manifest validation accepts the local sideload package.
- Replaced SVG manifest icons with PNG icons and raised the manifest version to `1.0.0.0`.

Minimal diff summary:

- Updated `addin/manifest.xml`.
- Replaced add-in SVG icon placeholders with PNG icon assets.
- Extended audit coverage for manifest version and icon extensions.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Add npm smoke-test package metadata

Semantic change:

- Added minimal Node package metadata so Microsoft Office add-in debugging tools can run from the repo root.
- Kept npm as local tooling only; workbook formulas remain the calculation engine.
- Made add-in smoke helpers find the standard Windows Node install even before a shell PATH refresh.

Minimal diff summary:

- Added `package.json` with add-in smoke, server, stop, and dev-server scripts.
- Updated README, Office add-in docs, and audit coverage for the npm smoke-test path.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Automated add-in smoke-test helper

Semantic change:

- Added Windows PowerShell helpers for running the local Office.js smoke test with less manual setup.
- Made the server-only path PowerShell-native so it does not require Node/npm.
- Kept the add-in as a setup and validation layer; workbook formulas remain the calculation engine.

Minimal diff summary:

- Added `tools/start_addin_smoke_test.ps1`, `tools/start_addin_dev_server.ps1`, and `tools/stop_addin_smoke_test.ps1`.
- Updated README and Office add-in docs with the one-command smoke-test path.
- Extended audit coverage for the add-in helper scripts.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Office.js add-in starter

Semantic change:

- Added a minimal Excel Office.js task-pane add-in scaffold for installing formula modules and starter sheets.
- Kept formula modules as the calculation engine; JavaScript is only the packaging, setup, and validation layer.

Minimal diff summary:

- Added `addin/manifest.xml`, task pane HTML/CSS/JS, and text SVG icon placeholders.
- Added `docs/office_addin.md`.
- Updated README, starter workbook docs, operating contract, and audit coverage.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Public push helper added

Semantic change:

- Added a local public-export helper that runs validation, commits, rebases, and pushes the public repo.

Minimal diff summary:

- Added `tools/push_public.ps1`.
- Documented the helper in the README and public release checklist.
- Extended audit coverage so the helper continues to run audit, formula lint, whitespace check, fetch/rebase, and push.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Workbook-driven cap setup

Semantic change:

- Moved public cap setup from module constants to a workbook input contract.
- Renamed the dated main-report module to the generic `CapitalPlanning` contract and `CAPITAL_PLANNING_REPORT()` entry point.
- Replaced the public projection header with `Annual Projected`.

Minimal diff summary:

- Added `samples/cap_setup_starter.tsv`.
- Updated `docs/starter_workbook.md`, README, and the import map to explain `Cap Setup`.
- Updated `kind.CapTable`, `kind.PortfolioCap`, and `kind.CapByBU(...)` to read from `Cap Setup`.
- Updated docs and audit checks so old dated tracker wording and hardcoded cap arrays do not return.

Visible impact:

- Main report totals: can change only when workbook `Cap Setup` values differ from previous constants.
- Subtotal flags: cap-related subtotal flags can change only from changed `Cap Setup` values.
- Cap remaining values: now come from workbook cap inputs rather than module constants.
- Candidate logic and planning-screen math: no intended change.

## 2026-04-25 - Excel runtime rationale documented

Semantic change:

- Added public-facing rationale for using Excel as the runtime while keeping workbook logic governed through text modules, Git, docs, and audit.
- Clarified that Excel is appropriate for planning review surfaces, not for transactional application requirements.

Minimal diff summary:

- Added `Why Excel` to the README.
- Added `Runtime Position` to the operating contract.
- Added audit coverage for the public rationale.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Search helper label cleanup

Semantic change:

- Removed legacy source/job shorthand wording from the search helper and replaced it with public `Job ID` wording.

Minimal diff summary:

- Updated `Search.Projects_Health` messages and local variable names.
- Added audit coverage to forbid legacy source/job shorthand wording.

Visible impact:

- Workbook behavior: health-message wording changed only.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.

## 2026-04-25 - Final public BU cleanup

Semantic change:

- Replaced work-looking BU codes and cap amounts with fictional public sample values.
- Added audit coverage so the removed BU codes and cap amounts cannot return silently.

Minimal diff summary:

- Updated the sample BU cap definitions.
- Updated starter table BU sample values to `BU-A: Sample Unit` and `BU-B: Sample Unit`.
- Extended public-safety audit checks.

Visible impact:

- Workbook behavior: public sample cap values changed to fictional placeholders.
- Main report totals: can change when using only the public sample workbook data.
- Subtotal flags: can change when using only the public sample workbook data.
- Cap remaining values: can change when using only the public sample workbook data.
- No private workbook should import this public sample cap table without replacing the fictional placeholders.

## 2026-04-25 - Starter workbook table added

Semantic change:

- Added a paste-ready public starter table so new users can create a blank workbook trial without inventing the source-table shape.
- Documented why the current finance block needs three columns per month.

Minimal diff summary:

- Added `samples/planning_table_starter.tsv`.
- Added `docs/starter_workbook.md`.
- Updated README and workbook import map with the starter flow.

Visible impact:

- Workbook behavior: no formula logic change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.
- New users get a concrete `Planning Table` shape for local testing.

## 2026-04-25 - Public release hardening second pass

Semantic change:

- Generalized remaining workbook-specific labels to public template names across formula modules, the staged-decision script, docs, and validation tooling.
- Added a public release checklist and strengthened the audit so old private workbook labels, local paths, URLs, email addresses, workbook binaries, and generated artifacts are blocked before export.

Minimal diff summary:

- Replaced old sheet/table/header vocabulary with `Planning Table`, `Planning Review`, `Decision Staging`, `Source ID`, `Job ID`, `Planning Notes`, and `Timeline`.
- Added `docs/public_release_checklist.md`.
- Extended `tools/audit_capex_module.py` with broader public-safety checks.

Visible impact:

- Workbook behavior: label contract changes only for the public template.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.
- Public release readiness is now checked directly by the audit.

## 2026-04-25 - Public template sanitization started

Semantic change:

- Converted the repo presentation from a private workbook workspace into a public-safe Excel formula-module template.
- Kept formula modules available as examples while removing public docs that named real workbook paths, workbook files, or organization-specific process details.

Minimal diff summary:

- Rewrote the README, operating contract, import map, planning-plugin menu, scenario matrix, and change log around the generic governed-formula-module pattern.
- Replaced workbook-specific lineage and inventory docs with generic public-template guidance.
- Replaced the static audit with a public-safety and formula-contract audit.

Visible impact:

- Workbook behavior: no intended change.
- Main report totals: no intended change.
- Subtotal flags: no intended change.
- Cap remaining values: no intended change.
- Public-template docs now describe the reusable pattern rather than a private workbook implementation.
