# Change Log

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
