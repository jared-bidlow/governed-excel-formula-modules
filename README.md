# Governed Excel Formula Modules

This repository is a public template for treating complex Excel workbook logic as source code.

It keeps workbook formulas in plain-text modules, pairs them with scenario documentation, and validates the module set with static audit tools. The pattern is designed for capital planning and forecast-review workbooks, but the repo does not include any real workbook, company data, or generated Excel artifacts.

## What This Demonstrates

- Excel `LAMBDA` / `LET` modules tracked as text.
- Dynamic-array planning screens built with functions such as `GROUPBY`, `PIVOTBY`, `FILTER`, `SORTBY`, `HSTACK`, and `VSTACK`.
- A workbook-binary boundary: formulas, docs, and tests are versioned; `.xlsx` files are not.
- Static checks for formula balance, named-formula size, public-safety strings, and important planning-screen contracts.
- Operator-facing docs that explain how a workbook user should import, review, and validate formula modules.

## For Technical Reviewers

Start with `docs/technical_review_guide.md` if you are reviewing this repo as evidence of systems work.

That guide explains how the pieces fit together: source-controlled formula modules, workbook contracts, static audits, starter data, Office.js installation tooling, and the public/private boundary. It is written for someone who wants to understand the engineering pattern without needing a private workbook or production data.

For the durable architecture map, see `docs/reference_architecture_tree.md`.

## Layout

```text
governed-excel-formula-modules/
+-- AGENTS.md
+-- README.md
+-- README_FIRST.md
+-- Start-AddIn.ps1
+-- addin/
|   +-- manifest.xml
|   +-- taskpane.html
|   +-- taskpane.js
|   \-- assets/
+-- modules/
|   +-- get.formula.txt
|   +-- kind.formula.txt
|   +-- capital_planning_report.formula.txt
|   +-- assets.formula.txt
|   +-- analysis.formula.txt
|   \-- supporting workbook modules
+-- docs/
|   +-- asset_setup_workflow.md
|   +-- asset_quick_start.md
|   +-- asset_evidence_power_query.md
|   +-- copilot_review_playbook.md
|   +-- database_import_contract.md
|   +-- notes_apply_workflow.md
|   +-- operating_contract.md
|   +-- planning_plugins.md
|   +-- power_platform_fabric_integration.md
|   +-- public_release_checklist.md
|   +-- reference_architecture_tree.md
|   +-- scenario_matrix.md
|   +-- starter_workbook.md
|   +-- technical_review_guide.md
|   +-- office_addin.md
|   +-- workbook_import_map.md
|   \-- change_log.md
+-- office-scripts/
|   +-- README.md
|   +-- apply_notes.ts
|   \-- apply_asset_mappings.ts
+-- samples/
|   +-- planning_table_starter.tsv
|   +-- cap_setup_starter.tsv
|   +-- asset_register_starter.tsv
|   +-- power-query/
|   \-- workflow starter TSVs
+-- package.json
\-- tools/
    +-- audit_capex_module.py
    +-- build_governance_starter_workbook.ps1
    +-- build_asset_evidence_pq_seed.ps1
    +-- install_asset_evidence_pq_workbook.ps1
    +-- start_asset_evidence_pq_installer.ps1
    +-- lint_formulas.py
    +-- start_addin_smoke_test.ps1
    +-- start_addin_dev_server.ps1
    +-- stop_addin_smoke_test.ps1
    \-- push_public.ps1
```

## Quick Checks

From the repo root:

```bash
python tools/audit_capex_module.py
python tools/lint_formulas.py modules/*.formula.txt
python tools/report_feature_status.py
```

The audit is intentionally text-only. It does not open Excel, edit workbook binaries, or require workbook data.

For release review, `npm run validate` runs the static audit, formula lint, and feature-status reporter. `npm run review:packet` writes an ignored review packet under `release_artifacts/review_packet/` so reviewers can see what is built, scaffolded, or still missing.

## Starter Workbook Editions

The default generated `Governance_Starter.xltx` is planning-only. Its visible flow is:

```text
Start Here -> Source Status -> Data Import Setup -> Planning Table -> Cap Setup -> Planning Review -> Analysis Hub
```

Asset workflow is optional. Start with Asset Hub only when you explicitly need project-to-asset tracking. `AssetsLite` adds `Asset Hub`; `AssetsFull` adds both `Asset Hub` and `Asset Finance Hub`. `SemanticTwin` adds `Semantic Map Hub` for optional REC/Brick semantic crosswalk review.

`tblBudgetInput` is the canonical formula source. `Planning Table` / `tblPlanningTable` is manual/staging/local writeback. After manual Planning Table edits or `ApplyNotes`, refresh or re-sync the current-workbook adapter before relying on formula outputs.

To validate, commit, rebase, and push the public repo in one local command:

```powershell
.\tools\push_public.ps1 -Message "Update public formula template"
```

## Worktree Workflow

Use `main` as the stable product branch. Use short-lived worktrees for concurrent Excel-work roles: `main` for pristine public template state, `work` for formula/add-in tasks, `review` for PR and workbook-contract review, `fuzz` for automated smoke/lint runs, and `scratch` for disposable workbook-reference analysis.

```powershell
.\tools\new_worktree.ps1 -Name ready-fix
```

See `docs/git_worktree_workflow.md` for the starter workflow.

## Why Excel

Excel is the right runtime for this pattern when the real work is planning, review, and decision support by people who already live in workbooks.

The goal is not to pretend that a workbook is a database or a web application. The goal is to make workbook logic governable:

- planners can inspect formulas, source rows, pivots, and exceptions without leaving Excel,
- dynamic arrays can produce live review screens without a separate service layer,
- `LET` and `LAMBDA` allow complex logic to be named, split, and reused,
- Git can track formula modules as source code while excluding workbook binaries,
- text audit can enforce size limits, public-safety rules, and key formula contracts.

A different stack is better when the primary need is multi-user transactions, permissions, APIs, durable storage, or application workflows. This repo is for the space where Excel is already the operating surface, and the missing discipline is source control, reviewability, and repeatable validation.

## Current Readiness

This repo works as a public source-code template:

- the text audit and formula lint pass from a clean checkout,
- the formula modules are importable plain-text examples,
- the Office.js starter add-in can install the modules into workbook names,
- the starter table gives a blank workbook the expected source-table shape,
- public-safety checks block private labels, paths, workbook binaries, and old sample codes.

This repo is not a turnkey workbook:

- it does not track workbook binaries in Git,
- it can generate a local governance starter `.xltx` / `.xlsx` under ignored `release_artifacts/`,
- the Office.js add-in is a starter installer, not a production Marketplace package,
- it does not prove runtime recalculation inside every Excel tenant,
- a real workbook owner still needs to map their own table names, headers, caps, and review process.

## Generated Governance Starter Template

For a local workbook artifact, build the generated starter:

```powershell
npm run build:governance-starter
```

or:

```powershell
.\tools\build_governance_starter_workbook.ps1
```

The script creates ignored artifacts under `release_artifacts/governance-starter/`:

```text
Governance_Starter.xlsx
Governance_Starter.xltx
```

The `.xltx` is the user-facing Excel template. The `.xlsx` is kept beside it for inspection and smoke testing. Both are generated from tracked text sources: formula modules in `modules/`, starter TSVs in `samples/`, and M templates in `samples/power-query/`.

The generated workbook opens on `Start Here` and keeps the default visible surface small: `Source Status`, `Data Import Setup`, `Planning Table`, `Cap Setup`, `Planning Review`, and `Analysis Hub`. Asset and semantic surfaces are opt-in by edition. `Start Here` includes workbook flow, the `tblBudgetInput` source rule, navigation links, and the hidden-backend explanation. Governed backend sheets such as `PQ Budget Input`, `PQ Budget QA`, `Validation Lists`, `Decision Staging`, asset workflow sheets, semantic setup sheets, and intermediate asset-evidence Power Query outputs are still generated, but hidden by default from the source-controlled workbook manifest. The manifest includes a `Presence` field so legacy sheet names can remain documented as `OptionalLegacy` without being created as primary workbook sheets.

The v0.5 data import bridge adds `Data Import Setup`, hidden `PQ Budget Input`, and hidden `PQ Budget QA` sheets. The generated template creates `tblDataSourceProfile`, `tblBudgetImportParameters`, `tblBudgetImportContract`, `tblBudgetInput`, `tblBudgetImportStatus`, and `tblBudgetImportIssues`. Formula modules now read the canonical `tblBudgetInput` table; `Planning Table` remains the manual starter source and current-workbook adapter source.

For v0.5, `Planning Table` is a manual/staging surface and governed report formulas consume `tblBudgetInput`. If `Planning Table` changes, refresh or re-sync the current-workbook adapter before relying on formula outputs.

The generated starter also includes a hidden `Automation Setup` worksheet. It explains how to import the optional `ApplyNotes.ts` release asset through Excel `Automate -> New Script`; the public template does not embed or auto-install Office Scripts.

The generated starter includes the first v0.4 asset finance bridge: hidden `Asset Finance Setup` / `tblAssetFinanceAssumptions`, the `AssetFinance` formula names, and `Asset Finance Hub` sections for depreciation, funding requirements, totals, and chart-ready feeds. The stacked hub sheets include clickable `Go to section` tables near the top so operators can jump down-page to the relevant output without browsing backend sheets. Those outputs read `tblAssetEvidence_ModelInputs`; mapped-only evidence remains reviewable in the Power Query status and mapping queue, but only rows with `PresentWithClassifiedEvidence = TRUE` feed the finance model outputs.

## Start From A Blank Workbook

For a first local trial without the generated template, create a blank workbook and follow `docs/starter_workbook.md`.

The paste-ready starter table is in:

```text
samples/planning_table_starter.tsv
samples/cap_setup_starter.tsv
samples/budget_import_contract_starter.tsv
```

Paste the planning table into `Planning Table!A2` and into `PQ Budget Input!A1` as `tblBudgetInput`, paste the cap table into `Cap Setup!A2`, set `Planning Review!M2` to a month abbreviation such as `Mar`, then import the formula modules. For a normal v0.5 start, use the generated `Governance_Starter.xltx` so the canonical import tables already exist.

See `docs/database_import_contract.md` for the canonical `tblBudgetInput` contract, `docs/power_platform_fabric_integration.md` for the later platform path, and `docs/copilot_review_playbook.md` for prompt-card guidance. Copilot may explain and summarize reviewed tables, but governed numeric calculations stay in native Excel formulas.

## Office.js Add-In Starter

The `addin/` folder contains a minimal Excel task-pane add-in. It creates starter sheets, installs workbook defined names from the text formula modules, and validates the workbook contract.

Use it as a packaging starter, not as a replacement calculation engine. The add-in installs native Excel formulas; the planning logic still lives in workbook named formulas after installation.

For operator-style local use after downloading the repo ZIP, start with:

```powershell
.\Start-AddIn.ps1
```

That launcher confirms you are using a workbook copy, installs npm dependencies when `node_modules` is missing, then starts the local add-in and launches Excel. It does not edit a workbook by itself; workbook changes happen only after you open a workbook copy and click setup or apply buttons.

The same operator launcher is available through npm:

```powershell
npm run start:addin
npm run excel:addin
```

For the shortest operator checklist, see `README_FIRST.md`.

To run the local smoke test on Windows:

```powershell
.\tools\start_addin_smoke_test.ps1
```

That helper runs the static checks, starts a local HTTPS server, and asks Excel desktop to sideload `addin/manifest.xml` when npm is available. If npm is not available, it still starts the local server and prints the manual sideload fallback.

With Node/npm installed, the same smoke path is also available as:

```powershell
npm run addin:smoke
npm run test:smoke
```

When the test session is done:

```powershell
.\tools\stop_addin_smoke_test.ps1
```

See `docs/office_addin.md`.

## Notes And Asset Workflows

The v0.2.0 workflow layer adds a controlled notes/status/timeline apply path and an optional asset setup path.

- `Setup Notes Workflow` creates `Planning Review!O:R` notes columns and `Decision Staging` / `tblDecisionStaging`.
- `office-scripts/apply_notes.ts` performs the two-pass prepare/apply writeback to `Planning Notes`, `Timeline`, `Comments`, and `Status`.
- `Setup Asset Workflow` is optional and creates asset review/apply tables for controlled workbook writes.
- Asset Evidence Power Query is a separate seed-workbook path. Source-controlled M templates live in `samples/power-query/asset-evidence/`; `tools/start_asset_evidence_pq_installer.ps1` provides the local button launcher, and `tools/install_asset_evidence_pq_workbook.ps1` installs the setup sheets, query definitions, and loaded output tables into a new target workbook copy.
- `modules/assets.formula.txt` contains review queues only; Office Scripts perform controlled writes.

See `docs/notes_apply_workflow.md`, `docs/asset_setup_workflow.md`, `docs/asset_evidence_power_query.md`, and `office-scripts/README.md`.

## Optional SemanticTwin Crosswalk

`SemanticTwin` is an opt-in starter edition for digital-twin-readiness review. It adds `Semantic Map Hub` and hidden `Semantic Map Setup` tables without changing the default planning workbook.

The semantic layer is optional. Use REC for buildings, rooms, spaces, real-estate context, and generic assets. Use Brick for equipment, points, sensors, meters, setpoints, commands, and building systems. This is not a full ontology import, not a JSON-LD/RDF exporter, and not a completed Azure Digital Twins or Fabric graph integration.

The first slice is a curated crosswalk:

- `tblOntologyNamespaces`
- `tblOntologyClassMap`
- `tblOntologyRelationshipMap`
- `tblProjectSemanticMap`
- `tblAssetSemanticMap`
- `Ontology.TRIPLE_EXPORT_QUEUE`
- `Ontology.ONTOLOGY_ISSUES`

See `docs/semantic_standards_strategy.md`.

## Core Pattern

The example workbook logic is split into modules:

- `get` owns workbook range extraction helpers.
- `kind` owns shared calculation, cap lookup, grouping, flag, and display helpers.
- `CapitalPlanning` owns the main `CAPITAL_PLANNING_REPORT()` formula.
- `Analysis` owns optional planning screens such as `BU_CAP_SCORECARD()` and `REFORECAST_QUEUE([groupBy])`.

Cap limits are workbook inputs. Create a `Cap Setup` sheet, paste `samples/cap_setup_starter.tsv` into `Cap Setup!A2`, and replace the fake caps with your own planning limits. `kind.CapByBU(...)` maps the BU code before any colon in `Planning Table[BU]` to that cap table, and `kind.PortfolioCap` sums the table.

The important implementation idea is the boundary, not the sample vocabulary: keep complex Excel logic in importable text modules, keep workbook binaries out of Git, and make formula behavior reviewable with docs and static checks.
