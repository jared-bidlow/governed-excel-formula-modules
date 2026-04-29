# Reference Architecture Tree

This document is the durable orientation map for the public formula-module reference architecture.

It explains how the repository tree works, which layer owns which behavior, and where the boundaries are. It is public-safe by design: no workbook binary, private workbook path, company data, or production source data is required to understand the architecture.

## Current Launch State

The v0.5 branch adds a canonical data-import bridge and generated workbook editions.

The launched shape is:

- capital-planning formula modules remain the primary reference architecture;
- the notes workflow is part of the normal add-in setup path;
- the asset workflow is opt-in and visible only in asset-enabled generated editions;
- the default generated starter is planning-only;
- formula modules create review surfaces;
- Office Scripts perform controlled workbook writes;
- workbook binaries and generated workbook artifacts remain out of scope.

## Top-Level Tree

```text
governed-excel-formula-modules/
+-- README.md
+-- AGENTS.md
+-- modules/
|   +-- controls.formula.txt
|   +-- get.formula.txt
|   +-- kind.formula.txt
|   +-- capital_planning_report.formula.txt
|   +-- analysis.formula.txt
|   +-- assets.formula.txt
|   +-- notes.formula.txt
|   +-- ready.formula.txt
|   +-- search.formula.txt
|   +-- defer.formula.txt
|   \-- phasing.formula.txt
+-- addin/
|   +-- manifest.xml
|   +-- taskpane.html
|   +-- taskpane.css
|   +-- taskpane.js
|   \-- assets/
+-- office-scripts/
|   +-- README.md
|   +-- apply_notes.ts
|   \-- apply_asset_mappings.ts
+-- samples/
|   +-- planning_table_starter.tsv
|   +-- cap_setup_starter.tsv
|   +-- decision_staging_starter.tsv
|   +-- asset_setup_starter.tsv
|   +-- semantic_assets_starter.tsv
|   +-- project_asset_map_starter.tsv
|   +-- asset_changes_starter.tsv
|   \-- asset_state_history_starter.tsv
+-- docs/
+-- tools/
\-- package.json
```

## Architecture Layers

| Layer | Primary files | Owns | Does not own |
|---|---|---|---|
| Contract and docs | `README.md`, `docs/*.md`, `AGENTS.md` | Public story, operating rules, workflow boundaries, scenario expectations | Workbook state or hidden implementation |
| Formula modules | `modules/*.formula.txt` | Native Excel named formulas, review surfaces, report logic, helper logic | Hidden writes, workbook mutation, generated artifacts |
| Starter data | `samples/*.tsv` | Public-safe table headers, blank asset starter rows, and demo rows under `samples/demo/` | Production data |
| Office.js add-in | `addin/taskpane.*`, `addin/manifest.xml` | Workbook setup, formula installation, validation, demo output placement | Business calculation engine |
| Office Scripts | `office-scripts/*.ts` | Explicit controlled writeback from staged workbook rows | Silent background mutation or source-control truth |
| Validation tools | `tools/audit_capex_module.py`, `tools/lint_formulas.py` | Public-safety checks, formula presence, formula size, docs and setup contracts | Runtime proof inside every Excel tenant |
| Git workflow | `docs/git_worktree_workflow.md`, `tools/new_worktree.ps1`, `tools/push_public.ps1` | Reviewable source changes, task concurrency, validated push path | Workbook copies or local operator state |

## Runtime Flow

```text
plain-text repo artifacts
        |
        v
Office.js add-in setup
        |
        +--> starter sheets and tables
        +--> validation lists and visible controls
        +--> workbook defined names from modules/*.formula.txt
        +--> demo output formulas
        +--> generated starter editions: Planning, AssetsLite, AssetsFull
        |
        v
Excel workbook runtime
        |
        +--> formula reports and review queues
        +--> operator review and staging
        |
        v
Office Scripts controlled apply
        |
        +--> staged notes/status/timeline writes
        +--> accepted asset mapping/change/history writes
```

The workbook is the runtime surface. Git is the source-control surface. The add-in installs and validates. Office Scripts apply staged changes. Those are separate roles.

Asset workflow is optional. Start with `Asset Hub` only when project-to-asset tracking is needed. Do not start with PQ asset evidence sheets. Do not start with `Asset State History`. Asset Finance is advanced and requires classified evidence.

## Formula Dependency Tree

The core formula dependency spine is:

```text
controls

get
  -> reads tblBudgetInput as the canonical formula source
  -> Planning Table remains manual/staging/local writeback

kind
  -> shared calculation helpers
  -> cap lookup
  -> accounting universe masks
  -> grouping, flags, future filters, closed-row filters

CapitalPlanning
  -> main CAPITAL_PLANNING_REPORT()
  -> uses get + kind

Analysis
  -> PM_SPEND_REPORT()
  -> WORKING_BUDGET_SCREEN()
  -> BU_CAP_SCORECARD()
  -> REFORECAST_QUEUE()
  -> BURNDOWN_SCREEN()
  -> uses get + kind
```

Supporting formula branches are:

```text
Assets
  -> PROJECT_PROMOTION_QUEUE
  -> ASSET_MAPPING_ISSUES
  -> ASSET_CHANGE_ISSUES
  -> INSTALLED_WITHOUT_EVIDENCE
  -> REPLACEMENT_SOURCE_TARGET_ISSUES
  -> review queues only

Notes
  -> meeting-note context and staging helpers

Ready
  -> internal-readiness export helpers

Search
  -> budget search and row health helpers

defer
  -> deferral and policy selector helpers

Phasing
  -> annual-to-month phasing helpers
```

Formula modules are expected to be importable as workbook named formulas. If a formula becomes too large, the preferred fix is to split helper formulas rather than widen one workbook name past Excel's practical save limits.

## Workbook Input Tree

The normal capital-planning workbook shape is:

```text
Planning Table
  -> job/project rows
  -> identity fields
  -> status, BU, PM, region, site, type
  -> annual projected and monthly projected/actual/budget triplets
  -> planning notes, readiness fields, resource fields, cancellation flag

Cap Setup
  -> BU
  -> Cap

Planning Review
  -> visible controls in B2:E2
  -> month controls in M2:N2
  -> main report spill at A4
  -> notes helper/input block at O:R

Validation Lists
  -> status lists
  -> yes/no lists
  -> group and filter lists
  -> optional asset workflow lists
```

The workbook contract is header-aware where possible. The public starter documents the current column layout, but formulas and setup logic should prefer stable header names over hidden positional assumptions unless the contiguous finance block is deliberately being read.

## Default Setup Path

The standard add-in path is:

```text
Setup + Install + Validate + Outputs
  -> create starter workbook sheets
  -> paste starter planning and cap tables
  -> install formula modules as workbook names
  -> bind visible controls
  -> validate required workbook contracts
  -> run Setup Notes Workflow
  -> insert demo output formulas
```

This path stays capital-planning focused. It includes notes setup because the notes workflow is part of the normal operator loop.

## Notes Apply Workflow

The notes workflow is a controlled writeback pattern:

```text
Planning Review notes inputs
        |
        v
Decision Staging / tblDecisionStaging
        |
        v
office-scripts/apply_notes.ts
        |
        +--> Run 1: prepare rows
        \--> Run 2: apply prepared rows
```

The script writes only controlled target fields on `Planning Table`:

- `Planning Notes`
- `Timeline`
- `Comments`
- `Status`

This is intentionally not a hidden write. The workbook user stages edits, reviews status, and then runs a controlled apply action.

## Optional Asset Workflow

The asset workflow is separate:

```text
Setup Asset Workflow
  -> Asset Register / tblAssets
  -> Asset Setup / tblAssetPromotionQueue
  -> Asset Setup / tblAssetMappingStaging
  -> Project Asset Map / tblProjectAssetMap
  -> Semantic Assets / tblSemanticAssets
  -> Asset Changes / tblAssetChanges
  -> Asset State History / tblAssetStateHistory
  -> asset dropdown and relationship lists on Validation Lists
```

Important boundary:

- `Setup Asset Workflow` is opt-in.
- It is not part of the normal setup button.
- Rerunning it recreates workflow tables from their headers.
- Treat it as starter/reset setup, not as a migration over populated production tables.

The asset table ownership model is:

| Table | Role |
|---|---|
| `tblAssets` | Durable asset register starter table. |
| `tblSemanticAssets` | Formula-facing candidate/proposal surface. |
| `tblAssetPromotionQueue` | Operator queue for accepted candidate assets. |
| `tblAssetMappingStaging` | Reviewed project-to-asset change staging. |
| `tblProjectAssetMap` | Current project-to-asset relationships. |
| `tblAssetChanges` | Applied mapping/change log. |
| `tblAssetStateHistory` | Asset state event trail. |

`modules/assets.formula.txt` is review-only. It can identify project promotion candidates and asset mapping issues, but it does not mutate workbook tables.

`office-scripts/apply_asset_mappings.ts` is the controlled-write layer. It can update project-to-asset mappings, change rows, and state-history rows when staging rows are accepted or ready. It does not create, overwrite, or enrich the durable `tblAssets` register.

## Validation Tree

The main validation gates are:

```text
tools/lint_formulas.py
  -> formula balance and text-shape checks

tools/audit_capex_module.py
  -> public-safety strings
  -> forbidden workbook/generated artifacts
  -> required formula names
  -> installed formula size budgets
  -> starter table shape
  -> add-in setup contracts
  -> notes workflow contracts
  -> optional asset workflow contracts
  -> docs and release checklist coverage

git diff --check
  -> whitespace and patch hygiene

tools/start_addin_smoke_test.ps1
  -> static checks
  -> local HTTPS add-in host
  -> Excel sideload when available
```

Passing static checks means the source-controlled architecture is internally consistent. It does not prove every workbook tenant, Excel version, or production data shape recalculates correctly.

## Public/Private Boundary

The public repo may contain:

- generic formula modules;
- generic docs and runbooks;
- public-safe starter TSVs;
- Office.js setup code;
- Office Scripts for controlled writeback;
- static validation tools.

The public repo must not contain:

- workbook binaries;
- production workbook copies;
- real project, employee, vendor, customer, or company data;
- local workbook paths;
- private implementation labels;
- generated workbook packages.

## Reality Checks And Tensions

These distinctions are intentional and should not be collapsed:

| Apparent conflict | Actual architecture rule |
|---|---|
| Excel is the runtime, but Git is the durable source. | Workbook state is operational; text modules, docs, samples, and scripts are source-controlled. |
| The add-in installs formulas, but it is not the calculation engine. | The add-in packages setup and validation; native Excel formulas do the calculation after install. |
| Formula modules can surface review queues, but they cannot apply decisions. | Formula outputs are review surfaces. Office Scripts are controlled write actions. |
| Notes setup is standard, but asset setup is optional. | Notes support the core planning review loop. Asset tracking is an opt-in extension. |
| `tblAssets` is durable inside the workbook, but not durable inside Git. | The table is a workbook data surface. The repo owns its schema starter and controlled scripts, not live asset rows. |
| Semantic mapping can support digital-twin readiness, but it is not the default workbook. | `SemanticTwin` is optional and uses curated REC and Brick crosswalk tables instead of full ontology imports. |
| Relationship dropdowns help coherence, but they are not strict referential integrity. | They are advisory while the register and mapping tables are being built. |
| Worktrees support concurrent work, but branches still own source history. | Worktrees are working folders. Git commits, PRs, and merges are the durable source-control record. |

## Deferred Or Missing By Design

The current architecture does not include:

- workbook binaries;
- production data;
- full RDF, Turtle, JSON-LD, or graph export;
- SHACL validation;
- full ontology publication or ontology dumps;
- a production AppSource package;
- strict project/asset referential integrity;
- a dedicated evidence table for asset evidence;
- a controlled script that creates or updates `tblAssets` as a source-system sync;
- a migration path for populated asset workflow tables.

Those are not accidental omissions. They are deferred boundaries until a later release chooses them explicitly.

## Best Entry Points

For review:

1. `README.md`
2. `docs/technical_review_guide.md`
3. `docs/operating_contract.md`
4. `docs/reference_architecture_tree.md`
5. `tools/audit_capex_module.py`

For workbook setup:

1. `docs/starter_workbook.md`
2. `docs/office_addin.md`
3. `samples/planning_table_starter.tsv`
4. `samples/cap_setup_starter.tsv`

For notes/apply:

1. `docs/notes_apply_workflow.md`
2. `office-scripts/apply_notes.ts`

For asset workflow:

1. `docs/asset_setup_workflow.md`
2. `docs/asset_tracker_next_steps.md`
3. `modules/assets.formula.txt`
4. `office-scripts/apply_asset_mappings.ts`

For optional SemanticTwin crosswalk:

1. `docs/semantic_standards_strategy.md`
2. `modules/ontology.formula.txt`
3. `samples/ontology_class_map_starter.tsv`
4. `samples/ontology_relationship_map_starter.tsv`

SemanticTwin is optional. REC is for buildings, rooms, spaces, real-estate context, and generic assets. Brick is for equipment, points, sensors, meters, setpoints, commands, and building systems. This is not a full ontology import.

## Maintenance Rule

When the architecture changes, update this map in the same pass as the docs, formula modules, add-in setup, Office Scripts, and audit checks that define the new behavior.
