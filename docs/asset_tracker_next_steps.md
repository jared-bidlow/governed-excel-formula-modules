# Asset Tracker Next Steps

This branch is the reference implementation for turning the governed formula starter into an asset-tracker starter. It should stay separate from the active Capital Planning workbook logic unless a later branch deliberately ports selected pieces back.

## Current State

- The default `Setup + Install + Validate + Outputs` path remains capital-planning focused.
- `Setup Notes Workflow` remains part of the default setup path and creates `Decision Staging` / `tblDecisionStaging`.
- `Setup Asset Workflow` is opt-in and creates the asset register, setup, mapping, change, and state-history sheets.
- `tools/build_governance_starter_workbook.ps1` generates `Governance_Starter.xltx` with those asset workflow sheets plus asset-evidence Power Query output sheets already present.
- The v0.4 branch `codex/asset-finance-model-modules` adds `Asset Finance Setup`, `tblAssetFinanceAssumptions`, the `AssetFinance` formula module, and depreciation, funding requirements, totals, and chart-ready output sheets.
- `tblAssets` is the durable asset register starter table.
- `tblProjectAssetMap`, `tblAssetChanges`, and `tblAssetStateHistory` are the controlled-write targets for `office-scripts/apply_asset_mappings.ts`.
- `modules/assets.formula.txt` provides review queues only; formulas do not mutate workbook tables.
- Asset Evidence Power Query is a separate opt-in seed-workbook path that creates source, rule, and override tables plus six loaded query output sheets from source-controlled M templates.

## Table Ownership

| Table | Role |
|---|---|
| `tblAssets` | Durable asset records and asset attributes. |
| `tblSemanticAssets` | Candidate asset proposals surfaced from workbook/project context. |
| `tblAssetPromotionQueue` | Operator queue for accepted candidate assets. |
| `tblAssetMappingStaging` | Reviewed staging rows for project-to-asset mapping changes. |
| `tblProjectAssetMap` | Current project-to-asset relationships. |
| `tblAssetChanges` | Applied mapping/change log. |
| `tblAssetStateHistory` | Applied asset state event trail. |
| `tblAssetEvidence_ModelInputs` | Loaded Power Query bridge table consumed by asset finance formulas. |
| `tblAssetFinanceAssumptions` | Operator assumptions for depreciation life, funding rule, and chart grouping. |

## Immediate Verification

1. Start from a workbook copy.
2. Run the add-in setup path for formulas and notes.
3. Run `Setup Asset Workflow` only when ready to create or reset asset tables.
4. Confirm `Asset Register` / `tblAssets` exists before entering real assets.
5. Confirm project key and asset ID columns show advisory dropdowns from `Validation Lists`.
6. Run `office-scripts/apply_asset_mappings.ts` only after staging rows are marked ready.
7. Build the seed workbook with `tools/build_asset_evidence_pq_seed.ps1` when the M templates change.
8. Run `tools/start_asset_evidence_pq_installer.ps1` for the button-driven local installer, or run `tools/install_asset_evidence_pq_workbook.ps1 -TargetWorkbookPath <workbook-copy.xlsx>` directly when evidence import sheets and loaded query tables are wanted.
9. Open the generated `.asset-evidence-pq.xlsx` output workbook and inspect Power Query load settings before treating it as the working copy.
10. For the generated `Governance_Starter.xltx`, refresh Power Query, inspect `PQ Asset Evidence Model Inputs`, then review `Asset Depreciation`, `Asset Funding Requirements`, `Asset Finance Totals`, and `Asset Finance Charts`.

## Follow-Up Decisions

- Decide whether `tblAssets` should be manually maintained, imported from a source system, or updated by a separate controlled apply script.
- Decide whether the v0.4 chart-ready feed tables should become native Excel chart objects in a later generated-workbook slice.
- Decide whether project keys should come from the planning table, a project master table, or only the asset workflow tables.
- Decide whether advisory relationship dropdowns should later become strict enough to reject unknown asset/project IDs.
