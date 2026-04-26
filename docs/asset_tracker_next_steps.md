# Asset Tracker Next Steps

This branch is the reference implementation for turning the governed formula starter into an asset-tracker starter. It should stay separate from the active Capital Planning workbook logic unless a later branch deliberately ports selected pieces back.

## Current State

- The default `Setup + Install + Validate + Outputs` path remains capital-planning focused.
- `Setup Notes Workflow` remains part of the default setup path and creates `Decision Staging` / `tblDecisionStaging`.
- `Setup Asset Workflow` is opt-in and creates the asset register, setup, mapping, change, and state-history sheets.
- `tblAssets` is the durable asset register starter table.
- `tblProjectAssetMap`, `tblAssetChanges`, and `tblAssetStateHistory` are the controlled-write targets for `office-scripts/apply_asset_mappings.ts`.
- `modules/assets.formula.txt` provides review queues only; formulas do not mutate workbook tables.

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

## Immediate Verification

1. Start from a workbook copy.
2. Run the add-in setup path for formulas and notes.
3. Run `Setup Asset Workflow` only when ready to create or reset asset tables.
4. Confirm `Asset Register` / `tblAssets` exists before entering real assets.
5. Confirm project key and asset ID columns show advisory dropdowns from `Validation Lists`.
6. Run `office-scripts/apply_asset_mappings.ts` only after staging rows are marked ready.

## Follow-Up Decisions

- Decide whether `tblAssets` should be manually maintained, imported from a source system, or updated by a separate controlled apply script.
- Decide whether asset evidence should remain as `EvidenceId` text or get a dedicated evidence table.
- Decide whether project keys should come from the planning table, a project master table, or only the asset workflow tables.
- Decide whether advisory relationship dropdowns should later become strict enough to reject unknown asset/project IDs.
