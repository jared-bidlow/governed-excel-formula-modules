# Asset Setup Workflow

The asset workflow is optional. It is not part of the default `Setup + Install + Validate + Outputs` path.

Use `Setup Asset Workflow` when a workbook needs asset-mapping review surfaces and controlled asset writeback tables.

Rerunning this setup recreates the asset workflow tables from their headers. Use it as a starter/reset action on a copy or before live asset data is entered, not as a migration over populated production tables.

## Created Sheets

The add-in creates these sheets:

- `Asset Register`
- `Asset Setup`
- `Project Asset Map`
- `Semantic Assets`
- `Asset Changes`
- `Asset State History`

## Created Tables

The add-in creates formatted starter tables aligned to `office-scripts/apply_asset_mappings.ts`:

- `tblAssets`
- `tblSemanticAssets`
- `tblAssetPromotionQueue`
- `tblAssetMappingStaging`
- `tblProjectAssetMap`
- `tblAssetChanges`
- `tblAssetStateHistory`

Starter TSV files under `samples/` provide public-safe example headers and fake starter rows for the staging, mapping, change, and history tables. `tblAssets` is created directly from the add-in's asset-register header contract.

## Asset Table Map

`tblAssets` is the starter asset register. It is the place for durable asset records such as `AssetID`, `AssetName`, `AssetType`, `Site`, `Location`, `Department`, `Owner`, `Status`, `Condition`, `Criticality`, replacement cost, review dates, and optional `LinkedProjectID`.

`tblSemanticAssets` is a formula-facing proposal surface. It holds candidate asset IDs and inferred project-to-asset changes for review before promotion.

`tblAssetPromotionQueue` is the operator queue for accepted candidate assets. Rows marked ready can be consumed by the controlled mapping script.

`tblAssetMappingStaging` is the staging table for specific project-to-asset changes. It carries change type, source asset, target asset, installed state, evidence, status, and apply fields.

`tblProjectAssetMap` is the current relationship table between projects and assets.

`tblAssetChanges` is the append/update log for applied asset mapping changes.

`tblAssetStateHistory` is the asset state event trail written by the controlled apply script when that target exists.

## Dropdowns And Relationships

The setup writes static dropdown values to `Validation Lists` for asset status, condition, criticality, change type, asset state, promotion status, mapping status, change status, and `Y/N` apply readiness.

The setup also adds relationship dropdown sources on `Validation Lists`:

- `Asset ID` spills the distinct asset IDs found across `tblAssets` and workflow tables.
- `Project Key` spills the distinct project keys found across `tblAssets` and workflow tables.

Those relationship lists are applied to asset ID and project key columns in the created tables. They are advisory dropdowns: they help keep project-to-asset links coherent, but they allow new IDs while the register and mapping tables are still being built.

## Review Formulas

`modules/assets.formula.txt` contains dynamic-array review formulas:

- `Assets.PROJECT_PROMOTION_QUEUE`
- `Assets.ASSET_MAPPING_ISSUES`
- `Assets.ASSET_CHANGE_ISSUES`
- `Assets.INSTALLED_WITHOUT_EVIDENCE`
- `Assets.REPLACEMENT_SOURCE_TARGET_ISSUES`

These formulas create review queues only. They do not write rows, mutate tables, export data, or call external services.

## Controlled Writes

`office-scripts/apply_asset_mappings.ts` is the controlled-write action. It reads accepted or ready rows from the asset staging tables, validates basic asset rules, and appends or updates controlled workbook tables when possible.

The script updates project-to-asset mappings, change rows, and state-history rows. It does not create or overwrite the durable `tblAssets` asset register.

The script can write apply status fields when the target columns exist:

- `ApplyStatus`
- `AppliedOn`
- `ApplyMessage`

The script is defensive when expected tables are missing and reports the issue instead of silently applying partial changes.

## Boundary

This release does not include RDF export, SHACL validation, ontology files, reports, or a Power Query bridge. Asset formulas review workbook state; Office Scripts perform controlled workbook writes.
