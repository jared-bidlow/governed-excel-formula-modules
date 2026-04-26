# Asset Setup Workflow

The asset workflow is optional. It is not part of the default `Setup + Install + Validate + Outputs` path.

Use `Setup Asset Workflow` when a workbook needs asset-mapping review surfaces and controlled asset writeback tables.

## Created Sheets

The add-in creates these sheets:

- `Asset Setup`
- `Project Asset Map`
- `Semantic Assets`
- `Asset Changes`
- `Asset State History`

## Created Tables

The add-in creates formatted starter tables aligned to `office-scripts/apply_asset_mappings.ts`:

- `tblSemanticAssets`
- `tblAssetPromotionQueue`
- `tblAssetMappingStaging`
- `tblProjectAssetMap`
- `tblAssetChanges`
- `tblAssetStateHistory`

Starter TSV files under `samples/` provide public-safe example headers and fake starter rows for these tables.

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

The script can write apply status fields when the target columns exist:

- `ApplyStatus`
- `AppliedOn`
- `ApplyMessage`

The script is defensive when expected tables are missing and reports the issue instead of silently applying partial changes.

## Boundary

This release does not include RDF export, SHACL validation, ontology files, reports, or a Power Query bridge. Asset formulas review workbook state; Office Scripts perform controlled workbook writes.
