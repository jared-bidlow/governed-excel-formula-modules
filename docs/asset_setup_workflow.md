# Asset Setup Workflow

The asset workflow is optional. It is not part of the default `Setup + Install + Validate + Outputs` path.

Start with `Asset Hub` when a workbook needs asset-mapping review surfaces and controlled asset writeback tables. If the workbook only needs capital planning, stay in `Planning Review` and `Analysis Hub`.

Do not start with PQ asset evidence sheets. Do not start with `Asset State History`. Asset Finance is advanced and requires classified evidence.

Rerunning this setup recreates the asset workflow tables from their headers. Use it as a starter/reset action on a copy or before live asset data is entered, not as a migration over populated production tables.

For a new workbook, `tools/build_governance_starter_workbook.ps1 -Edition AssetsLite` can generate an asset-enabled starter template with `Asset Hub` visible. Use `-Edition AssetsFull` only when asset evidence finance outputs are in scope. The add-in button remains useful for blank-workbook setup and controlled resets.

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

Starter TSV files under `samples/` provide public-safe headers and blank starter rows for the staging, mapping, change, and history tables. Demo rows live under `samples/demo/asset_workflow/` so the public starter does not look pre-populated with live assets.

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

- `Assets.ASSET_START_HERE`
- `Assets.ASSET_WORKFLOW_STATUS`
- `Assets.ASSET_NEXT_ACTIONS`
- `Assets.ASSET_TABLE_MAP`
- `Assets.ASSET_GLOSSARY`
- `Assets.ASSET_REVIEW_QUEUE`
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

## Asset Evidence Power Query

Asset Evidence Power Query is a separate optional seed-workbook path for asset evidence imports. It is not part of the normal `Setup + Install + Validate + Outputs` path and is not exposed as task-pane copy buttons.

`Setup Asset Workflow` does not create `tblAssetEvidenceSource`, `tblAssetEvidenceRules`, `tblAssetEvidenceOverrides`, or any `PQ Asset Evidence ...` output sheets. It owns the asset register, mapping, change, and state-history workflow tables only.

Build `release_artifacts/asset-evidence-pq/Asset_Evidence_PQ_Seed.xlsx`, then use `tools/install_asset_evidence_pq_workbook.ps1` to install the seed-owned sheets into a new target workbook copy. The installed setup sheet carries `tblAssetEvidenceSource`, `tblAssetEvidenceRules`, and `tblAssetEvidenceOverrides`; the installed output sheets carry loaded tables for `qAssetEvidence_Normalized`, `qAssetEvidence_Classified`, `qAssetEvidence_Linked`, `qAssetEvidence_Status`, `qAssetEvidence_ModelInputs`, and `qQA_AssetEvidence_MappingQueue` from source-controlled M templates.

The evidence bridge keeps mapped context and classified evidence separate. Asset, project, or context fields can create mapped evidence, but `PresentWithClassifiedEvidence` requires a classified category plus classifier metadata from a rule or override.

## Boundary

This release keeps asset workflow separate from semantic mapping. Asset formulas review workbook state; Office Scripts perform controlled workbook writes; the asset evidence Power Query seed provides setup tables and public-safe M templates only.

This asset setup slice does not include RDF export, SHACL validation, or finished asset reports. The Power Query seed provides setup tables and evidence M templates only.

`SemanticTwin` is optional and sits after the asset workflow. Use REC for buildings, spaces, rooms, real-estate context, and generic assets. Use Brick for equipment, points, sensors, meters, setpoints, commands, and building systems. The semantic crosswalk is not a full ontology import, not SHACL validation, and not a completed graph or Azure Digital Twins integration.
