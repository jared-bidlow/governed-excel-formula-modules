# Asset Evidence Power Query

This branch adds an optional asset-evidence Power Query seed workbook for workbooks where the main operating object is an asset, project, funding requirement, or financial model input rather than a folder or file.

The Office.js add-in does not create Power Query queries. The asset-evidence import path is:

- `samples/power-query/asset-evidence/*.m` are the Power Query module sources.
- `tools/build_asset_evidence_pq_seed.ps1` builds `release_artifacts/asset-evidence-pq/Asset_Evidence_PQ_Seed.xlsx`.
- `tools/install_asset_evidence_pq_workbook.ps1` installs the same setup sheets, query definitions, and loaded output tables into a new copy of a target workbook.

The generated seed workbook is a release artifact, not the source of truth. Rebuild it whenever the M templates change.

For a new workbook, prefer the generated governance starter template:

```powershell
.\tools\build_governance_starter_workbook.ps1
```

That build creates `release_artifacts/governance-starter/Governance_Starter.xltx` with the planning starter, optional asset workflow tables, and these asset-evidence queries already loaded. Use the installer below when adding the same query surface to an existing workbook copy.

## Operator Flow

1. Run the normal add-in setup on a workbook copy when formula modules and starter sheets are needed, or start from `Governance_Starter.xltx`.
2. Build the seed workbook with `tools/build_asset_evidence_pq_seed.ps1`.
3. Run `tools/install_asset_evidence_pq_workbook.ps1 -TargetWorkbookPath <workbook-copy.xlsx>`.
4. Open the new `.asset-evidence-pq.xlsx` output workbook.
5. Confirm the installed `Asset Evidence Setup` sheet, six `PQ Asset Evidence ...` output sheets, and workbook queries are present.
6. Refresh Power Query and review load settings before treating the output workbook as the working copy.

The install script does not edit the original target workbook. It creates a new output copy, adds the seed-owned setup and output sheets, creates the query definitions from the source-controlled M files, refreshes the loaded query tables, and saves the output workbook.

If the output workbook already contains the asset-evidence seed sheets or matching query names, pass `-ReplaceExisting` to replace only those seed-owned objects. Pass `-Force` only when replacing the output workbook file itself.

`Setup Asset Workflow` is not a prerequisite for these asset-evidence setup tables. It can be useful when the workbook also needs an asset register, project-to-asset mappings, change logs, or state history, but Power Query should not load into those add-in-created workflow tables.

## Build Command

```powershell
.\tools\build_asset_evidence_pq_seed.ps1
```

or:

```powershell
npm run build:asset-evidence-pq-seed
```

## Install Command

For a button-driven local install on Windows:

```powershell
.\tools\start_asset_evidence_pq_installer.ps1
```

That launcher lets the operator browse for a workbook copy and click `Install Asset Evidence PQ`. It still writes a separate output workbook and does not edit the original target workbook.

The same operation can be run directly:

```powershell
.\tools\install_asset_evidence_pq_workbook.ps1 -TargetWorkbookPath "<path-to-workbook-copy.xlsx>"
```

Optional output path:

```powershell
.\tools\install_asset_evidence_pq_workbook.ps1 -TargetWorkbookPath "<path-to-workbook-copy.xlsx>" -OutputPath "<path-to-output-workbook.xlsx>"
```

## Installed Sheets

| Sheet | Role |
|---|---|
| `Asset Evidence Setup` | Starter source, rule, and override tables. |
| `PQ Asset Evidence Normalized` | Loaded table for `qAssetEvidence_Normalized`. |
| `PQ Asset Evidence Classified` | Loaded table for `qAssetEvidence_Classified`. |
| `PQ Asset Evidence Linked` | Loaded table for `qAssetEvidence_Linked`. |
| `PQ Asset Evidence Status` | Loaded table for `qAssetEvidence_Status`. |
| `PQ Asset Evidence Model Inputs` | Loaded table for `qAssetEvidence_ModelInputs`. |
| `PQ Asset Evidence Mapping Queue` | Loaded table for `qQA_AssetEvidence_MappingQueue`. |

## Setup Tables

| Table | Role |
|---|---|
| `tblAssetEvidenceSource` | Public-safe source evidence rows for assets, projects, funding, depreciation, and model inputs. |
| `tblAssetEvidenceRules` | Rule-driven classifier hints with active/draft/inactive status. |
| `tblAssetEvidenceOverrides` | Reviewed classifier overrides with source metadata and review status. |

## Query Outputs

| Query | Output |
|---|---|
| `qAssetEvidence_Normalized` | Cleans and types `tblAssetEvidenceSource`. |
| `qAssetEvidence_Classified` | Applies overrides first, then active rules, and emits classifier metadata. |
| `qAssetEvidence_Linked` | Adds asset/project/context mapping hints without treating them as true classification. |
| `qAssetEvidence_Status` | Emits `PresentWithSourceEvidence`, `PresentWithMappedEvidence`, and `PresentWithClassifiedEvidence`. |
| `qAssetEvidence_ModelInputs` | Narrow feed for later depreciation, funding requirement, total, and chart modules. |
| `qQA_AssetEvidence_MappingQueue` | Review queue for missing mappings, missing classification, or mapped/classified conflicts. |

## Evidence Rule

`PresentWithMappedEvidence` is allowed to come from structural context: `ContextCategoryId`, `ContextCategoryName`, `AssetId`, or `ProjectKey`.

`PresentWithClassifiedEvidence` requires both a classified category and classifier metadata. The metadata can come from `ClassifierSourceType`, `ClassifierSourceLabel`, or `ClassifierRuleId`.

Structural hints by themselves are not true classified evidence. A mapped asset, project key, or context category can send a row to review, but it cannot satisfy the classified-evidence gate unless a rule or override supplies classifier metadata.

## Source-Control Contract

The workbook artifact supports operator install, but source control remains text-first:

- Review and diff M changes in `samples/power-query/asset-evidence/*.m`.
- Rebuild the seed workbook with `tools/build_asset_evidence_pq_seed.ps1`.
- Install the seed shape into target workbook copies with `tools/install_asset_evidence_pq_workbook.ps1`.
- Keep generated workbook artifacts under `release_artifacts/`.

This slice does not implement depreciation calculations, funding requirement modules, total rollups, or charting outputs. It creates the governed import bridge those modules can consume later.
