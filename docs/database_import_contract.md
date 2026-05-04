# Data Import Bridge Contract

This repo keeps the budget source boundary simple:

```text
Planning Table or workbook input -> Power Query adapter -> tblBudgetInput -> formula modules
```

Excel remains the review and calculation surface. The current operator workflow does not build a database app, app workflow, external writeback, or migration path.

Asset workflow is optional and separate from the budget import source boundary. `tblBudgetInput` remains the canonical planning formula source whether or not an asset edition is generated.

## Canonical Tables

The generated `Governance_Starter.xltx` and the Office.js blank-workbook setup create:

| Sheet | Table | Purpose |
|---|---|---|
| `Data Import Setup` | `tblDataSourceProfile` | Public-safe source profile and adapter selection metadata. |
| `Data Import Setup` | `tblBudgetImportParameters` | Operator-facing import parameters. |
| `Data Import Setup` | `tblBudgetImportContract` | The required 64-column wide planning contract. |
| `PQ Budget Input` | `tblBudgetInput` | Canonical loaded budget rows consumed by formulas. |
| `PQ Budget QA` | `tblBudgetImportStatus` | Refresh and source-status output rows. |
| `PQ Budget QA` | `tblBudgetImportIssues` | Schema or source review issues. |

`tblBudgetInput` preserves the existing 64-column planning-table shape for this release. `tblBudgetInput` is the canonical formula source. The `Planning Table` worksheet remains useful as the manual/starter input and local writeback surface, but formulas read `tblBudgetInput[#All]` through `modules/get.formula.txt`.

The generated starter keeps `PQ Budget Input` and `PQ Budget QA` hidden by default. Operators normally review import health through `Source Status` and use `Data Import Setup` plus `Planning Table` for setup and manual/local writeback.

## Integration Bridge Boundary

The optional `Integration Bridge` sheet stages a reviewed-evidence handoff without changing the canonical planning contract.

| Sheet | Table | Purpose |
|---|---|---|
| `Integration Bridge` | `tblFinancialProjectRegisterExport` | Project identity export for a separate evidence review workspace. |
| `Integration Bridge` | `tblApprovedProjectEvidence` | Approved evidence links imported back as advisory context. |

`tblFinancialProjectRegisterExport` uses this public-safe shape:

```text
Source ID | Job ID | ProjectKey | Project Description | Status | BU | Category | Site | PM
```

For the bridge only, `ProjectKey` is derived as:

```text
Source ID & "-" & Job ID
```

`tblApprovedProjectEvidence` accepts approved rows with:

```text
ProjectKey | EvidenceId | EvidenceType | EvidencePath | EvidenceName | Extension | DocumentAreaID | DocumentAreaName | CategoryID | CategoryName | DateModified | ReviewStatus | ApprovedOn | ReviewerNotes | StatusSignal
```

Approved evidence is advisory. It does not auto-create financial projects, does not update official project status from documentation text, and does not treat raw file paths as financial project keys.

## Power Query Templates

Budget import templates live under:

```text
samples/power-query/budget-input/
```

The normal/default source adapter is:

- `qBudget_Source_CurrentWorkbook`

Optional placeholder adapters are also tracked:

- `qBudget_Source_AzureSql`
- `qBudget_Source_Dataverse`
- `qBudget_Source_FabricSqlEndpoint`
- `qBudget_Source_Selected`

The canonical shaping queries are:

- `qBudget_Normalized`
- `qBudget_WideContract`
- `qBudget_Input`
- `qBudget_Status`
- `qBudget_Issues`

The optional integration bridge templates live under:

```text
samples/power-query/integration-bridge/
```

The bridge templates are:

- `qBridge_FinancialProjectRegister`
- `qBridge_ApprovedProjectEvidence`

The non-current-workbook source adapters are placeholder examples. They use public-safe placeholder names only and are not part of the current operator workflow. Do not commit real server names, tenant names, workspace names, connection strings, credentials, tokens, private URLs, or local workbook paths.

`qBudget_Source_Selected` reads `tblBudgetImportParameters[ActiveAdapter]`. `CurrentWorkbook` is the default and expected operator path. `AzureSQL`, `Dataverse`, and `FabricSqlEndpoint` are placeholder adapter names that a private workbook owner can wire later while preserving the same `tblBudgetInput` contract. `qBudget_Normalized`, `qBudget_Issues`, and `qBudget_Status` use the selected adapter output.

## Operator Flow

1. Open `Governance_Starter.xltx` as a workbook copy. Use `Governance_Starter_AssetsLite.xltx` or `Governance_Starter_AssetsFull.xltx` only when the optional asset workflow is in scope.
2. Use `Planning Table` / `tblPlanningTable` for manual starter data, or refresh the current-workbook adapter.
3. Refresh Power Query so `qBudget_Input` loads `tblBudgetInput`.
4. Review `Source Status`; unhide `PQ Budget QA` only when troubleshooting `tblBudgetImportStatus` or `tblBudgetImportIssues`.
5. Review formula outputs that now consume `tblBudgetInput`.

If `ApplyNotes.ts` writes back to `Planning Table`, or if an operator edits `Planning Table` manually, refresh or re-sync the current-workbook budget adapter before relying on formula outputs that read `tblBudgetInput`.

## Source Profile Rules

The public source profile is intentionally descriptive:

- use placeholders for servers, environments, databases, and workspaces,
- use `vBudgetPlanningWorkbookContract` as the example source object,
- store credentials in the operator's Excel/Power Query connection context, not in Git,
- keep all tracked files free of private workbook paths and generated artifacts.

## Excluded From v0.5

- No new Trace, Variance, or Scenario formula modules.
- No expanded AssetFinance calculations.
- No policy-driven defer module.
- No external app implementation.
- No workspace automation.
- No direct database writeback.
- No workbook binaries in tracked source.
