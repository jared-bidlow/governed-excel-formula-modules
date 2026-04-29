# Data Import Bridge Contract

v0.5 adds a canonical import layer in front of the existing formula modules:

```text
External or manual source -> Power Query adapter -> tblBudgetInput -> formula modules
```

Excel remains the review and calculation surface. This slice does not build a database app, Power App, Fabric workspace, or direct database writeback.

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

## Power Query Templates

Budget import templates live under:

```text
samples/power-query/budget-input/
```

The source adapters are:

- `qBudget_Source_CurrentWorkbook`
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

The database-oriented adapters are placeholder templates. They use public-safe placeholder names only. Do not commit real server names, tenant names, workspace names, connection strings, credentials, tokens, private URLs, or local workbook paths.

`qBudget_Source_Selected` reads `tblBudgetImportParameters[ActiveAdapter]` and selects `CurrentWorkbook`, `AzureSQL`, `Dataverse`, or `FabricSqlEndpoint`. `qBudget_Normalized`, `qBudget_Issues`, and `qBudget_Status` use that selected adapter path.

## Operator Flow

1. Open `Governance_Starter.xltx` as a workbook copy.
2. Use `Planning Table` / `tblPlanningTable` for manual starter data, or import a Power Query adapter.
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
- No Power App implementation.
- No Fabric workspace automation.
- No direct database writeback.
- No workbook binaries in tracked source.
