# Optional Platform Adapters

This repo is an Excel-first formula and workbook-template repo. The current operator workflow is:

```text
Planning Table or workbook input
  -> Power Query current-workbook adapter
  -> tblBudgetInput
  -> governed Excel formula modules
  -> review screens and CSV handoff surfaces
```

The normal path is the generated Governance Starter workbook, `tblBudgetInput`, the Integration Bridge, source-controlled formulas, and static validation.

## Placeholder Adapters

The files below are optional placeholder adapters:

- `samples/power-query/budget-input/qBudget_Source_AzureSql.m`
- `samples/power-query/budget-input/qBudget_Source_Dataverse.m`
- `samples/power-query/budget-input/qBudget_Source_FabricSqlEndpoint.m`
- `samples/power-query/budget-input/qBudget_Source_Selected.m`

They are kept so a private workbook owner can extend the same 64-column `tblBudgetInput` contract later without changing formula modules. They are not part of the current operator package, and they are not a recommended next step for this public template.

## Boundaries

The placeholder adapters do not create environments, workspaces, apps, flows, tables, semantic models, deployment pipelines, or direct database writeback.

Tracked files must not contain real tenant names, server names, workspace names, database names, credentials, tokens, private URLs, local workbook paths, or generated workbook binaries.
