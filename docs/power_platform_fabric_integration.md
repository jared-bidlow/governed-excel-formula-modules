# Power Platform And Fabric Integration Path

The v0.5 bridge keeps the repo workbook-centered and database-ready without making Power Platform or Fabric required dependencies.

## Recommended Shape

```text
Dataverse, Fabric, SQL, SharePoint export, or manual workbook source
  -> Power Query adapter
  -> qBudget_WideContract
  -> qBudget_Input
  -> tblBudgetInput
  -> governed Excel formula modules
```

## Fabric Path

Fabric should curate source data into a stable planning-contract view before Excel imports it.

```text
Source systems
  -> Dataflow or warehouse shaping
  -> vBudgetPlanningWorkbookContract
  -> qBudget_Source_FabricSqlEndpoint
  -> tblBudgetInput
```

For v0.5, the repo ships only the placeholder adapter and documentation. It does not create a workspace, lakehouse, warehouse, semantic model, shortcut, gateway, or deployment pipeline.

## Power Platform Path

Power Platform is the later workflow path for forms, approvals, and controlled writeback. A future implementation can model planning rows, decisions, asset evidence, and refresh logs in Dataverse, then expose a curated workbook-contract view for Excel Power Query.

For v0.5, the repo ships only the import contract and review guidance. It does not create Dataverse tables, Power Apps, Power Automate flows, environment-specific ALM assets, or direct database writeback.

## Public-Safe Boundary

Tracked files must not contain:

- real tenant names,
- real server names,
- real workspace names,
- real database names,
- credentials,
- tokens,
- private URLs,
- local workbook paths,
- generated workbook binaries.

Use the placeholder adapter values in `samples/power-query/budget-input/` and replace them only inside a private workbook copy or private deployment process.
