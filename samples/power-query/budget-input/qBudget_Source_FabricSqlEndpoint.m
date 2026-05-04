let
    // Optional placeholder adapter. Not part of the current operator workflow.
    EndpointName = "FABRIC_SQL_ENDPOINT_PLACEHOLDER",
    DatabaseName = "DATABASE_OR_WORKSPACE_PLACEHOLDER",
    ViewName = "vBudgetPlanningWorkbookContract",
    Source = Sql.Database(EndpointName, DatabaseName),
    Data = Source{[Schema = "dbo", Item = ViewName]}[Data]
in
    Data
