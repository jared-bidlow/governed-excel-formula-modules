let
    // Optional placeholder adapter. Not part of the current operator workflow.
    ServerName = "SERVER_OR_ENDPOINT_PLACEHOLDER",
    DatabaseName = "DATABASE_OR_WORKSPACE_PLACEHOLDER",
    ViewName = "vBudgetPlanningWorkbookContract",
    Source = Sql.Database(ServerName, DatabaseName),
    Data = Source{[Schema = "dbo", Item = ViewName]}[Data]
in
    Data
