let
    // Optional placeholder adapter. Not part of the current operator workflow.
    Environment = "DATAVERSE_ENVIRONMENT_PLACEHOLDER",
    EntityName = "gef_budgetplanningworkbookcontract",
    Source = CommonDataService.Database(Environment),
    Data = Source{[Schema = "dbo", Item = EntityName]}[Data]
in
    Data
