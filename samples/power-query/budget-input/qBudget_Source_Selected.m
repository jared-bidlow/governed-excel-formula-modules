let
    Parameters = Excel.CurrentWorkbook(){[Name = "tblBudgetImportParameters"]}[Content],
    ActiveAdapter = try Parameters{[Parameter = "ActiveAdapter"]}[Value] otherwise "CurrentWorkbook",
    Source =
        if ActiveAdapter = "AzureSQL" then qBudget_Source_AzureSql
        else if ActiveAdapter = "Dataverse" then qBudget_Source_Dataverse
        else if ActiveAdapter = "FabricSqlEndpoint" then qBudget_Source_FabricSqlEndpoint
        else qBudget_Source_CurrentWorkbook
in
    Source
