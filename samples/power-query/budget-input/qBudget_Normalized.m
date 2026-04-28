let
    Source = qBudget_Source_CurrentWorkbook,
    Contract = Excel.CurrentWorkbook(){[Name = "tblBudgetImportContract"]}[Content],
    ContractColumns = Table.Column(Contract, "ColumnName"),
    Selected = Table.SelectColumns(Source, ContractColumns, MissingField.UseNull),
    Typed = Table.TransformColumnTypes(
        Selected,
        List.Transform(ContractColumns, each {_, type any})
    )
in
    Typed
