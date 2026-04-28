let
    Source = qBudget_Normalized,
    Contract = Excel.CurrentWorkbook(){[Name = "tblBudgetImportContract"]}[Content],
    ContractColumns = Table.Column(Contract, "ColumnName"),
    WideContract = Table.ReorderColumns(Source, ContractColumns, MissingField.UseNull)
in
    WideContract
