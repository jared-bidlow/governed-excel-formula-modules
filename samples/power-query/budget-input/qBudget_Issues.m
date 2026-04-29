let
    Source = qBudget_Source_Selected,
    Contract = Excel.CurrentWorkbook(){[Name = "tblBudgetImportContract"]}[Content],
    ExpectedColumns = Table.Column(Contract, "ColumnName"),
    ActualColumns = Table.ColumnNames(Source),
    MissingColumns = List.Difference(ExpectedColumns, ActualColumns),
    ExtraColumns = List.Difference(ActualColumns, ExpectedColumns),
    MissingRows = List.Transform(
        MissingColumns,
        each {"MissingColumn", "Error", _, "", "Required canonical budget input column is missing.", "Open"}
    ),
    ExtraRows = List.Transform(
        ExtraColumns,
        each {"ExtraColumn", "Review", _, "", "Source column is not in the canonical budget input contract.", "Open"}
    ),
    Rows = MissingRows & ExtraRows,
    Issues = if List.Count(Rows) = 0 then
        #table(
            {"IssueType", "Severity", "ColumnName", "SourceKey", "Message", "ReviewStatus"},
            {{"SchemaOK", "Info", "tblBudgetInput", "", "Source columns match the canonical budget input contract.", "Closed"}}
        )
    else
        #table(
            {"IssueType", "Severity", "ColumnName", "SourceKey", "Message", "ReviewStatus"},
            Rows
        )
in
    Issues
