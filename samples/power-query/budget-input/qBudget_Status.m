let
    Parameters = Excel.CurrentWorkbook(){[Name = "tblBudgetImportParameters"]}[Content],
    SourceMode = try Parameters{[Parameter = "ActiveAdapter"]}[Value] otherwise "CurrentWorkbook",
    SourceRows = qBudget_Input,
    Status = #table(
        {"QueryName", "SourceMode", "LastRefreshUtc", "RowCount", "Status", "Message"},
        {
            {"qBudget_Input", SourceMode, DateTimeZone.ToText(DateTimeZone.UtcNow()), Table.RowCount(SourceRows), "Ready", "Loaded canonical budget input rows."},
            {"qBudget_Status", SourceMode, DateTimeZone.ToText(DateTimeZone.UtcNow()), 1, "Ready", "Status query refreshed."}
        }
    )
in
    Status
