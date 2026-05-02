let
    Source = Excel.CurrentWorkbook(){[Name = "tblBudgetInput"]}[Content],
    RequiredColumns = {"Source ID", "Job ID", "Project Description", "Status", "BU", "Category", "Site", "PM"},
    Shaped = Table.SelectColumns(Source, RequiredColumns, MissingField.UseNull),
    HasText = (value as any) as logical =>
        value <> null and Text.Trim(Text.From(value)) <> "",
    WithProjectKey = Table.AddColumn(
        Shaped,
        "ProjectKey",
        each
            let
                SourceId = if HasText([Source ID]) then Text.Trim(Text.From([Source ID])) else null,
                JobId = if HasText([Job ID]) then Text.Trim(Text.From([Job ID])) else null
            in
                if SourceId <> null and JobId <> null then SourceId & "-" & JobId else null,
        type nullable text
    ),
    Output = Table.ReorderColumns(
        WithProjectKey,
        {"Source ID", "Job ID", "ProjectKey", "Project Description", "Status", "BU", "Category", "Site", "PM"},
        MissingField.UseNull
    )
in
    Output
