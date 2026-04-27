let
    Source = Excel.CurrentWorkbook(){[Name = "tblAssetEvidenceSource"]}[Content],
    ExpectedColumns = {
        "EvidenceId",
        "SourceSystem",
        "SourceRecordType",
        "SourceRecordKey",
        "ProjectKey",
        "AssetId",
        "AssetLabel",
        "AssetType",
        "EvidenceDate",
        "Amount",
        "FundingSource",
        "DepreciationClass",
        "ContextCategoryId",
        "ContextCategoryName",
        "Description"
    },
    WithExpectedColumns =
        List.Accumulate(
            ExpectedColumns,
            Source,
            (state, columnName) =>
                if Table.HasColumns(state, columnName)
                then state
                else Table.AddColumn(state, columnName, each null)
        ),
    Selected = Table.SelectColumns(WithExpectedColumns, ExpectedColumns),
    TextColumns = List.RemoveItems(ExpectedColumns, {"EvidenceDate", "Amount"}),
    Trimmed =
        Table.TransformColumns(
            Selected,
            List.Transform(
                TextColumns,
                (columnName) => {columnName, each if _ = null then null else Text.Trim(Text.From(_)), type nullable text}
            )
        ),
    Typed =
        Table.TransformColumns(
            Trimmed,
            {
                {"EvidenceDate", each try Date.From(_) otherwise null, type nullable date},
                {"Amount", each try Number.From(_) otherwise null, type nullable number}
            }
        )
in
    Typed
