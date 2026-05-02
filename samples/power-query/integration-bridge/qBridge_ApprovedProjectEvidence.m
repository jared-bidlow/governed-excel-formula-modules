let
    Source = Excel.CurrentWorkbook(){[Name = "tblApprovedProjectEvidence"]}[Content],
    RequiredColumns = {
        "ProjectKey",
        "EvidenceId",
        "EvidenceType",
        "EvidencePath",
        "EvidenceName",
        "Extension",
        "DocumentAreaID",
        "DocumentAreaName",
        "CategoryID",
        "CategoryName",
        "DateModified",
        "ReviewStatus",
        "ApprovedOn",
        "ReviewerNotes",
        "StatusSignal"
    },
    Shaped = Table.SelectColumns(Source, RequiredColumns, MissingField.UseNull),
    ApprovedOnly = Table.SelectRows(
        Shaped,
        each
            let
                ReviewStatus = try [ReviewStatus] otherwise "",
                ReviewStatusText = if ReviewStatus = null then "" else Text.From(ReviewStatus)
            in
                Text.Upper(Text.Trim(ReviewStatusText)) = "APPROVED"
    )
in
    ApprovedOnly
