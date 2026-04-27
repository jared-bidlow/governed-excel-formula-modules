let
    Source = qAssetEvidence_Status,
    IssueRows =
        Table.SelectRows(
            Source,
            each [ReviewIssue] <> null and Text.Trim(Text.From([ReviewIssue])) <> ""
        ),
    Selected =
        Table.SelectColumns(
            IssueRows,
            {
                "EvidenceId",
                "ProjectKey",
                "AssetId",
                "AssetLabel",
                "AssetType",
                "MappedCategoryId",
                "MappedCategoryName",
                "ClassifiedCategoryId",
                "ClassifiedCategoryName",
                "ClassifierSourceType",
                "ClassifierSourceLabel",
                "ClassifierRuleId",
                "PresentWithMappedEvidence",
                "PresentWithClassifiedEvidence",
                "ReviewIssue"
            },
            MissingField.UseNull
        )
in
    Selected
