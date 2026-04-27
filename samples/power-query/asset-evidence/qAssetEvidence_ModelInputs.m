let
    Source = qAssetEvidence_Status,
    Selected =
        Table.SelectColumns(
            Source,
            {
                "EvidenceId",
                "ProjectKey",
                "AssetId",
                "AssetLabel",
                "AssetType",
                "EvidenceDate",
                "Amount",
                "FundingSource",
                "DepreciationClass",
                "ClassifiedCategoryId",
                "ClassifiedCategoryName",
                "PresentWithSourceEvidence",
                "PresentWithMappedEvidence",
                "PresentWithClassifiedEvidence",
                "ReviewIssue"
            },
            MissingField.UseNull
        )
in
    Selected
