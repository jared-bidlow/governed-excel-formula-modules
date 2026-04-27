let
    Source = qAssetEvidence_Classified,
    HasText = (value as any) as logical =>
        value <> null and Text.Trim(Text.From(value)) <> "",
    WithMappedCategoryId =
        Table.AddColumn(
            Source,
            "MappedCategoryId",
            each if HasText([ContextCategoryId]) then Text.Trim(Text.From([ContextCategoryId])) else null,
            type nullable text
        ),
    WithMappedCategoryName =
        Table.AddColumn(
            WithMappedCategoryId,
            "MappedCategoryName",
            each if HasText([ContextCategoryName]) then Text.Trim(Text.From([ContextCategoryName])) else null,
            type nullable text
        ),
    WithSourceEvidence =
        Table.AddColumn(
            WithMappedCategoryName,
            "HasSourceEvidence",
            each HasText([EvidenceId]) or HasText([SourceRecordKey]) or HasText([Description]),
            type logical
        ),
    WithMappedEvidence =
        Table.AddColumn(
            WithSourceEvidence,
            "HasMappedEvidence",
            each
                HasText([MappedCategoryId])
                    or HasText([MappedCategoryName])
                    or HasText([AssetId])
                    or HasText([ProjectKey]),
            type logical
        )
in
    WithMappedEvidence
