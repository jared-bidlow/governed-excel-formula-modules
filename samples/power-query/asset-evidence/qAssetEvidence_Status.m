let
    Source = qAssetEvidence_Linked,
    HasText = (value as any) as logical =>
        value <> null and Text.Trim(Text.From(value)) <> "",
    WithClassifiedEvidence =
        Table.AddColumn(
            Source,
            "HasClassifiedEvidence",
            each
                (HasText([ClassifiedCategoryId]) or HasText([ClassifiedCategoryName]))
                    and (
                        HasText([ClassifierSourceType])
                            or HasText([ClassifierSourceLabel])
                            or HasText([ClassifierRuleId])
                    ),
            type logical
        ),
    WithSourceStatus =
        Table.AddColumn(
            WithClassifiedEvidence,
            "PresentWithSourceEvidence",
            each [HasSourceEvidence],
            type logical
        ),
    WithMappedStatus =
        Table.AddColumn(
            WithSourceStatus,
            "PresentWithMappedEvidence",
            each [HasMappedEvidence],
            type logical
        ),
    WithClassifiedStatus =
        Table.AddColumn(
            WithMappedStatus,
            "PresentWithClassifiedEvidence",
            each [HasClassifiedEvidence],
            type logical
        ),
    WithReviewIssue =
        Table.AddColumn(
            WithClassifiedStatus,
            "ReviewIssue",
            each
                if not [PresentWithSourceEvidence] then
                    "Missing source evidence"
                else if [PresentWithMappedEvidence] and not [PresentWithClassifiedEvidence] then
                    "Mapped context requires classifier review"
                else if [PresentWithClassifiedEvidence] and not [PresentWithMappedEvidence] then
                    "Classified evidence has no asset or project mapping"
                else
                    "",
            type text
        )
in
    WithReviewIssue
