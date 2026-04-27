let
    Evidence = qAssetEvidence_Normalized,
    RulesSource = Excel.CurrentWorkbook(){[Name = "tblAssetEvidenceRules"]}[Content],
    OverridesSource = Excel.CurrentWorkbook(){[Name = "tblAssetEvidenceOverrides"]}[Content],
    NormalizeText = (value as any) as text =>
        if value = null then "" else Text.Trim(Text.From(value)),
    EnsureColumns = (table as table, columns as list) as table =>
        List.Accumulate(
            columns,
            table,
            (state, columnName) =>
                if Table.HasColumns(state, columnName)
                then state
                else Table.AddColumn(state, columnName, each null)
        ),
    RuleColumns = {"RuleId", "MatchField", "MatchText", "CategoryId", "CategoryName", "RulePriority", "RuleStatus"},
    OverrideColumns = {
        "EvidenceId",
        "CategoryId",
        "CategoryName",
        "ClassifierSourceType",
        "ClassifierSourceLabel",
        "ClassifierRuleId",
        "OverrideReason",
        "ReviewStatus"
    },
    RulesSelected = Table.SelectColumns(EnsureColumns(RulesSource, RuleColumns), RuleColumns),
    OverridesSelected = Table.SelectColumns(EnsureColumns(OverridesSource, OverrideColumns), OverrideColumns),
    ActiveRules =
        Table.Sort(
            Table.SelectRows(
                RulesSelected,
                each Text.Lower(NormalizeText([RuleStatus])) = "active" and NormalizeText([MatchText]) <> ""
            ),
            {{"RulePriority", Order.Ascending}, {"RuleId", Order.Ascending}}
        ),
    RuleRecords = Table.ToRecords(ActiveRules),
    ClassifyByRule = (row as record) as nullable record =>
        let
            Matches =
                List.Select(
                    RuleRecords,
                    (rule as record) =>
                        let
                            FieldName = NormalizeText(Record.FieldOrDefault(rule, "MatchField", "Description")),
                            FieldValue = Text.Lower(NormalizeText(Record.FieldOrDefault(row, FieldName, ""))),
                            Needle = Text.Lower(NormalizeText(Record.FieldOrDefault(rule, "MatchText", "")))
                        in
                            FieldName <> "" and Needle <> "" and Text.Contains(FieldValue, Needle)
                )
        in
            if List.IsEmpty(Matches) then null else Matches{0},
    MergedOverrides =
        Table.NestedJoin(
            Evidence,
            {"EvidenceId"},
            OverridesSelected,
            {"EvidenceId"},
            "OverrideRows",
            JoinKind.LeftOuter
        ),
    ExpandedOverrides =
        Table.ExpandTableColumn(
            MergedOverrides,
            "OverrideRows",
            {
                "CategoryId",
                "CategoryName",
                "ClassifierSourceType",
                "ClassifierSourceLabel",
                "ClassifierRuleId",
                "OverrideReason",
                "ReviewStatus"
            },
            {
                "OverrideCategoryId",
                "OverrideCategoryName",
                "OverrideClassifierSourceType",
                "OverrideClassifierSourceLabel",
                "OverrideClassifierRuleId",
                "OverrideReason",
                "OverrideReviewStatus"
            }
        ),
    WithRuleRecord = Table.AddColumn(ExpandedOverrides, "RuleRecord", each ClassifyByRule(_), type nullable record),
    WithClassifiedCategoryId =
        Table.AddColumn(
            WithRuleRecord,
            "ClassifiedCategoryId",
            each
                let
                    OverrideCategory = NormalizeText([OverrideCategoryId]),
                    Rule = [RuleRecord]
                in
                    if OverrideCategory <> ""
                    then OverrideCategory
                    else if Rule <> null
                    then NormalizeText(Record.FieldOrDefault(Rule, "CategoryId", ""))
                    else null,
            type nullable text
        ),
    WithClassifiedCategoryName =
        Table.AddColumn(
            WithClassifiedCategoryId,
            "ClassifiedCategoryName",
            each
                let
                    OverrideCategory = NormalizeText([OverrideCategoryName]),
                    Rule = [RuleRecord]
                in
                    if OverrideCategory <> ""
                    then OverrideCategory
                    else if Rule <> null
                    then NormalizeText(Record.FieldOrDefault(Rule, "CategoryName", ""))
                    else null,
            type nullable text
        ),
    WithClassifierSourceType =
        Table.AddColumn(
            WithClassifiedCategoryName,
            "ClassifierSourceType",
            each
                let
                    OverrideCategory = NormalizeText([OverrideCategoryId]),
                    OverrideSource = NormalizeText([OverrideClassifierSourceType]),
                    Rule = [RuleRecord]
                in
                    if OverrideCategory <> ""
                    then if OverrideSource <> "" then OverrideSource else "override"
                    else if Rule <> null
                    then "rule"
                    else null,
            type nullable text
        ),
    WithClassifierSourceLabel =
        Table.AddColumn(
            WithClassifierSourceType,
            "ClassifierSourceLabel",
            each
                let
                    OverrideCategory = NormalizeText([OverrideCategoryId]),
                    OverrideLabel = NormalizeText([OverrideClassifierSourceLabel]),
                    Rule = [RuleRecord]
                in
                    if OverrideCategory <> ""
                    then OverrideLabel
                    else if Rule <> null
                    then
                        NormalizeText(Record.FieldOrDefault(Rule, "MatchField", "Description"))
                            & " contains "
                            & NormalizeText(Record.FieldOrDefault(Rule, "MatchText", ""))
                    else null,
            type nullable text
        ),
    WithClassifierRuleId =
        Table.AddColumn(
            WithClassifierSourceLabel,
            "ClassifierRuleId",
            each
                let
                    OverrideCategory = NormalizeText([OverrideCategoryId]),
                    OverrideRule = NormalizeText([OverrideClassifierRuleId]),
                    Rule = [RuleRecord]
                in
                    if OverrideCategory <> ""
                    then OverrideRule
                    else if Rule <> null
                    then NormalizeText(Record.FieldOrDefault(Rule, "RuleId", ""))
                    else null,
            type nullable text
        ),
    WithReviewStatus =
        Table.AddColumn(
            WithClassifierRuleId,
            "ClassificationReviewStatus",
            each NormalizeText([OverrideReviewStatus]),
            type nullable text
        ),
    Cleaned =
        Table.RemoveColumns(
            WithReviewStatus,
            {
                "OverrideCategoryId",
                "OverrideCategoryName",
                "OverrideClassifierSourceType",
                "OverrideClassifierSourceLabel",
                "OverrideClassifierRuleId",
                "OverrideReason",
                "OverrideReviewStatus",
                "RuleRecord"
            }
        )
in
    Cleaned
