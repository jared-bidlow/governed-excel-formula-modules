let
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
    EmptyTyped = #table(
        type table [
            ProjectKey = nullable text,
            EvidenceId = nullable text,
            EvidenceType = nullable text,
            EvidencePath = nullable text,
            EvidenceName = nullable text,
            Extension = nullable text,
            DocumentAreaID = nullable text,
            DocumentAreaName = nullable text,
            CategoryID = nullable text,
            CategoryName = nullable text,
            DateModified = nullable text,
            ReviewStatus = nullable text,
            ApprovedOn = nullable text,
            ReviewerNotes = nullable text,
            StatusSignal = nullable text
        ],
        {}
    ),
    Config =
        try Excel.CurrentWorkbook(){[Name = "tblIntegrationBridgeConfig"]}[Content]
        otherwise #table({"Setting", "Value", "Notes"}, {}),
    GetConfigValue = (SettingName as text, DefaultValue as nullable text) as nullable text =>
        let
            MatchingRows = Table.SelectRows(
                Config,
                each
                    let
                        SettingValue = try [Setting] otherwise null,
                        SettingText = if SettingValue = null then "" else Text.From(SettingValue)
                    in
                        Text.Upper(Text.Trim(SettingText)) = Text.Upper(SettingName)
            ),
            RawValue =
                if Table.RowCount(MatchingRows) = 0 then
                    DefaultValue
                else
                    try Text.From(MatchingRows{0}[Value]) otherwise DefaultValue,
            TrimmedValue = if RawValue = null then null else Text.Trim(RawValue)
        in
            if TrimmedValue = null or TrimmedValue = "" then DefaultValue else TrimmedValue,
    IntegrationRepoRoot = GetConfigValue("IntegrationRepoRoot", ""),
    ApprovedProjectEvidenceCsvRelativePath =
        GetConfigValue("ApprovedProjectEvidenceCsvRelativePath", "data\exports\approved_project_evidence.csv"),
    CleanRoot =
        if IntegrationRepoRoot = null or IntegrationRepoRoot = "" or IntegrationRepoRoot = "<LOCAL_INTEGRATION_REPO_V1>" then
            ""
        else if Text.End(IntegrationRepoRoot, 1) = "\" or Text.End(IntegrationRepoRoot, 1) = "/" then
            Text.Start(IntegrationRepoRoot, Text.Length(IntegrationRepoRoot) - 1)
        else
            IntegrationRepoRoot,
    CleanRelativePath =
        if ApprovedProjectEvidenceCsvRelativePath = null or ApprovedProjectEvidenceCsvRelativePath = "" then
            "data\exports\approved_project_evidence.csv"
        else if Text.Start(ApprovedProjectEvidenceCsvRelativePath, 1) = "\" or Text.Start(ApprovedProjectEvidenceCsvRelativePath, 1) = "/" then
            Text.Range(ApprovedProjectEvidenceCsvRelativePath, 1)
        else
            ApprovedProjectEvidenceCsvRelativePath,
    CsvPath = if CleanRoot = "" then "" else CleanRoot & "\" & CleanRelativePath,
    CsvBinary = if CsvPath = "" then null else try File.Contents(CsvPath) otherwise null,
    Imported =
        if CsvBinary = null then
            EmptyTyped
        else
            try
                let
                    CsvRows = Csv.Document(CsvBinary, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
                    PromotedHeaders = Table.PromoteHeaders(CsvRows, [PromoteAllScalars = true]),
                    Shaped = Table.SelectColumns(PromotedHeaders, RequiredColumns, MissingField.UseNull),
                    Typed = Table.TransformColumnTypes(
                        Shaped,
                        List.Transform(RequiredColumns, each {_, type text})
                    )
                in
                    Typed
            otherwise
                EmptyTyped,
    ApprovedOnly = Table.SelectRows(
        Imported,
        each
            let
                ReviewStatus = try [ReviewStatus] otherwise "",
                ReviewStatusText = if ReviewStatus = null then "" else Text.From(ReviewStatus)
            in
                Text.Upper(Text.Trim(ReviewStatusText)) = "APPROVED"
    )
in
    ApprovedOnly
