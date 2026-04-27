[CmdletBinding()]
param(
    [string]$OutputDirectory = "release_artifacts\governance-starter"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$artifactRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot "release_artifacts"))

if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $resolvedOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
} else {
    $resolvedOutputDirectory = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory))
}

if (-not $resolvedOutputDirectory.StartsWith($artifactRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputDirectory must stay under release_artifacts: $resolvedOutputDirectory"
}

$scratchDirectory = Join-Path $resolvedOutputDirectory "_scratch"
$coreWorkbookPath = Join-Path $scratchDirectory "Governance_Starter_Core.xlsx"
$starterWorkbookPath = Join-Path $resolvedOutputDirectory "Governance_Starter.xlsx"
$templateWorkbookPath = Join-Path $resolvedOutputDirectory "Governance_Starter.xltx"
$assetEvidenceInstaller = Join-Path $PSScriptRoot "install_asset_evidence_pq_workbook.ps1"

New-Item -ItemType Directory -Path $resolvedOutputDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $scratchDirectory -Force | Out-Null

foreach ($path in @($coreWorkbookPath, $starterWorkbookPath, $templateWorkbookPath)) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force
    }
}

function Resolve-RepoPath {
    param([string]$RelativePath)
    return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $RelativePath))
}

function Read-TsvMatrix {
    param([string]$RelativePath)

    $path = Resolve-RepoPath $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing TSV source: $path"
    }

    $lines = [System.IO.File]::ReadAllLines($path, [System.Text.Encoding]::UTF8)
    $rows = New-Object 'object[][]' $lines.Count
    for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
        $cells = $lines[$lineIndex] -split "`t", -1
        $values = New-Object 'object[]' $cells.Count
        for ($cellIndex = 0; $cellIndex -lt $cells.Count; $cellIndex++) {
            $values[$cellIndex] = Convert-CellValue $cells[$cellIndex]
        }
        $rows[$lineIndex] = $values
    }
    return ,$rows
}

function Convert-CellValue {
    param([string]$Value)

    if ($null -eq $Value) { return "" }
    $text = $Value.Trim()
    if ($text -eq "") { return "" }
    if ($text -eq "TRUE") { return $true }
    if ($text -eq "FALSE") { return $false }
    $number = 0.0
    if ([double]::TryParse($text, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    return $Value
}

function Write-Matrix {
    param(
        [object]$Worksheet,
        [string]$TopLeft,
        [object]$Rows
    )

    if ($Rows.Count -eq 0) {
        return $null
    }

    $rowCount = $Rows.Count
    $columnCount = $Rows[0].Count
    $range = $Worksheet.Range($TopLeft).Resize($rowCount, $columnCount)
    $data = New-Object 'object[,]' $rowCount, $columnCount

    for ($row = 0; $row -lt $rowCount; $row++) {
        for ($column = 0; $column -lt $columnCount; $column++) {
            $data[$row, $column] = $Rows[$row][$column]
        }
    }

    $range.Value2 = $data
    return ,$range
}

function Add-Worksheet {
    param(
        [object]$Workbook,
        [string]$Name
    )

    foreach ($sheet in $Workbook.Worksheets) {
        if ([string]::Compare($sheet.Name, $Name, $true) -eq 0) {
            return $sheet
        }
    }

    $sheet = $Workbook.Worksheets.Add([System.Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    $sheet.Name = $Name
    return $sheet
}

function Add-TableFromMatrix {
    param(
        [object]$Worksheet,
        [string]$TableName,
        [string]$TopLeft,
        [object]$Rows,
        [string]$Style = "TableStyleMedium2"
    )

    $range = Write-Matrix -Worksheet $Worksheet -TopLeft $TopLeft -Rows $Rows
    $table = $Worksheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $TableName
    $table.TableStyle = $Style
    Format-TableHeader -Table $table
    return $table
}

function Format-TableHeader {
    param([object]$Table)

    try {
        [void]($Table.HeaderRowRange.Font.Bold = $true)
        [void]($Table.HeaderRowRange.Font.Color = 0)
        [void]($Table.HeaderRowRange.Interior.Color = 16247773)
    } catch {
        Write-Warning "Skipped table header formatting for $($Table.Name): $($_.Exception.Message)"
    }
}

function Column-Name {
    param([int]$Index)

    $name = ""
    $value = $Index
    while ($value -gt 0) {
        $value--
        $name = [string][char]([int](65 + ($value % 26))) + $name
        $value = [int][math]::Floor($value / 26)
    }
    return $name
}

function Add-ValidationList {
    param(
        [object]$Range,
        [string]$Source,
        [switch]$AllowUnknown
    )

    try {
        [void]$Range.Validation.Delete()
        [void]$Range.Validation.Add(3, 1, 1, $Source)
        if ($AllowUnknown) {
            $Range.Validation.ShowError = $false
        }
    } catch {
        Write-Warning "Skipped validation on $($Range.Address()): $($_.Exception.Message)"
    }
}

function Add-NonNegativeValidation {
    param([object]$Range)

    try {
        [void]$Range.Validation.Delete()
        [void]$Range.Validation.Add(2, 1, 7, "0")
    } catch {
        Write-Warning "Skipped non-negative validation on $($Range.Address()): $($_.Exception.Message)"
    }
}

function Set-NumberFormat {
    param(
        [object]$Range,
        [string]$Format
    )

    [void]($Range.NumberFormat = $Format)
}

function Set-RangeFormula {
    param(
        [object]$Range,
        [string]$Formula
    )

    try {
        $Range.Formula2 = $Formula
    } catch {
        $Range.Formula = $Formula
    }
}

function Set-FreezeRows {
    param(
        [object]$Excel,
        [object]$Worksheet,
        [int]$Rows
    )

    try {
        [void]$Worksheet.Activate()
        [void]($Excel.ActiveWindow.FreezePanes = $false)
        [void]$Worksheet.Range("A$($Rows + 1)").Select()
        [void]($Excel.ActiveWindow.FreezePanes = $true)
    } catch {
        Write-Warning "Skipped freeze panes on $($Worksheet.Name): $($_.Exception.Message)"
    }
}

function Remove-BlockComments {
    param([string]$Text)
    return [regex]::Replace($Text, "/\*[\s\S]*?\*/", "")
}

function Compact-FormulaBody {
    param([string]$Text)

    $clean = Remove-BlockComments $Text
    $builder = [System.Text.StringBuilder]::new()
    $inString = $false
    $inSheetQuote = $false
    $index = 0

    while ($index -lt $clean.Length) {
        $char = $clean[$index]

        if ($char -eq '"') {
            [void]$builder.Append($char)
            if ($inString -and $index + 1 -lt $clean.Length -and $clean[$index + 1] -eq '"') {
                $index++
                [void]$builder.Append($clean[$index])
            } else {
                $inString = -not $inString
            }
            $index++
            continue
        }

        if (-not $inString -and $char -eq "'") {
            $inSheetQuote = -not $inSheetQuote
            [void]$builder.Append($char)
            $index++
            continue
        }

        if (-not $inString -and -not $inSheetQuote -and [char]::IsWhiteSpace($char)) {
            $index++
            continue
        }

        [void]$builder.Append($char)
        $index++
    }

    return $builder.ToString().Trim()
}

function Parse-FormulaModule {
    param([string]$RelativePath)

    $path = Resolve-RepoPath $RelativePath
    $source = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8).Replace("`r", "")
    $matches = [regex]::Matches($source, "(?m)^([A-Za-z_][A-Za-z0-9_]*)\s*=")
    $items = @()

    for ($index = 0; $index -lt $matches.Count; $index++) {
        $match = $matches[$index]
        $name = $match.Groups[1].Value
        $start = $match.Index + $match.Length
        $end = if ($index + 1 -lt $matches.Count) { $matches[$index + 1].Index } else { $source.Length }
        $body = $source.Substring($start, $end - $start).Trim()
        $body = [regex]::Replace($body, ";\s*$", "").Trim()
        $items += [pscustomobject]@{
            Name = $name
            Formula = "=" + (Compact-FormulaBody $body)
        }
    }

    return $items
}

function Set-WorkbookName {
    param(
        [object]$Workbook,
        [string]$Name,
        [string]$RefersTo,
        [string]$Comment
    )

    try {
        [void]$Workbook.Names.Item($Name).Delete()
    } catch {
    }

    $definedName = $Workbook.Names.Add($Name, $RefersTo)
    try {
        $definedName.Comment = $Comment
    } catch {
    }
}

function Install-FormulaModules {
    param([object]$Workbook)

    $moduleFiles = @(
        @{ Prefix = "Controls"; Path = "modules\controls.formula.txt" },
        @{ Prefix = "get"; Path = "modules\get.formula.txt" },
        @{ Prefix = "kind"; Path = "modules\kind.formula.txt" },
        @{ Prefix = "CapitalPlanning"; Path = "modules\capital_planning_report.formula.txt" },
        @{ Prefix = "Analysis"; Path = "modules\analysis.formula.txt" },
        @{ Prefix = "defer"; Path = "modules\defer.formula.txt" },
        @{ Prefix = "Notes"; Path = "modules\notes.formula.txt" },
        @{ Prefix = "Phasing"; Path = "modules\phasing.formula.txt" },
        @{ Prefix = "Ready"; Path = "modules\ready.formula.txt" },
        @{ Prefix = "Search"; Path = "modules\search.formula.txt" },
        @{ Prefix = "Assets"; Path = "modules\assets.formula.txt" }
    )

    $aliases = @{}
    $installedCount = 0

    foreach ($module in $moduleFiles) {
        $parsed = Parse-FormulaModule $module.Path
        foreach ($item in $parsed) {
            Set-WorkbookName `
                -Workbook $Workbook `
                -Name "$($module.Prefix).$($item.Name)" `
                -RefersTo $item.Formula `
                -Comment "Governed formula module: $($module.Prefix)"
            $installedCount++

            if (-not $aliases.ContainsKey($item.Name)) {
                $aliases[$item.Name] = $module.Prefix
                Set-WorkbookName `
                    -Workbook $Workbook `
                    -Name $item.Name `
                    -RefersTo $item.Formula `
                    -Comment "Governed formula compatibility alias: $($module.Prefix)"
                $installedCount++
            }
        }
    }

    Set-WorkbookName -Workbook $Workbook -Name "PM_Filter_Dropdowns" -RefersTo "='Planning Review'!`$B`$2" -Comment "Governed formula visible control: Planning Review!B2"
    Set-WorkbookName -Workbook $Workbook -Name "Future_Filter_Mode" -RefersTo "='Planning Review'!`$C`$2" -Comment "Governed formula visible control: Planning Review!C2"
    Set-WorkbookName -Workbook $Workbook -Name "HideClosed_Status" -RefersTo "='Planning Review'!`$D`$2" -Comment "Governed formula visible control: Planning Review!D2"
    Set-WorkbookName -Workbook $Workbook -Name "Burndown_Cut_Target" -RefersTo "='Planning Review'!`$E`$2" -Comment "Governed formula visible control: Planning Review!E2"

    return $installedCount
}

$dropdownLists = [ordered]@{
    months = @("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    groupFields = @("Revised Group", "Site", "Region", "PM", "BU", "Category")
    futureFilters = @("All", "Exclude Future", "Keep F1 Only", "Keep F1+F2")
    closedRows = @("SHOW", "HIDE")
    statuses = @("Active", "Hold", "Closed", "In Service", "Skipping", "Canceled")
    yesNo = @("Y", "N")
    assetStatuses = @("planned", "active", "in_service", "maintenance", "retired")
    assetConditions = @("new", "good", "fair", "poor", "critical")
    assetCriticalities = @("low", "medium", "high", "critical")
    assetChangeTypes = @("new_asset", "replace_asset", "upgrade_asset")
    assetStates = @("mapped", "planned", "installed", "retired")
    assetPromotionStatuses = @("draft", "review", "accepted", "ready", "project_ready", "rejected")
    assetMappingStatuses = @("draft", "active", "ready", "needs_review", "inactive")
    assetChangeStatuses = @("draft", "ready", "applied", "needs_review", "blocked")
}

$validationColumns = @(
    @{ Key = "months"; Header = "Month" },
    @{ Key = "groupFields"; Header = "Group Field" },
    @{ Key = "futureFilters"; Header = "Future Filter" },
    @{ Key = "closedRows"; Header = "Closed Rows" },
    @{ Key = "statuses"; Header = "Status" },
    @{ Key = "yesNo"; Header = "Yes No" },
    @{ Key = "assetStatuses"; Header = "Asset Status" },
    @{ Key = "assetConditions"; Header = "Asset Condition" },
    @{ Key = "assetCriticalities"; Header = "Asset Criticality" },
    @{ Key = "assetChangeTypes"; Header = "Asset Change Type" },
    @{ Key = "assetStates"; Header = "Asset State" },
    @{ Key = "assetPromotionStatuses"; Header = "Asset Promotion Status" },
    @{ Key = "assetMappingStatuses"; Header = "Asset Mapping Status" },
    @{ Key = "assetChangeStatuses"; Header = "Asset Change Status" }
)

function Get-ListValidationSource {
    param([string]$ListKey)

    for ($index = 0; $index -lt $validationColumns.Count; $index++) {
        if ($validationColumns[$index].Key -eq $ListKey) {
            $column = Column-Name ($index + 1)
            $endRow = $dropdownLists[$ListKey].Count + 1
            return "='Validation Lists'!`$$column`$2:`$$column`$$endRow"
        }
    }

    throw "Unknown validation list: $ListKey"
}

function Get-TableColumnRange {
    param(
        [object]$Table,
        [string]$Header
    )

    try {
        $range = $Table.ListColumns.Item($Header).DataBodyRange
        return ,$range
    } catch {
        return $null
    }
}

function Apply-TableListValidation {
    param(
        [object]$Table,
        [string]$Header,
        [string]$ListKey,
        [switch]$AllowUnknown
    )

    $range = Get-TableColumnRange -Table $Table -Header $Header
    if ($null -ne $range) {
        [void](Add-ValidationList -Range $range -Source (Get-ListValidationSource $ListKey) -AllowUnknown:$AllowUnknown)
    }
}

function Apply-PlanningTableValidation {
    param([object]$Worksheet)

    $headers = @((Read-TsvMatrix "samples\planning_table_starter.tsv")[0])
    $rules = @(
        @{ Header = "Status"; List = "statuses" },
        @{ Header = "Chargeable"; List = "yesNo" },
        @{ Header = "Internal Eligible"; List = "yesNo" },
        @{ Header = "Canceled"; List = "yesNo" }
    )

    foreach ($rule in $rules) {
        $index = [array]::IndexOf($headers, $rule.Header)
        if ($index -ge 0) {
            $column = Column-Name ($index + 1)
            [void](Add-ValidationList -Range $Worksheet.Range("$column`3:$column`2000") -Source (Get-ListValidationSource $rule.List))
        }
    }
}

function Build-ValidationLists {
    param([object]$Worksheet)

    $maxCount = ($validationColumns | ForEach-Object { $dropdownLists[$_.Key].Count } | Measure-Object -Maximum).Maximum
    $rows = @()
    $rows += ,@($validationColumns | ForEach-Object { $_.Header })

    for ($row = 0; $row -lt $maxCount; $row++) {
        $values = @()
        foreach ($column in $validationColumns) {
            $list = $dropdownLists[$column.Key]
            $values += if ($row -lt $list.Count) { $list[$row] } else { "" }
        }
        $rows += ,$values
    }

    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "A1" -Rows $rows)
    $headerRange = $Worksheet.Range("A1").Resize(1, $validationColumns.Count)
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 16247773
    [void]$Worksheet.Columns.AutoFit()
}

function Build-AssetRelationshipLists {
    param([object]$Worksheet)

    $startColumn = $validationColumns.Count + 1
    $assetColumn = Column-Name $startColumn
    $projectColumn = Column-Name ($startColumn + 1)
    $Worksheet.Range("$assetColumn`1").Value2 = "Asset ID"
    $Worksheet.Range("$projectColumn`1").Value2 = "Project Key"
    try {
        [void](Set-RangeFormula -Range $Worksheet.Range("$assetColumn`2") -Formula '=LET(ids,VSTACK(tblAssets[AssetID],tblProjectAssetMap[AssetId],tblAssetChanges[SourceAssetId],tblAssetChanges[TargetAssetId],tblAssetStateHistory[AssetId]),IFERROR(SORT(UNIQUE(FILTER(ids,ids<>""))),""))')
        [void](Set-RangeFormula -Range $Worksheet.Range("$projectColumn`2") -Formula '=LET(keys,VSTACK(tblAssets[LinkedProjectID],tblSemanticAssets[ProjectKey],tblAssetPromotionQueue[ProjectKey],tblAssetMappingStaging[ProjectKey],tblProjectAssetMap[ProjectKey],tblAssetChanges[ProjectKey],tblAssetStateHistory[ProjectKey]),IFERROR(SORT(UNIQUE(FILTER(keys,keys<>""))),""))')
    } catch {
        Write-Warning "Skipped asset relationship spill formulas: $($_.Exception.Message)"
    }
    $Worksheet.Range("$assetColumn`1:$projectColumn`1").Font.Bold = $true
    $Worksheet.Range("$assetColumn`1:$projectColumn`1").Interior.Color = 16247773
    [void]$Worksheet.Columns.AutoFit()
}

function Build-PlanningReview {
    param([object]$Worksheet)

    $Worksheet.Range("A1:N3").Clear()
    $Worksheet.Range("A1").Value2 = "Planning Review"
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "B1" -Rows @(
        ,@("Group", "Future Filter", "Closed Rows", "Burndown Cut Target")
    ))
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "A2" -Rows @(
        ,@("Controls", "BU", "All", "SHOW", 0)
    ))
    $Worksheet.Range("A3").Value2 = "Main report spill starts at A4. Columns O:R are reserved for notes."
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "M1" -Rows @(
        ,@("Report As Of", "Defer As Of"),
        ,@("Mar", "Mar")
    ))

    $Worksheet.Range("A1").Font.Bold = $true
    $Worksheet.Range("A1").Font.Size = 16
    $Worksheet.Range("B1:E1").Font.Bold = $true
    $Worksheet.Range("M1:N1").Font.Bold = $true
    $Worksheet.Range("A2:E2").Interior.Color = 16448250
    $Worksheet.Range("B2:E2").Interior.Color = 13431551
    $Worksheet.Range("M2:N2").Interior.Color = 13431551
    [void](Set-NumberFormat -Range $Worksheet.Range("E2") -Format "$#,##0")

    [void](Add-ValidationList -Range $Worksheet.Range("B2") -Source (Get-ListValidationSource "groupFields"))
    [void](Add-ValidationList -Range $Worksheet.Range("C2") -Source (Get-ListValidationSource "futureFilters"))
    [void](Add-ValidationList -Range $Worksheet.Range("D2") -Source (Get-ListValidationSource "closedRows"))
    [void](Add-ValidationList -Range $Worksheet.Range("M2:N2") -Source (Get-ListValidationSource "months"))
    [void](Add-NonNegativeValidation -Range $Worksheet.Range("E2"))

    [void]$Worksheet.Columns.AutoFit()
}

function Build-PlanningReviewNotes {
    param([object]$Worksheet)

    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "O1" -Rows @(
        ,@("ApplyNotes Control", "Run 1: Prepare", "Run 2: Apply", "After Apply"),
        ,@("Type updates in P:R", "Run ApplyNotes once", "Run ApplyNotes again", "P:R clears"),
        ,@("Check Decision Staging", "Rows should say Prepared", "Prepared rows write back", "Column O refreshes"),
        ,@("ExistingMeetingNotes", "NewPlanningNotes", "NewTimeline", "NewStatus")
    ))
    $Worksheet.Range("O1:R3").Interior.Color = 14873826
    $Worksheet.Range("O1:R4").Font.Bold = $true
    $Worksheet.Range("O5").Formula2 = '=IFERROR(Notes.Existing,"")'
    $Worksheet.Range("O5:O200").Interior.Color = 16448250
    $Worksheet.Range("P5:R200").Interior.Color = 13431551
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "P5" -Rows @(
        ,@("Review forecast against latest meeting note", "Apr", "Review")
    ))
    [void](Add-ValidationList -Range $Worksheet.Range("R5:R200") -Source (Get-ListValidationSource "statuses"))
    [void]$Worksheet.Columns.AutoFit()
}

function Build-AutomationSetup {
    param([object]$Worksheet)

    $Worksheet.Range("A1:F40").Clear()
    $Worksheet.Range("A1").Value2 = "Automation Setup"
    $Worksheet.Range("A2").Value2 = "Optional Office Scripts are distributed as source files. Import them into Excel Automate when writeback automation is wanted."
    $Worksheet.Range("A1").Font.Bold = $true
    $Worksheet.Range("A1").Font.Size = 16
    $Worksheet.Range("A2").WrapText = $true

    $automationRows = New-Object 'object[][]' 7
    $automationRows[0] = [object[]]@("Step", "Action", "Why it matters")
    $automationRows[1] = [object[]]@("1", "Download ApplyNotes.ts from the GitHub release assets.", "The script is source-controlled and shipped separately from the workbook template.")
    $automationRows[2] = [object[]]@("2", "In Excel, open Automate > New Script.", "Office Scripts are created and stored through the user's Microsoft 365 account.")
    $automationRows[3] = [object[]]@("3", "Replace the default script with the contents of ApplyNotes.ts.", "This keeps script installation explicit instead of auto-installing tenant automation.")
    $automationRows[4] = [object[]]@("4", "Save the script as ApplyNotes.", "The operator can then run the same reviewed script against this workbook.")
    $automationRows[5] = [object[]]@("5", "Run ApplyNotes once to prepare Decision Staging, inspect ApplyMessage, then run it again to apply.", "The two-pass flow prevents hidden writeback and keeps staged changes reviewable.")
    $automationRows[6] = [object[]]@("Alternate", "When using the add-in, click Copy ApplyNotes Script and paste it into Automate > New Script.", "The add-in can help copy source text, but it does not install scripts automatically.")
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblAutomationSetup" -TopLeft "A4" -Rows $automationRows)

    $Worksheet.Range("A13").Value2 = "Release assets"
    $Worksheet.Range("A13").Font.Bold = $true
    $releaseRows = New-Object 'object[][]' 2
    $releaseRows[0] = [object[]]@("Governance_Starter.xltx", "Workbook template with sheets, tables, formulas, and Power Query outputs.")
    $releaseRows[1] = [object[]]@("ApplyNotes.ts", "Optional Office Script source for notes/status/timeline writeback.")
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "A14" -Rows $releaseRows)

    $Worksheet.Range("A18").Value2 = "Security boundary"
    $Worksheet.Range("A18").Font.Bold = $true
    $Worksheet.Range("A19").Value2 = "The public template does not embed VBA or auto-install Office Scripts. Users decide whether to import and run the optional script."
    $Worksheet.Range("A19").WrapText = $true
    $Worksheet.Range("A:F").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    $Worksheet.Columns.Item(2).ColumnWidth = 42
    $Worksheet.Columns.Item(3).ColumnWidth = 58
}

function Configure-DecisionStagingFormulas {
    param([object]$Table)

    $sourceRows = "DROP(Notes.FromArrayv,1)"
    $formulaColumns = @(
        @{ Header = "GroupType"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),2),""""))" },
        @{ Header = "GroupValue"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),3),""""))" },
        @{ Header = "Category"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),4),""""))" },
        @{ Header = "ProjDesc"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),5),""""))" },
        @{ Header = "AnnualProj"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),6),""""))" },
        @{ Header = "ActualsYTD"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),7),""""))" },
        @{ Header = "ExistingMeetingNotes"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),8),""""))" },
        @{ Header = "NewPlanningNotes"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),9),""""))" },
        @{ Header = "NewTimeline"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),10),""""))" },
        @{ Header = "NewStatus"; Formula = "=IF([@ReviewRow]="""","""",IFERROR(INDEX($sourceRows,XMATCH([@ReviewRow],CHOOSECOLS($sourceRows,1),0),11),""""))" },
        @{ Header = "ApplyAction"; Formula = '=IF(OR([@NewPlanningNotes]<>"",[@NewTimeline]<>"",[@NewStatus]<>""),"NOTE_TIMELINE_STATUS","")' },
        @{ Header = "PlanningNotes_New"; Formula = '=IF([@NewPlanningNotes]<>"",[@NewPlanningNotes],"")' },
        @{ Header = "Timeline_New"; Formula = '=IF([@NewTimeline]<>"",[@NewTimeline],"")' },
        @{ Header = "Comments_New"; Formula = '=""' },
        @{ Header = "Status_New"; Formula = '=IF([@NewStatus]<>"",[@NewStatus],"")' },
        @{ Header = "BudgetMatchCount"; Formula = '=IF([@ProjDesc]="","",SUMPRODUCT(--(INDEX(''Planning Table''!$A$3:$BM$200,,XMATCH("Project Description",''Planning Table''!$A$2:$BM$2,0))=[@ProjDesc])))' },
        @{ Header = "KeyStatus"; Formula = '=IF([@ProjDesc]="","",IF([@BudgetMatchCount]=1,"OK","BLOCKED"))' },
        @{ Header = "ApplyReady"; Formula = '=AND([@ProjDesc]<>"",[@BudgetMatchCount]=1,OR([@NewPlanningNotes]<>"",[@NewTimeline]<>"",[@NewStatus]<>""))' }
    )

    foreach ($item in $formulaColumns) {
        $range = Get-TableColumnRange -Table $Table -Header $item.Header
        if ($null -ne $range) {
            [void](Set-RangeFormula -Range $range -Formula $item.Formula)
        }
    }
}

function Build-AssetWorkflowTables {
    param(
        [object]$Workbook,
        [object]$Excel
    )

    $assetRegisterSheet = Add-Worksheet -Workbook $Workbook -Name "Asset Register"
    $assetRegister = Add-TableFromMatrix -Worksheet $assetRegisterSheet -TableName "tblAssets" -TopLeft "A1" -Rows (Read-TsvMatrix "samples\asset_register_starter.tsv")
    [void](Apply-TableListValidation -Table $assetRegister -Header "Status" -ListKey "assetStatuses")
    [void](Apply-TableListValidation -Table $assetRegister -Header "Condition" -ListKey "assetConditions")
    [void](Apply-TableListValidation -Table $assetRegister -Header "Criticality" -ListKey "assetCriticalities")

    $semanticSheet = Add-Worksheet -Workbook $Workbook -Name "Semantic Assets"
    [void](Add-TableFromMatrix -Worksheet $semanticSheet -TableName "tblSemanticAssets" -TopLeft "A1" -Rows (Read-TsvMatrix "samples\semantic_assets_starter.tsv"))

    $assetSetupSheet = Add-Worksheet -Workbook $Workbook -Name "Asset Setup"
    $assetSetupRows = Read-TsvMatrix "samples\asset_setup_starter.tsv"
    $promotionRows = @($assetSetupRows[0], $assetSetupRows[1])
    $mappingRows = @($assetSetupRows[2], $assetSetupRows[3])
    $promotionTable = Add-TableFromMatrix -Worksheet $assetSetupSheet -TableName "tblAssetPromotionQueue" -TopLeft "A1" -Rows $promotionRows
    $mappingTable = Add-TableFromMatrix -Worksheet $assetSetupSheet -TableName "tblAssetMappingStaging" -TopLeft "A6" -Rows $mappingRows

    $projectMapSheet = Add-Worksheet -Workbook $Workbook -Name "Project Asset Map"
    $projectMapTable = Add-TableFromMatrix -Worksheet $projectMapSheet -TableName "tblProjectAssetMap" -TopLeft "A1" -Rows (Read-TsvMatrix "samples\project_asset_map_starter.tsv")

    $assetChangesSheet = Add-Worksheet -Workbook $Workbook -Name "Asset Changes"
    $assetChangesTable = Add-TableFromMatrix -Worksheet $assetChangesSheet -TableName "tblAssetChanges" -TopLeft "A1" -Rows (Read-TsvMatrix "samples\asset_changes_starter.tsv")

    $assetHistorySheet = Add-Worksheet -Workbook $Workbook -Name "Asset State History"
    $assetHistoryTable = Add-TableFromMatrix -Worksheet $assetHistorySheet -TableName "tblAssetStateHistory" -TopLeft "A1" -Rows (Read-TsvMatrix "samples\asset_state_history_starter.tsv")

    foreach ($table in @($promotionTable, $mappingTable, $projectMapTable, $assetChangesTable, $assetHistoryTable)) {
        [void](Apply-TableListValidation -Table $table -Header "ChangeType" -ListKey "assetChangeTypes")
        [void](Apply-TableListValidation -Table $table -Header "ProposedChangeType" -ListKey "assetChangeTypes")
        [void](Apply-TableListValidation -Table $table -Header "InstalledState" -ListKey "assetStates")
        [void](Apply-TableListValidation -Table $table -Header "AssetState" -ListKey "assetStates")
        [void](Apply-TableListValidation -Table $table -Header "PromotionStatus" -ListKey "assetPromotionStatuses")
        [void](Apply-TableListValidation -Table $table -Header "MappingStatus" -ListKey "assetMappingStatuses")
        [void](Apply-TableListValidation -Table $table -Header "ChangeStatus" -ListKey "assetChangeStatuses")
        [void](Apply-TableListValidation -Table $table -Header "ApplyReady" -ListKey "yesNo")
    }

    foreach ($sheet in @($assetRegisterSheet, $semanticSheet, $assetSetupSheet, $projectMapSheet, $assetChangesSheet, $assetHistorySheet)) {
        [void](Set-FreezeRows -Excel $Excel -Worksheet $sheet -Rows 1)
        [void]$sheet.Columns.AutoFit()
    }
}

function Build-DemoOutputs {
    param([object]$Workbook)

    $outputs = @(
        @{ Sheet = "Planning Review"; Formula = "=CapitalPlanning.CAPITAL_PLANNING_REPORT()" },
        @{ Sheet = "BU Cap Scorecard"; Title = "BU Cap Scorecard"; Note = "Cap and spend posture by BU."; Formula = "=Analysis.BU_CAP_SCORECARD()" },
        @{ Sheet = "Reforecast Queue"; Title = "Reforecast Queue"; Note = "Grouped action queue for forecast review."; Formula = "=Analysis.REFORECAST_QUEUE()" },
        @{ Sheet = "PM Spend Report"; Title = "PM Spend Report"; Note = "Existing-work summary and job detail."; Formula = "=Analysis.PM_SPEND_REPORT()" },
        @{ Sheet = "Working Budget"; Title = "Working Budget Screen"; Note = "Current-job screening before budget drafting."; Formula = "=Analysis.WORKING_BUDGET_SCREEN()" },
        @{ Sheet = "Burndown"; Title = "Burndown Screen"; Note = "Meeting view of remaining burn and drivers."; Formula = "=Analysis.BURNDOWN_SCREEN()" },
        @{ Sheet = "Internal Jobs"; Title = "Internal Jobs Export"; Note = "Header-driven internal work export for readiness smoke testing."; Formula = "=Ready.InternalJobs_Export()" },
        @{ Sheet = "Asset Review"; Title = "Asset Review"; Note = "Asset workflow issue queues from source-controlled formula modules."; Formula = "=Assets.ASSET_MAPPING_ISSUES" }
    )

    foreach ($output in $outputs) {
        $sheet = Add-Worksheet -Workbook $Workbook -Name $output.Sheet
        if ($output.Sheet -ne "Planning Review") {
            $sheet.Range("A1:Z300").Clear()
            $sheet.Range("A1").Value2 = $output.Title
            $sheet.Range("A2").Value2 = $output.Note
            $sheet.Range("A1").Font.Bold = $true
            $sheet.Range("A1").Font.Size = 16
            $sheet.Range("A2").Font.Italic = $true
        }
        [void](Set-RangeFormula -Range $sheet.Range("A4") -Formula $output.Formula)
        [void]$sheet.Columns.AutoFit()
    }
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
        [void]$workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $planningSheet = $workbook.Worksheets.Item(1)
    $planningSheet.Name = "Planning Table"
    [void](Add-TableFromMatrix -Worksheet $planningSheet -TableName "tblPlanningTable" -TopLeft "A2" -Rows (Read-TsvMatrix "samples\planning_table_starter.tsv"))
    $planningSheet.Range("A2:BL2").Font.Bold = $true
    $planningSheet.Range("A2:BL2").Font.Color = 0
    $planningSheet.Range("A2:BL2").Interior.Color = 16247773
    foreach ($address in @("F2", "G2", "O2", "P2", "BE2")) {
        $planningSheet.Range($address).Interior.Color = 13431551
    }
    [void](Set-NumberFormat -Range $planningSheet.Range("O3:AZ234") -Format "$#,##0")
    [void](Set-NumberFormat -Range $planningSheet.Range("BJ3:BJ234") -Format "0")
    [void](Set-FreezeRows -Excel $excel -Worksheet $planningSheet -Rows 2)
    [void]$planningSheet.Columns.AutoFit()

    $capSheet = Add-Worksheet -Workbook $workbook -Name "Cap Setup"
    [void](Add-TableFromMatrix -Worksheet $capSheet -TableName "tblCapSetup" -TopLeft "A2" -Rows (Read-TsvMatrix "samples\cap_setup_starter.tsv"))
    $capSheet.Range("A2:B2").Font.Bold = $true
    $capSheet.Range("A2:B2").Font.Color = 0
    $capSheet.Range("A2:B2").Interior.Color = 16247773
    [void](Set-NumberFormat -Range $capSheet.Range("B3:B100") -Format "$#,##0")
    [void](Add-NonNegativeValidation -Range $capSheet.Range("B3:B100"))
    [void](Set-FreezeRows -Excel $excel -Worksheet $capSheet -Rows 2)
    [void]$capSheet.Columns.AutoFit()

    $validationSheet = Add-Worksheet -Workbook $workbook -Name "Validation Lists"
    [void](Build-ValidationLists -Worksheet $validationSheet)
    [void](Apply-PlanningTableValidation -Worksheet $planningSheet)

    $reviewSheet = Add-Worksheet -Workbook $workbook -Name "Planning Review"
    [void](Build-PlanningReview -Worksheet $reviewSheet)
    [void](Build-PlanningReviewNotes -Worksheet $reviewSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $reviewSheet -Rows 3)

    $automationSheet = Add-Worksheet -Workbook $workbook -Name "Automation Setup"
    [void](Build-AutomationSetup -Worksheet $automationSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $automationSheet -Rows 4)

    $stagingSheet = Add-Worksheet -Workbook $workbook -Name "Decision Staging"
    $stagingRows = Read-TsvMatrix "samples\decision_staging_starter.tsv"
    $blankStagingBody = @($stagingRows[0], @(1..$stagingRows[0].Count | ForEach-Object { "" }))
    $stagingTable = Add-TableFromMatrix -Worksheet $stagingSheet -TableName "tblDecisionStaging" -TopLeft "A1" -Rows $blankStagingBody
    $stagingTable.ListColumns.Item("ReviewRow").DataBodyRange.Value2 = 5
    [void](Configure-DecisionStagingFormulas -Table $stagingTable)
    [void](Set-FreezeRows -Excel $excel -Worksheet $stagingSheet -Rows 1)
    [void]$stagingSheet.Columns.AutoFit()

    Build-AssetWorkflowTables -Workbook $workbook -Excel $excel | Out-Null
    Build-AssetRelationshipLists -Worksheet $validationSheet | Out-Null

    $installedNames = Install-FormulaModules -Workbook $workbook
    Build-DemoOutputs -Workbook $workbook | Out-Null

    [void]$reviewSheet.Activate()
    [void]$workbook.SaveAs($coreWorkbookPath, 51)
    Write-Host "Built core governance workbook with $installedNames defined names: $coreWorkbookPath"
} finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

& $assetEvidenceInstaller `
    -TargetWorkbookPath $coreWorkbookPath `
    -OutputPath $starterWorkbookPath `
    -Force

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($starterWorkbookPath)
    [void]$workbook.Worksheets.Item("Planning Review").Activate()
    [void]$workbook.SaveAs($templateWorkbookPath, 54)
    Write-Host "Built governance starter workbook: $starterWorkbookPath"
    Write-Host "Built governance starter template: $templateWorkbookPath"
} finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if (Test-Path -LiteralPath $scratchDirectory) {
    Remove-Item -LiteralPath $scratchDirectory -Recurse -Force
}
