[CmdletBinding()]
param(
    [string]$OutputDirectory = "release_artifacts\governance-starter",
    [ValidateSet("Planning", "AssetsLite", "AssetsFull", "SemanticTwin")]
    [string]$Edition = "Planning"
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
$artifactSuffix = if ($Edition -eq "Planning") { "" } else { "_$Edition" }
$coreWorkbookPath = Join-Path $scratchDirectory "Governance_Starter$artifactSuffix`_Core.xlsx"
$starterWorkbookPath = Join-Path $resolvedOutputDirectory "Governance_Starter$artifactSuffix.xlsx"
$templateWorkbookPath = Join-Path $resolvedOutputDirectory "Governance_Starter$artifactSuffix.xltx"
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

function Ensure-TableRows {
    param([object]$Rows)

    if ($Rows.Count -gt 1) {
        return ,$Rows
    }

    $blank = New-Object 'object[]' $Rows[0].Count
    for ($index = 0; $index -lt $Rows[0].Count; $index++) {
        $blank[$index] = ""
    }
    return ,@($Rows[0], $blank)
}

function New-BlankRowLike {
    param([object]$HeaderRow)

    $blank = New-Object 'object[]' $HeaderRow.Count
    for ($index = 0; $index -lt $HeaderRow.Count; $index++) {
        $blank[$index] = ""
    }
    return ,$blank
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

function Get-WorksheetOrNull {
    param(
        [object]$Workbook,
        [string]$Name
    )

    foreach ($sheet in $Workbook.Worksheets) {
        if ([string]::Compare($sheet.Name, $Name, $true) -eq 0) {
            return $sheet
        }
    }

    return $null
}

function Remove-WorkbookTableIfExists {
    param(
        [object]$Workbook,
        [string]$TableName
    )

    foreach ($sheet in $Workbook.Worksheets) {
        for ($index = $sheet.ListObjects.Count; $index -ge 1; $index--) {
            $table = $sheet.ListObjects.Item($index)
            if ([string]::Compare($table.Name, $TableName, $true) -eq 0) {
                [void]$table.Delete()
            }
        }
    }
}

function Add-TableFromMatrix {
    param(
        [object]$Worksheet,
        [string]$TableName,
        [string]$TopLeft,
        [object]$Rows,
        [string]$Style = "TableStyleMedium2"
    )

    Remove-WorkbookTableIfExists -Workbook $Worksheet.Parent -TableName $TableName
    $Rows = Ensure-TableRows -Rows $Rows
    $range = Write-Matrix -Worksheet $Worksheet -TopLeft $TopLeft -Rows $Rows
    try {
        $table = $Worksheet.ListObjects.Add(1, $range, $null, 1)
    } catch {
        throw "Failed to create table $TableName on sheet $($Worksheet.Name) at $TopLeft`: $($_.Exception.Message)"
    }
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
        [switch]$AllowUnknown,
        [string]$InputTitle = "",
        [string]$InputMessage = ""
    )

    try {
        [void]$Range.Validation.Delete()
        [void]$Range.Validation.Add(3, 1, 1, $Source)
        $Range.Validation.IgnoreBlank = $true
        if ($InputMessage -ne "") {
            $Range.Validation.ShowInput = $true
            $Range.Validation.InputTitle = $InputTitle
            $Range.Validation.InputMessage = $InputMessage
        }
        if ($AllowUnknown) {
            $Range.Validation.ShowError = $false
        }
    } catch {
        Write-Warning "Skipped validation on $($Range.Address()): $($_.Exception.Message)"
    }
}

function Add-NonNegativeValidation {
    param(
        [object]$Range,
        [string]$InputTitle = "",
        [string]$InputMessage = ""
    )

    try {
        [void]$Range.Validation.Delete()
        [void]$Range.Validation.Add(2, 1, 7, "0")
        $Range.Validation.IgnoreBlank = $true
        if ($InputMessage -ne "") {
            $Range.Validation.ShowInput = $true
            $Range.Validation.InputTitle = $InputTitle
            $Range.Validation.InputMessage = $InputMessage
        }
    } catch {
        Write-Warning "Skipped non-negative validation on $($Range.Address()): $($_.Exception.Message)"
    }
}

function Add-InputMessage {
    param(
        [object]$Range,
        [string]$InputTitle,
        [string]$InputMessage
    )

    try {
        [void]$Range.Validation.Delete()
        [void]$Range.Validation.Add(7, 1, 1, "=TRUE")
        $Range.Validation.IgnoreBlank = $true
        $Range.Validation.ShowInput = $true
        $Range.Validation.InputTitle = $InputTitle
        $Range.Validation.InputMessage = $InputMessage
        $Range.Validation.ShowError = $false
    } catch {
        Write-Warning "Skipped input message on $($Range.Address()): $($_.Exception.Message)"
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

function Set-InternalWorksheetLink {
    param(
        [object]$Range,
        [string]$SheetName,
        [string]$DisplayText,
        [string]$TargetCell = "A1"
    )

    $safeSheet = $SheetName.Replace("'", "''")
    $subAddress = "'$safeSheet'!$TargetCell"
    try {
        $Range.Hyperlinks.Delete()
    } catch {
        # Newly-created cells usually have no existing hyperlinks.
    }
    $Range.Value2 = $DisplayText
    [void]$Range.Worksheet.Hyperlinks.Add($Range, "", $subAddress, "", $DisplayText)
    $Range.Font.Color = 12611584
    $Range.Font.Underline = 2
}

function Set-MergedPanel {
    param(
        [object]$Worksheet,
        [string]$Address,
        [string]$Text,
        [int]$FillColor
    )

    $range = $Worksheet.Range($Address)
    $range.Merge() | Out-Null
    $range.Value2 = $Text
    $range.WrapText = $true
    $range.VerticalAlignment = -4160
    $range.Interior.Color = $FillColor
    $range.Borders.LineStyle = 1
    $range.Borders.Color = 14277081
}

function Normalize-GeneratedSheetRows {
    param(
        [object]$Worksheet,
        [int[]]$SectionRows = @(),
        [int]$DefaultHeight = 20
    )

    try {
        $Worksheet.UsedRange.Rows.RowHeight = $DefaultHeight
        $Worksheet.Rows.Item(1).RowHeight = 28
        $Worksheet.Rows.Item(2).RowHeight = 34
        $Worksheet.Rows.Item(3).RowHeight = 8
        foreach ($row in $SectionRows) {
            $Worksheet.Rows.Item($row).RowHeight = 24
            $Worksheet.Rows.Item($row + 1).RowHeight = 30
        }
    } catch {
        Write-Warning "Skipped row-height normalization on $($Worksheet.Name): $($_.Exception.Message)"
    }
}

function Apply-HubColumnWidthTemplate {
    param(
        [object]$Worksheet,
        [string]$Template
    )

    if ($Template -eq "SourceStatus") {
        $widths = @(24, 34, 18, 18, 18, 18)
    } elseif ($Template -eq "PlanningReview") {
        $widths = @(18, 18, 18, 16, 18, 16, 16, 16, 16, 16, 16, 16, 18, 18, 38, 36, 20, 20)
    } elseif ($Template -eq "AssetFinance") {
        $widths = @(24, 24, 22, 22, 20, 22, 20, 20, 20, 22, 24, 24, 18, 18, 18, 18)
    } elseif ($Template -eq "Asset") {
        $widths = @(26, 24, 36, 22, 22, 22, 22, 22, 22, 24, 24, 18, 18, 18, 18)
    } elseif ($Template -eq "Semantic") {
        $widths = @(26, 24, 26, 30, 26, 24, 24, 18, 18, 20, 22, 18, 18, 18, 18)
    } else {
        $widths = @(24, 22, 42, 18, 18, 18, 18, 18, 20, 20, 20, 18, 18, 18, 18)
    }

    for ($index = 0; $index -lt $widths.Count; $index++) {
        $Worksheet.Columns.Item($index + 1).ColumnWidth = $widths[$index]
    }
}

function Set-DocumentPropertyIfExists {
    param(
        [object]$Properties,
        [string]$Name,
        [string]$Value
    )

    try {
        $Properties.Item($Name).Value = $Value
    } catch {
        return
    }
}

function Set-PublicWorkbookProperties {
    param([object]$Workbook)

    $publicName = "Governed Excel Formula Modules"
    Set-DocumentPropertyIfExists -Properties $Workbook.BuiltinDocumentProperties -Name "Author" -Value $publicName
    Set-DocumentPropertyIfExists -Properties $Workbook.BuiltinDocumentProperties -Name "Last Author" -Value $publicName
    Set-DocumentPropertyIfExists -Properties $Workbook.BuiltinDocumentProperties -Name "Company" -Value $publicName
    Set-DocumentPropertyIfExists -Properties $Workbook.BuiltinDocumentProperties -Name "Manager" -Value ""
}

function Sanitize-WorkbookPackage {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $archive = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entries = @($archive.Entries | Where-Object {
            $_.FullName.EndsWith(".xml", [System.StringComparison]::OrdinalIgnoreCase) -or
            $_.FullName.EndsWith(".rels", [System.StringComparison]::OrdinalIgnoreCase)
        })
        foreach ($entry in $entries) {
            $entryName = $entry.FullName
            $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $text = $reader.ReadToEnd()
            } finally {
                $reader.Close()
            }

            $clean = [regex]::Replace($text, '(?s)<mc:AlternateContent>.*?<x15ac:absPath\b.*?</mc:AlternateContent>', '')
            $clean = [regex]::Replace($clean, '<x15ac:absPath\b[^>]*/>', '')
            $clean = [regex]::Replace($clean, '(?s)<dc:creator>.*?</dc:creator>', '<dc:creator>Governed Excel Formula Modules</dc:creator>')
            $clean = [regex]::Replace($clean, '(?s)<cp:lastModifiedBy>.*?</cp:lastModifiedBy>', '<cp:lastModifiedBy>Governed Excel Formula Modules</cp:lastModifiedBy>')
            $clean = [regex]::Replace($clean, '(?s)<Company>.*?</Company>', '<Company>Governed Excel Formula Modules</Company>')
            $clean = [regex]::Replace($clean, '(?s)<Manager>.*?</Manager>', '<Manager></Manager>')

            if ($clean -ne $text) {
                $entry.Delete()
                $newEntry = $archive.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
                $writer = [System.IO.StreamWriter]::new($newEntry.Open(), [System.Text.Encoding]::UTF8)
                try {
                    $writer.Write($clean)
                } finally {
                    $writer.Close()
                }
            }
        }
    } finally {
        $archive.Dispose()
    }
}

function Assert-WorkbookPackagePublic {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression | Out-Null
    Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
    $forbidden = @(
        ("C:" + "\Users"),
        ("/" + "Users/"),
        ("." + "codex"),
        ("ja" + "red"),
        ("C:" + "\Truth"),
        ("iCloud" + "Drive"),
        ("One" + "Drive"),
        ("public" + "-exports"),
        ("release" + "_artifacts")
    )
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($entry in $archive.Entries) {
            if (-not ($entry.FullName.EndsWith(".xml", [System.StringComparison]::OrdinalIgnoreCase) -or $entry.FullName.EndsWith(".rels", [System.StringComparison]::OrdinalIgnoreCase))) {
                continue
            }
            $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
            try {
                $text = $reader.ReadToEnd()
            } finally {
                $reader.Close()
            }
            foreach ($needle in $forbidden) {
                if ($text.IndexOf($needle, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    throw "Release artifact contains forbidden metadata '$needle' in $($entry.FullName): $Path"
                }
            }
        }
    } finally {
        $archive.Dispose()
    }
}

function Assert-NoVisibleWorkbookErrors {
    param([object]$Workbook)

    $blocked = @("#N/A", "#REF!", "#VALUE!", "#NAME?", "#DIV/0!")
    foreach ($sheet in $Workbook.Worksheets) {
        if ($sheet.Visible -ne -1) { continue }
        $used = $sheet.UsedRange
        foreach ($cell in $used.Cells) {
            $text = [string]$cell.Text
            if ($blocked -contains $text) {
                throw "Visible workbook error $text at $($sheet.Name)!$($cell.Address($false, $false))"
            }
        }
    }
}

function Color-Rgb {
    param(
        [int]$Red,
        [int]$Green,
        [int]$Blue
    )

    return $Red + ($Green * 256) + ($Blue * 65536)
}

function Format-PageHeader {
    param(
        [object]$Worksheet,
        [string]$Title,
        [string]$Subtitle,
        [string]$BandRange = "A1:I3"
    )

    $headerFill = Color-Rgb 31 78 121
    $subtitleFill = Color-Rgb 47 117 181
    $Worksheet.Cells.Font.Name = "Aptos"
    $band = $Worksheet.Range($BandRange)
    $band.Interior.Color = $headerFill
    $band.Font.Color = 16777215
    $band.WrapText = $true
    $Worksheet.Range("A1").Value2 = $Title
    $Worksheet.Range("A1").Font.Bold = $true
    $Worksheet.Range("A1").Font.Size = 20
    $Worksheet.Range("A2").Value2 = $Subtitle
    $Worksheet.Range("A2").Interior.Color = $subtitleFill
    $Worksheet.Range("A2").Font.Color = 16777215
    $Worksheet.Range("A2").Font.Size = 10
    $Worksheet.Range("A2").WrapText = $true
    $Worksheet.Rows.Item(1).RowHeight = 28
    $Worksheet.Rows.Item(2).RowHeight = 34
    $Worksheet.Rows.Item(3).RowHeight = 8
}

function Format-SectionHeader {
    param(
        [object]$Anchor,
        [string]$Title,
        [string]$Note
    )

    $sectionFill = Color-Rgb 226 239 218
    $noteFill = Color-Rgb 242 242 242
    $sectionRange = $Anchor.Resize(1, 8)
    $noteRange = $Anchor.Offset(1, 0).Resize(1, 8)
    $sectionRange.Interior.Color = $sectionFill
    $sectionRange.Font.Color = 0
    $sectionRange.Font.Bold = $true
    $sectionRange.Borders.LineStyle = 1
    $sectionRange.Borders.Color = Color-Rgb 169 208 142
    $Anchor.Value2 = $Title
    $noteRange.Interior.Color = $noteFill
    $noteRange.Font.Italic = $true
    $noteRange.Font.Color = Color-Rgb 89 89 89
    $noteRange.WrapText = $true
    $Anchor.Offset(1, 0).Value2 = $Note
}

function Add-HubSection {
    param(
        [object]$Worksheet,
        [string]$Cell,
        [string]$Title,
        [string]$Note,
        [string]$Formula
    )

    $anchor = $Worksheet.Range($Cell)
    Format-SectionHeader -Anchor $anchor -Title $Title -Note $Note
    [void](Set-RangeFormula -Range $anchor.Offset(3, 0) -Formula $Formula)
}

function Add-HubTableOfContents {
    param(
        [object]$Worksheet,
        [string]$TableName,
        [object[]]$Sections,
        [string]$TopLeft = "A4"
    )

    $rows = New-Object 'object[][]' ($Sections.Count + 1)
    $rows[0] = [object[]]@("Go to section", "What it shows")
    for ($index = 0; $index -lt $Sections.Count; $index++) {
        $rows[$index + 1] = [object[]]@($Sections[$index].Title, $Sections[$index].Note)
    }

    $table = Add-TableFromMatrix -Worksheet $Worksheet -TableName $TableName -TopLeft $TopLeft -Rows $rows
    for ($index = 0; $index -lt $Sections.Count; $index++) {
        $cell = $table.DataBodyRange.Cells.Item($index + 1, 1)
        Set-InternalWorksheetLink -Range $cell -SheetName $Worksheet.Name -DisplayText $Sections[$index].Title -TargetCell $Sections[$index].Cell
    }
    Set-HubTableOfContentsColumnWidths -Table $table
    return $table
}

function Set-HubTableOfContentsColumnWidths {
    param([object]$Table)

    $Table.Range.Columns.Item(1).ColumnWidth = 28
    $Table.Range.Columns.Item(2).ColumnWidth = 64
}

function Format-HubSheet {
    param(
        [object]$Worksheet,
        [string]$Title,
        [string]$Note,
        [string]$ClearRange = "A1:Z340"
    )

    $Worksheet.Range($ClearRange).Clear()
    Format-PageHeader -Worksheet $Worksheet -Title $Title -Subtitle $Note -BandRange "A1:Z3"
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

    try {
        $definedName = $Workbook.Names.Add($Name, $RefersTo)
    } catch {
        throw "Failed to install workbook name $Name`: $($_.Exception.Message)"
    }
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
        @{ Prefix = "Source"; Path = "modules\source.formula.txt" },
        @{ Prefix = "Assets"; Path = "modules\assets.formula.txt" },
        @{ Prefix = "AssetFinance"; Path = "modules\asset_finance.formula.txt" },
        @{ Prefix = "Ontology"; Path = "modules\ontology.formula.txt" }
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
    statuses = @("Active", "Review", "Hold", "Closed", "In Service", "Skipping", "Canceled")
    yesNo = @("Y", "N")
    booleanFlags = @("TRUE", "FALSE")
    assetTypes = @("Equipment", "Building", "Vehicle", "System", "Space", "Other")
    assetStatuses = @("planned", "active", "in_service", "maintenance", "retired")
    assetConditions = @("new", "good", "fair", "poor", "critical")
    assetCriticalities = @("low", "medium", "high", "critical")
    assetChangeTypes = @("new_asset", "replace_asset", "upgrade_asset")
    assetStates = @("mapped", "planned", "installed", "retired")
    assetPromotionStatuses = @("draft", "review", "accepted", "ready", "project_ready", "rejected")
    assetMappingStatuses = @("draft", "active", "ready", "needs_review", "inactive")
    assetChangeStatuses = @("draft", "ready", "applied", "needs_review", "blocked")
    assetWorkflowModes = @("Off", "Map existing assets", "Create candidate assets", "Track replacements/upgrades", "Asset finance from evidence")
}

$validationColumns = @(
    @{ Key = "months"; Header = "Month" },
    @{ Key = "groupFields"; Header = "Group Field" },
    @{ Key = "futureFilters"; Header = "Future Filter" },
    @{ Key = "closedRows"; Header = "Closed Rows" },
    @{ Key = "statuses"; Header = "Status" },
    @{ Key = "yesNo"; Header = "Yes No" },
    @{ Key = "booleanFlags"; Header = "Boolean Flag" },
    @{ Key = "assetTypes"; Header = "Asset Type" },
    @{ Key = "assetStatuses"; Header = "Asset Status" },
    @{ Key = "assetConditions"; Header = "Asset Condition" },
    @{ Key = "assetCriticalities"; Header = "Asset Criticality" },
    @{ Key = "assetChangeTypes"; Header = "Asset Change Type" },
    @{ Key = "assetStates"; Header = "Asset State" },
    @{ Key = "assetPromotionStatuses"; Header = "Asset Promotion Status" },
    @{ Key = "assetMappingStatuses"; Header = "Asset Mapping Status" },
    @{ Key = "assetChangeStatuses"; Header = "Asset Change Status" },
    @{ Key = "assetWorkflowModes"; Header = "Asset Workflow Mode" }
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
        [switch]$AllowUnknown,
        [string]$InputTitle = "",
        [string]$InputMessage = ""
    )

    $range = Get-TableColumnRange -Table $Table -Header $Header
    if ($null -ne $range) {
        [void](Add-ValidationList -Range $range -Source (Get-ListValidationSource $ListKey) -AllowUnknown:$AllowUnknown -InputTitle $InputTitle -InputMessage $InputMessage)
    }
}

function Apply-TableSourceValidation {
    param(
        [object]$Table,
        [string]$Header,
        [string]$Source,
        [switch]$AllowUnknown,
        [string]$InputTitle = "",
        [string]$InputMessage = ""
    )

    $range = Get-TableColumnRange -Table $Table -Header $Header
    if ($null -ne $range) {
        [void](Add-ValidationList -Range $range -Source $Source -AllowUnknown:$AllowUnknown -InputTitle $InputTitle -InputMessage $InputMessage)
    }
}

function Apply-TableInputMessage {
    param(
        [object]$Table,
        [string]$Header,
        [string]$InputTitle,
        [string]$InputMessage
    )

    $range = Get-TableColumnRange -Table $Table -Header $Header
    if ($null -ne $range) {
        [void](Add-InputMessage -Range $range -InputTitle $InputTitle -InputMessage $InputMessage)
    }
}

function Apply-TableNonNegativeValidation {
    param(
        [object]$Table,
        [string]$Header,
        [string]$InputTitle = "",
        [string]$InputMessage = ""
    )

    $range = Get-TableColumnRange -Table $Table -Header $Header
    if ($null -ne $range) {
        [void](Add-NonNegativeValidation -Range $range -InputTitle $InputTitle -InputMessage $InputMessage)
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
        [void](Set-RangeFormula -Range $Worksheet.Range("$projectColumn`2") -Formula '=LET(keys,VSTACK(tblPlanningTable[Source ID],tblPlanningTable[Job ID],tblPlanningTable[Project Description],tblAssets[LinkedProjectID],tblSemanticAssets[ProjectKey],tblAssetPromotionQueue[ProjectKey],tblAssetMappingStaging[ProjectKey],tblProjectAssetMap[ProjectKey],tblAssetChanges[ProjectKey],tblAssetStateHistory[ProjectKey]),IFERROR(SORT(UNIQUE(FILTER(keys,keys<>""))),""))')
    } catch {
        Write-Warning "Skipped asset relationship spill formulas: $($_.Exception.Message)"
    }
    $Worksheet.Range("$assetColumn`1:$projectColumn`1").Font.Bold = $true
    $Worksheet.Range("$assetColumn`1:$projectColumn`1").Interior.Color = 16247773
    Set-WorkbookName -Workbook $Worksheet.Parent -Name "Asset_ID_Dropdowns" -RefersTo "='Validation Lists'!`$$assetColumn`$2#" -Comment "Advisory asset ID dropdowns for optional asset workflow."
    Set-WorkbookName -Workbook $Worksheet.Parent -Name "Asset_Project_Key_Dropdowns" -RefersTo "='Validation Lists'!`$$projectColumn`$2#" -Comment "Advisory project key dropdowns for optional asset workflow."
    [void]$Worksheet.Columns.AutoFit()
}

function Get-AssetRelationshipValidationSource {
    param([string]$ListKey)

    $indexMap = @{
        assetIds = 0
        projectKeys = 1
    }
    if (-not $indexMap.ContainsKey($ListKey)) {
        throw "Unknown asset relationship list: $ListKey"
    }
    $column = Column-Name ($validationColumns.Count + 1 + $indexMap[$ListKey])
    return "='Validation Lists'!`$$column`$2#"
}

function Apply-AssetRegisterValidation {
    param([object]$Table)

    Apply-TableInputMessage `
        -Table $Table `
        -Header "AssetID" `
        -InputTitle "AssetID" `
        -InputMessage "Enter a stable asset identifier, e.g. AHU-001."
    Apply-TableInputMessage `
        -Table $Table `
        -Header "AssetName" `
        -InputTitle "AssetName" `
        -InputMessage "Enter a plain-English asset name."
    Apply-TableSourceValidation `
        -Table $Table `
        -Header "AssetType" `
        -Source ($dropdownLists["assetTypes"] -join ",") `
        -InputTitle "AssetType" `
        -InputMessage "Choose a simple type such as Equipment, Building, Vehicle, System, Space, or Other."
    Apply-TableSourceValidation `
        -Table $Table `
        -Header "Status" `
        -Source ($dropdownLists["assetStatuses"] -join ",") `
        -InputTitle "Status" `
        -InputMessage "Choose the current lifecycle status."
    Apply-TableSourceValidation -Table $Table -Header "Condition" -Source ($dropdownLists["assetConditions"] -join ",")
    Apply-TableSourceValidation -Table $Table -Header "Criticality" -Source ($dropdownLists["assetCriticalities"] -join ",")
    Apply-TableNonNegativeValidation `
        -Table $Table `
        -Header "ReplacementCost" `
        -InputTitle "ReplacementCost" `
        -InputMessage "Enter a non-negative replacement cost, or leave blank."
    Apply-TableNonNegativeValidation `
        -Table $Table `
        -Header "UsefulLifeYears" `
        -InputTitle "UsefulLifeYears" `
        -InputMessage "Enter a non-negative useful life in years, or leave blank."
}

function Apply-AssetRegisterRelationshipValidation {
    param([object]$Workbook)

    try {
        $table = $Workbook.Worksheets.Item("Asset Register").ListObjects.Item("tblAssets")
        Apply-TableSourceValidation `
            -Table $table `
            -Header "LinkedProjectID" `
            -Source "=Asset_Project_Key_Dropdowns" `
            -AllowUnknown `
            -InputTitle "LinkedProjectID" `
            -InputMessage "Optional project/job key from the current workbook planning data. This does not imply external refresh or sync."
    } catch {
        Write-Warning "Skipped Asset Register relationship validation: $($_.Exception.Message)"
    }
}

function Build-PlanningReview {
    param([object]$Worksheet)

    $Worksheet.Range("A1:N3").Clear()
    $Worksheet.Range("A1").Value2 = "Planning Review"
    $reviewHeaderRows = New-Object 'object[][]' 1
    $reviewHeaderRows[0] = [object[]]@("Group", "Future Filter", "Closed Rows", "Burndown Cut Target")
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "B1" -Rows $reviewHeaderRows)
    $reviewControlRows = New-Object 'object[][]' 1
    $reviewControlRows[0] = [object[]]@("Controls", "BU", "All", "SHOW", 0)
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "A2" -Rows $reviewControlRows)
    $Worksheet.Range("A3").Value2 = "Main report spill starts at A4. Columns O:R are reserved for notes."
    $reviewMonthRows = New-Object 'object[][]' 2
    $reviewMonthRows[0] = [object[]]@("Report As Of Month", "Defer As Of Month")
    $reviewMonthRows[1] = [object[]]@("Mar", "Mar")
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "M1" -Rows $reviewMonthRows)

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
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "PlanningReview"
}

function Build-PlanningReviewNotes {
    param([object]$Worksheet)

    $notesFlowRows = New-Object 'object[][]' 3
    $notesFlowRows[0] = [object[]]@("ApplyNotes Control", "Run 1: Prepare", "Run 2: Apply", "After Apply")
    $notesFlowRows[1] = [object[]]@("Type updates in P:R", "Run ApplyNotes once", "Run ApplyNotes again", "P:R clears")
    $notesFlowRows[2] = [object[]]@("Check Decision Staging", "Rows should say Prepared", "Prepared rows write back", "Column O refreshes")
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "O1" -Rows $notesFlowRows)
    $notesHeaderRows = New-Object 'object[][]' 1
    $notesHeaderRows[0] = [object[]]@("ExistingMeetingNotes", "NewPlanningNotes", "NewTimeline", "NewStatus")
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "O4" -Rows $notesHeaderRows)
    $Worksheet.Range("O1:R3").Interior.Color = 14873826
    $Worksheet.Range("O4:R4").Interior.Color = 14281447
    $Worksheet.Range("O1:R1").Font.Bold = $true
    $Worksheet.Range("O4:R4").Font.Bold = $true
    $Worksheet.Range("O5").Formula2 = '=IFERROR(Notes.Existing,"")'
    $Worksheet.Range("O5:O200").Interior.Color = 16448250
    $Worksheet.Range("P5:R200").Interior.Color = 13431551
    [void](Write-Matrix -Worksheet $Worksheet -TopLeft "P5" -Rows @(
        ,@("Review forecast against latest meeting note", "Apr", "Review")
    ))
    [void](Add-ValidationList -Range $Worksheet.Range("R5:R200") -Source (Get-ListValidationSource "statuses"))
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "PlanningReview"
}

function Build-AutomationSetup {
    param([object]$Worksheet)

    $Worksheet.Range("A1:F40").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Automation Setup" `
        -Subtitle "Optional Office Scripts are distributed as source files. Import them into Excel Automate when writeback automation is wanted." `
        -BandRange "A1:F3"

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
    $releaseRows[0] = [object[]]@("Governance_Starter$artifactSuffix.xltx", "Workbook template with sheets, tables, formulas, and Power Query outputs.")
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

function Build-AssetFinanceSetup {
    param([object]$Worksheet)

    $Worksheet.Range("A1:F40").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Asset Finance Setup" `
        -Subtitle "Assumptions for AssetFinance formulas. Finance outputs read tblAssetEvidence_ModelInputs and keep mapped-only evidence out of final model outputs." `
        -BandRange "A1:F3"

    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblAssetFinanceAssumptions" `
        -TopLeft "A4" `
        -Rows (Read-TsvMatrix "samples\asset_finance_assumptions_starter.tsv"))

    [void](Set-NumberFormat -Range $Worksheet.Range("B5:B100") -Format "0")
    [void](Add-NonNegativeValidation -Range $Worksheet.Range("B5:B100"))
    $Worksheet.Range("A:F").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
}

function Build-DataImportSetup {
    param([object]$Worksheet)

    $Worksheet.Range("A1:H100").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Data Import Setup" `
        -Subtitle "Public-safe source profile and canonical budget import contract. Formulas read tblBudgetInput; Planning Table remains the manual starter source." `
        -BandRange "A1:H3"

    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblDataSourceProfile" `
        -TopLeft "A4" `
        -Rows (Read-TsvMatrix "samples\data_source_profile_starter.tsv"))

    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblBudgetImportParameters" `
        -TopLeft "E4" `
        -Rows (Read-TsvMatrix "samples\budget_import_parameters_starter.tsv"))

    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblBudgetImportContract" `
        -TopLeft "A16" `
        -Rows (Read-TsvMatrix "samples\budget_import_contract_starter.tsv"))

    $Worksheet.Range("A:H").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    $Worksheet.Columns.Item(1).ColumnWidth = 24
    $Worksheet.Columns.Item(2).ColumnWidth = 34
    $Worksheet.Columns.Item(3).ColumnWidth = 58
    $Worksheet.Columns.Item(4).ColumnWidth = 64
    $Worksheet.Columns.Item(5).ColumnWidth = 24
    $Worksheet.Columns.Item(6).ColumnWidth = 34
    $Worksheet.Columns.Item(7).ColumnWidth = 58
    $Worksheet.Columns.Item(8).ColumnWidth = 18
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @() -DefaultHeight 20
    $Worksheet.Rows.Item(2).RowHeight = 36
}

function Build-BudgetInput {
    param([object]$Worksheet)

    $Worksheet.Range("A1:BL234").Clear()
    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblBudgetInput" `
        -TopLeft "A1" `
        -Rows (Read-TsvMatrix "samples\planning_table_starter.tsv"))

    [void](Set-NumberFormat -Range $Worksheet.Range("O2:AZ234") -Format "$#,##0")
    [void](Set-NumberFormat -Range $Worksheet.Range("BJ2:BJ234") -Format "0")
    $Worksheet.Range("A:BL").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
}

function Build-BudgetQA {
    param([object]$Worksheet)

    $Worksheet.Range("A1:F80").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Budget Import QA" `
        -Subtitle "Hidden import status and issue tables surfaced through Source Status." `
        -BandRange "A1:F2"
    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblBudgetImportStatus" `
        -TopLeft "A3" `
        -Rows (Read-TsvMatrix "samples\budget_import_status_starter.tsv"))

    $Worksheet.Range("A9").Value2 = "Budget Import Issues"
    $Worksheet.Range("A9").Font.Bold = $true
    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblBudgetImportIssues" `
        -TopLeft "A11" `
        -Rows (Read-TsvMatrix "samples\budget_import_issues_starter.tsv"))

    $Worksheet.Range("A:F").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
}

function Build-IntegrationBridge {
    param([object]$Worksheet)

    $Worksheet.Range("A1:O120").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Integration Bridge" `
        -Subtitle 'Optional reviewed-evidence handoff. ProjectKey is Source ID & "-" & Job ID. Approved evidence is advisory only.' `
        -BandRange "A1:O3"

    Format-SectionHeader `
        -Anchor $Worksheet.Range("A5") `
        -Title "Financial project register export" `
        -Note 'Export this reviewed project identity shape to the evidence review workspace. It does not create official financial projects.'
    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblFinancialProjectRegisterExport" `
        -TopLeft "A7" `
        -Rows (Read-TsvMatrix "samples\financial_project_register_export_starter.tsv"))

    Set-MergedPanel `
        -Worksheet $Worksheet `
        -Address "J7:O13" `
        -Text 'ProjectKey is derived for the bridge as Source ID & "-" & Job ID. Raw file paths are evidence context, not project identity. The bridge exports workbook project identity for review; it does not make new financial projects.' `
        -FillColor 13431551

    Format-SectionHeader `
        -Anchor $Worksheet.Range("A17") `
        -Title "Approved evidence import" `
        -Note "Paste or load approved evidence rows only. These links are advisory context for review, not workbook status updates."
    Set-MergedPanel `
        -Worksheet $Worksheet `
        -Address "A19:O23" `
        -Text "Approved evidence remains separate from generated candidates and manual review decisions. It may support review, but it must not overwrite planning status, create projects, or turn documentation signals into official finance status." `
        -FillColor 16448250

    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblApprovedProjectEvidence" `
        -TopLeft "A26" `
        -Rows (Read-TsvMatrix "samples\approved_project_evidence_starter.tsv"))

    $Worksheet.Range("A:O").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    $Worksheet.Columns.Item(1).ColumnWidth = 18
    $Worksheet.Columns.Item(3).ColumnWidth = 22
    $Worksheet.Columns.Item(4).ColumnWidth = 28
    $Worksheet.Columns.Item(9).ColumnWidth = 18
    $Worksheet.Columns.Item(10).ColumnWidth = 24
    $Worksheet.Columns.Item(11).ColumnWidth = 18
    $Worksheet.Columns.Item(12).ColumnWidth = 18
    $Worksheet.Columns.Item(14).ColumnWidth = 30
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @(5, 17) -DefaultHeight 20
    $Worksheet.Rows.Item(2).RowHeight = 36
}

function Build-WorkbookManifest {
    param([object]$Worksheet)

    $Worksheet.Range("A1:L220").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Workbook Manifest" `
        -Subtitle "Source-controlled sheet/table map used for workbook navigation and visibility." `
        -BandRange "A1:L3"
    [void](Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblWorkbookManifest" `
        -TopLeft "A4" `
        -Rows (Read-TsvMatrix "samples\workbook_manifest.tsv"))
    $Worksheet.Range("A:L").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    $Worksheet.Columns.Item(6).ColumnWidth = 18
    $Worksheet.Columns.Item(7).ColumnWidth = 20
    $Worksheet.Columns.Item(8).ColumnWidth = 24
    $Worksheet.Columns.Item(12).ColumnWidth = 52
}

function Build-StartHere {
    param([object]$Worksheet)

    $Worksheet.Range("A1:M80").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Start Here" `
        -Subtitle "Use this workbook left to right: check source health, configure imports, refresh the canonical input, review outputs, then use hidden backend sheets only when troubleshooting." `
        -BandRange "A1:M3"

    Format-SectionHeader `
        -Anchor $Worksheet.Range("A5") `
        -Title "Workbook flow" `
        -Note "The workbook keeps data intake, canonical source rows, formula outputs, review hubs, and optional staged writeback as separate surfaces."
    $flowRows = New-Object 'object[][]' 8
    $flowRows[0] = [object[]]@("Step", "From", "To", "Purpose")
    $flowRows[1] = [object[]]@("1", "Manual workbook source / optional placeholder adapter", "Power Query or current-workbook adapter", "Select and shape incoming planning data.")
    $flowRows[2] = [object[]]@("2", "Power Query or current-workbook adapter", "tblBudgetInput", "Load the canonical 64-column planning contract.")
    $flowRows[3] = [object[]]@("3", "tblBudgetInput", "Governed formula modules", "Keep formulas independent of the source system.")
    $flowRows[4] = [object[]]@("4", "Governed formula modules", "Planning Review / Analysis Hub / Asset Hub / Asset Finance Hub", "Review controlled outputs on a smaller visible sheet set.")
    $flowRows[5] = [object[]]@("5", "Planning Review P:R", "Decision Staging", "Prepare optional notes/status/timeline writeback.")
    $flowRows[6] = [object[]]@("6", "Decision Staging", "Planning Table", "Apply reviewed writeback, then refresh or re-sync before relying on outputs.")
    $flowRows[7] = [object[]]@("7", "tblBudgetInput and approved evidence rows", "Integration Bridge", "Exchange reviewed evidence mappings as advisory context only.")
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblStartHereFlow" -TopLeft "A7" -Rows $flowRows)

    Format-SectionHeader `
        -Anchor $Worksheet.Range("F5") `
        -Title "Key rule" `
        -Note "tblBudgetInput is the canonical formula source. Planning Table is manual/staging/local writeback. After Planning Table edits or ApplyNotes, refresh or re-sync before relying on formula outputs."
    Set-MergedPanel `
        -Worksheet $Worksheet `
        -Address "F7:M13" `
        -Text "Key source boundary: tblBudgetInput is the canonical formula source. Planning Table is a manual/staging/local writeback surface. If Planning Table changes manually or through ApplyNotes, refresh or re-sync the current-workbook adapter before relying on Planning Review, Analysis Hub, Asset Hub, or Asset Finance Hub outputs. Assets are optional. In AssetsLite, start with Asset Hub, then enter simple assets in Asset Register. In Planning, ignore asset sheets. Asset Evidence, Asset Finance, and Semantic Map are advanced paths." `
        -FillColor 13431551

    Format-SectionHeader `
        -Anchor $Worksheet.Range("A16") `
        -Title "Go to" `
        -Note "Use these visible sheets for the normal left-to-right operator flow."
    $navRows = New-Object 'object[][]' 10
    $navRows[0] = [object[]]@("Sheet", "Use it for", "Normal action")
    $navRows[1] = [object[]]@("Source Status", "Check freshness and import issues.", "Review first.")
    $navRows[2] = [object[]]@("Data Import Setup", "Configure source mode and schema.", "Update source profile and contract.")
    $navRows[3] = [object[]]@("Integration Bridge", "Reviewed evidence handoff.", "Export project keys or paste approved evidence.")
    $navRows[4] = [object[]]@("Planning Table", "Manual starter/local writeback.", "Edit only when using manual/current-workbook mode.")
    $navRows[5] = [object[]]@("Cap Setup", "BU cap limits.", "Review or update caps.")
    $navRows[6] = [object[]]@("Planning Review", "Main planning report.", "Run meeting review and enter P:R notes.")
    $navRows[7] = [object[]]@("Analysis Hub", "Planning analysis outputs.", "Review scorecards, queues, burndown, and readiness.")
    $navRows[8] = [object[]]@("Asset Hub", "Optional asset workflow guide.", "AssetsLite users start here, then enter simple assets in Asset Register.")
    $navRows[9] = [object[]]@("Asset Finance Hub", "Advanced asset finance outputs.", "Use only after classified asset evidence exists.")
    $navTable = Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblStartHereNavigation" -TopLeft "A18" -Rows $navRows
    for ($index = 1; $index -le 9; $index++) {
        $sheetName = [string]$navRows[$index][0]
        Set-InternalWorksheetLink -Range $navTable.DataBodyRange.Cells.Item($index, 1) -SheetName $sheetName -DisplayText $sheetName
    }

    Format-SectionHeader `
        -Anchor $Worksheet.Range("F16") `
        -Title "Backend/admin sheets" `
        -Note "Hidden sheets keep the governed backend available without putting every implementation table in the normal operator path."
    Set-MergedPanel `
        -Worksheet $Worksheet `
        -Address "F18:M24" `
        -Text "Hidden backend/admin sheets are still part of the governed workbook. PQ Budget Input, PQ Budget QA, Validation Lists, Decision Staging, Automation Setup, asset setup tables, Workbook Manifest, and intermediate asset-evidence outputs stay hidden by default. Integration Bridge is visible because it is an operator handoff surface. Unhide backend sheets only for troubleshooting, administration, or release checks." `
        -FillColor 16448250
    $Worksheet.Range("A:M").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    $Worksheet.Columns.Item(3).ColumnWidth = 42
    $Worksheet.Columns.Item(4).ColumnWidth = 52
    $Worksheet.Columns.Item(6).ColumnWidth = 16
    $Worksheet.Columns.Item(7).ColumnWidth = 18
    $Worksheet.Columns.Item(8).ColumnWidth = 18
    $Worksheet.Columns.Item(9).ColumnWidth = 18
    $Worksheet.Columns.Item(10).ColumnWidth = 18
    $Worksheet.Columns.Item(11).ColumnWidth = 18
    $Worksheet.Columns.Item(12).ColumnWidth = 18
    $Worksheet.Columns.Item(13).ColumnWidth = 18
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @(5, 16) -DefaultHeight 20
    $Worksheet.Rows.Item(7).RowHeight = 24
    $Worksheet.Rows.Item(8).RowHeight = 24
    $Worksheet.Rows.Item(9).RowHeight = 24
    $Worksheet.Rows.Item(10).RowHeight = 24
    $Worksheet.Rows.Item(11).RowHeight = 24
    $Worksheet.Rows.Item(12).RowHeight = 24
    $Worksheet.Rows.Item(13).RowHeight = 24
    foreach ($row in 18..24) {
        $Worksheet.Rows.Item($row).RowHeight = 22
    }
}

function Build-SourceStatus {
    param([object]$Worksheet)

    Format-HubSheet `
        -Worksheet $Worksheet `
        -Title "Source Status" `
        -Note "Canonical budget import status and source trust checks." `
        -ClearRange "A1:Z80"
    [void](Set-RangeFormula -Range $Worksheet.Range("A4") -Formula "=Source.SOURCE_STATUS()")
    $Worksheet.Range("A:Z").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "SourceStatus"
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @() -DefaultHeight 20
}

function Build-AnalysisHub {
    param([object]$Worksheet)

    Format-HubSheet `
        -Worksheet $Worksheet `
        -Title "Analysis Hub" `
        -Note "Planning outputs grouped into one review surface instead of separate demo sheets." `
        -ClearRange "A1:Z360"
    $sections = @(
        @{ Cell = "A14"; Title = "BU Cap Scorecard"; Note = "Cap and spend posture by BU."; Formula = "=Analysis.BU_CAP_SCORECARD()" },
        @{ Cell = "A52"; Title = "Reforecast Queue"; Note = "Grouped action queue for forecast review."; Formula = "=Analysis.REFORECAST_QUEUE()" },
        @{ Cell = "A114"; Title = "PM Spend Report"; Note = "Existing-work summary and job detail."; Formula = "=Analysis.PM_SPEND_REPORT()" },
        @{ Cell = "A176"; Title = "Working Budget Screen"; Note = "Current-job screening before budget drafting."; Formula = "=Analysis.WORKING_BUDGET_SCREEN()" },
        @{ Cell = "A238"; Title = "Burndown Screen"; Note = "Meeting view of remaining burn and drivers."; Formula = "=Analysis.BURNDOWN_SCREEN()" },
        @{ Cell = "A300"; Title = "Internal Jobs Export"; Note = "Header-driven internal work export for readiness smoke testing."; Formula = "=Ready.InternalJobs_Export()" }
    )
    $toc = Add-HubTableOfContents -Worksheet $Worksheet -TableName "tblAnalysisHubSections" -Sections $sections
    foreach ($section in $sections) {
        Add-HubSection -Worksheet $Worksheet -Cell $section.Cell -Title $section.Title -Note $section.Note -Formula $section.Formula
    }
    $Worksheet.Range("A:Z").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "Analysis"
    Set-HubTableOfContentsColumnWidths -Table $toc
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @(14, 52, 114, 176, 238, 300) -DefaultHeight 20
}

function Build-AssetHub {
    param([object]$Worksheet)

    Format-HubSheet `
        -Worksheet $Worksheet `
        -Title "Asset Hub" `
        -Note "Optional workflow for entering simple assets, then connecting projects to assets only when that extra tracking is in scope." `
        -ClearRange "A1:Z460"

    $sections = @(
        @{ Cell = "D4"; Title = "Asset workflow mode"; Note = "Choose whether optional asset tracking is in scope." },
        @{ Cell = "A18"; Title = "Simple asset entry"; Note = "The first asset task is just entering an asset in Asset Register." },
        @{ Cell = "A30"; Title = "Minimum field guide"; Note = "Minimum and optional fields for tblAssets." },
        @{ Cell = "A50"; Title = "Asset register start here"; Note = "Short path for entering one simple asset safely." },
        @{ Cell = "A70"; Title = "Asset register status"; Note = "Counts for entered register rows and required-field readiness." },
        @{ Cell = "A94"; Title = "Asset register issues"; Note = "Register-level issues and advisory LinkedProjectID checks." },
        @{ Cell = "A128"; Title = "What should I do next?"; Note = "Mode-aware next action after simple asset entry." },
        @{ Cell = "A150"; Title = "Review queues"; Note = "Friendly overview of asset queues before technical issue sections." },
        @{ Cell = "A176"; Title = "Asset Mapping Issues"; Note = "Project-to-asset mapping issues for review." },
        @{ Cell = "A216"; Title = "Project Promotion Queue"; Note = "Candidate asset promotion rows." },
        @{ Cell = "A256"; Title = "Asset Change Issues"; Note = "Change staging issues before controlled apply." },
        @{ Cell = "A296"; Title = "Installed Without Evidence"; Note = "Installed-state assets missing evidence links." },
        @{ Cell = "A336"; Title = "Replacement Source/Target Issues"; Note = "Replacement rows missing required source or target asset context." },
        @{ Cell = "A386"; Title = "Asset terms"; Note = "Plain-language glossary for asset workflow terminology." },
        @{ Cell = "A416"; Title = "Admin / troubleshooting table map"; Note = "Hidden asset tables for intentional administration and troubleshooting." }
    )
    $toc = Add-HubTableOfContents -Worksheet $Worksheet -TableName "tblAssetHubSections" -Sections $sections -TopLeft "A4"

    Format-SectionHeader `
        -Anchor $Worksheet.Range("D4") `
        -Title "Asset workflow mode" `
        -Note "Assets are optional. Leave the mode Off when the workbook is only being used for capital planning."
    $settingsTable = Add-TableFromMatrix `
        -Worksheet $Worksheet `
        -TableName "tblAssetWorkflowSettings" `
        -TopLeft "D6" `
        -Rows (Read-TsvMatrix "samples\asset_workflow_settings_starter.tsv")
    [void](Add-ValidationList -Range $settingsTable.DataBodyRange.Cells.Item(1, 2) -Source (Get-ListValidationSource "assetWorkflowModes"))

    Format-SectionHeader `
        -Anchor $Worksheet.Range("A18") `
        -Title "Simple asset entry" `
        -Note "To enter one asset, go to Asset Register. Asset Evidence, Asset Finance, Semantic Map, Asset State History, and PQ Asset Evidence are not needed for this."
    $Worksheet.Range("A21").Value2 = "To enter one asset, go to Asset Register."
    $Worksheet.Range("A21").Font.Bold = $true
    $Worksheet.Range("A22").Value2 = "Minimum fields: AssetID, AssetName, AssetType, Status."
    $Worksheet.Range("A23").Value2 = "Helpful optional fields: Site, Location, Owner, Condition, Criticality, ReplacementCost, UsefulLifeYears, LinkedProjectID."
    $Worksheet.Range("A24").Value2 = "LinkedProjectID is optional and advisory. Manual IDs are allowed and do not imply external refresh or sync."
    Set-InternalWorksheetLink -Range $Worksheet.Range("D21") -SheetName "Asset Register" -DisplayText "Open Asset Register"
    $Worksheet.Range("A21:H24").WrapText = $true

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A30" `
        -Title "Minimum field guide" `
        -Note "Native Excel validation supports the fields below; LinkedProjectID is advisory and allows blanks or manual IDs." `
        -Formula "=Assets.ASSET_REGISTER_FIELD_GUIDE"

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A50" `
        -Title "Asset register start here" `
        -Note "Enter a simple asset first. Mapping, evidence, finance, and semantic layers are later paths." `
        -Formula "=Assets.ASSET_REGISTER_START_HERE"

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A70" `
        -Title "Asset register status" `
        -Note "Shows whether tblAssets has entered rows and whether required fields are ready." `
        -Formula "=Assets.ASSET_REGISTER_STATUS"

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A94" `
        -Title "Asset register issues" `
        -Note "Flags missing IDs, duplicate IDs, missing names/statuses, negative values, and advisory project links." `
        -Formula "=Assets.ASSET_REGISTER_ISSUES"

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A128" `
        -Title "What should I do next?" `
        -Note "The next action responds to the selected asset workflow mode and whether asset data is present." `
        -Formula "=Assets.ASSET_NEXT_ACTIONS"

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A150" `
        -Title "Review queues" `
        -Note "Friendly map of the technical queues below. Use these only after choosing an asset path." `
        -Formula "=Assets.ASSET_REVIEW_QUEUE"

    $technicalSections = @(
        @{ Cell = "A176"; Title = "Asset Mapping Issues"; Note = "Project-to-asset mapping issues for review."; Formula = "=Assets.ASSET_MAPPING_ISSUES" },
        @{ Cell = "A216"; Title = "Project Promotion Queue"; Note = "Candidate asset promotion rows."; Formula = "=Assets.PROJECT_PROMOTION_QUEUE" },
        @{ Cell = "A256"; Title = "Asset Change Issues"; Note = "Change staging issues before controlled apply."; Formula = "=Assets.ASSET_CHANGE_ISSUES" },
        @{ Cell = "A296"; Title = "Installed Without Evidence"; Note = "Installed-state assets missing evidence links."; Formula = "=Assets.INSTALLED_WITHOUT_EVIDENCE" },
        @{ Cell = "A336"; Title = "Replacement Source/Target Issues"; Note = "Replacement rows missing required source or target asset context."; Formula = "=Assets.REPLACEMENT_SOURCE_TARGET_ISSUES" }
    )
    foreach ($section in $technicalSections) {
        Add-HubSection -Worksheet $Worksheet -Cell $section.Cell -Title $section.Title -Note $section.Note -Formula $section.Formula
    }

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A386" `
        -Title "Asset terms" `
        -Note "Plain-language glossary for asset workflow terminology." `
        -Formula "=Assets.ASSET_GLOSSARY"

    Add-HubSection `
        -Worksheet $Worksheet `
        -Cell "A416" `
        -Title "Admin / troubleshooting table map" `
        -Note "Hidden asset tables are still available for operators who intentionally enable the asset workflow." `
        -Formula "=Assets.ASSET_TABLE_MAP"

    $Worksheet.Range("A:Z").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "Asset"
    Set-HubTableOfContentsColumnWidths -Table $toc
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @(4, 18, 30, 50, 70, 94, 128, 150, 176, 216, 256, 296, 336, 386, 416) -DefaultHeight 20
}

function Build-AssetFinanceHub {
    param([object]$Worksheet)

    Format-HubSheet `
        -Worksheet $Worksheet `
        -Title "Asset Finance Hub" `
        -Note "Optional finance outputs. Use this only when classified asset evidence exists and is ready for AssetFinance." `
        -ClearRange "A1:Z300"

    $sections = @(
        @{ Cell = "A12"; Title = "Finance gate"; Note = "Asset finance reads tblAssetEvidence_ModelInputs, not raw or mapped-only asset rows."; Formula = "=AssetFinance.FINANCE_START_HERE" },
        @{ Cell = "A26"; Title = "Readiness status"; Note = "Classified evidence count and unsupported-assumption counts before reviewing finance outputs."; Formula = "=AssetFinance.FINANCE_READINESS_STATUS" },
        @{ Cell = "A50"; Title = "Asset Depreciation"; Note = "Classified asset evidence converted to depreciation-ready rows."; Formula = "=AssetFinance.DEPRECIATION_SCHEDULE" },
        @{ Cell = "A124"; Title = "Asset Funding Requirements"; Note = "Classified asset evidence grouped into funding requirements."; Formula = "=AssetFinance.FUNDING_REQUIREMENTS" },
        @{ Cell = "A184"; Title = "Asset Finance Totals"; Note = "Asset finance summary totals from classified model inputs."; Formula = "=AssetFinance.FINANCE_TOTALS" },
        @{ Cell = "A214"; Title = "Asset Finance Charts"; Note = "Chart-ready asset finance feeds; no native chart objects yet."; Formula = "=AssetFinance.CHART_FEEDS" }
    )
    $toc = Add-HubTableOfContents -Worksheet $Worksheet -TableName "tblAssetFinanceHubSections" -Sections $sections -TopLeft "A4"
    foreach ($section in $sections) {
        Add-HubSection -Worksheet $Worksheet -Cell $section.Cell -Title $section.Title -Note $section.Note -Formula $section.Formula
    }
    $Worksheet.Range("A:Z").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "AssetFinance"
    Set-HubTableOfContentsColumnWidths -Table $toc
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @(12, 26, 50, 124, 184, 214) -DefaultHeight 20
}

function Build-SemanticMapSetup {
    param([object]$Worksheet)

    $Worksheet.Range("A1:K90").Clear()
    Format-PageHeader `
        -Worksheet $Worksheet `
        -Title "Semantic Map Setup" `
        -Subtitle "Hidden reference crosswalk tables for private workbook extension. Do not add private endpoints or full ontology dumps." `
        -BandRange "A1:K3"

    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblOntologyNamespaces" -TopLeft "A4" -Rows (Read-TsvMatrix "samples\ontology_namespaces_starter.tsv"))
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblOntologyClassMap" -TopLeft "A10" -Rows (Read-TsvMatrix "samples\ontology_class_map_starter.tsv"))
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblOntologyRelationshipMap" -TopLeft "A28" -Rows (Read-TsvMatrix "samples\ontology_relationship_map_starter.tsv"))
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblProjectSemanticMap" -TopLeft "A40" -Rows (Read-TsvMatrix "samples\project_semantic_map_starter.tsv"))
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblAssetSemanticMap" -TopLeft "A45" -Rows (Read-TsvMatrix "samples\asset_semantic_map_starter.tsv"))
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblOntologyExportQueue" -TopLeft "A50" -Rows (Read-TsvMatrix "samples\ontology_export_queue_starter.tsv"))
    [void](Add-TableFromMatrix -Worksheet $Worksheet -TableName "tblOntologyIssues" -TopLeft "A55" -Rows (Read-TsvMatrix "samples\ontology_issues_starter.tsv"))

    $Worksheet.Range("A:K").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "Semantic"
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -DefaultHeight 20
}

function Build-SemanticMapHub {
    param([object]$Worksheet)

    Format-HubSheet `
        -Worksheet $Worksheet `
        -Title "Semantic Map Hub" `
        -Note "Reference-only semantic crosswalk. Keep it separate from the normal planning and asset operator workflow." `
        -ClearRange "A1:Z270"

    $sections = @(
        @{ Cell = "A15"; Title = "Start here"; Note = "Plain-language guidance for the reference crosswalk."; Formula = "=Ontology.ONTOLOGY_START_HERE" },
        @{ Cell = "A35"; Title = "Semantic mapping status"; Note = "Counts for project mappings, asset mappings, export rows, and ontology issues."; Formula = "=Ontology.SEMANTIC_MAPPING_STATUS" },
        @{ Cell = "A55"; Title = "Ontology issues"; Note = "Rows missing identifiers, classes, predicates, or objects."; Formula = "=Ontology.ONTOLOGY_ISSUES" },
        @{ Cell = "A100"; Title = "Triple export queue"; Note = "Reviewable Subject-Predicate-Object rows for private extension."; Formula = "=Ontology.TRIPLE_EXPORT_QUEUE" },
        @{ Cell = "A150"; Title = "Class map"; Note = "Curated class labels. This is intentionally not a full ontology dump."; Formula = "=Ontology.CLASS_MAP" },
        @{ Cell = "A190"; Title = "Relationship map"; Note = "Curated relationships for location, composition, feeds, points, and project impact."; Formula = "=Ontology.RELATIONSHIP_MAP" },
        @{ Cell = "A225"; Title = "Reference export note"; Note = "Guidance only. This slice does not implement a deployed graph export."; Formula = "=Ontology.JSONLD_EXPORT_HELP" }
    )
    $toc = Add-HubTableOfContents -Worksheet $Worksheet -TableName "tblSemanticMapHubSections" -Sections $sections -TopLeft "A4"
    foreach ($section in $sections) {
        Add-HubSection -Worksheet $Worksheet -Cell $section.Cell -Title $section.Title -Note $section.Note -Formula $section.Formula
    }

    $Worksheet.Range("A:Z").WrapText = $true
    [void]$Worksheet.Columns.AutoFit()
    Apply-HubColumnWidthTemplate -Worksheet $Worksheet -Template "Semantic"
    Set-HubTableOfContentsColumnWidths -Table $toc
    Normalize-GeneratedSheetRows -Worksheet $Worksheet -SectionRows @(15, 35, 55, 100, 150, 190, 225) -DefaultHeight 20
}

function Apply-WorkbookManifestVisibility {
    param(
        [object]$Workbook,
        [string]$Edition = "Planning"
    )

    $visibleValue = -1
    $hiddenValue = 0
    $rows = Read-TsvMatrix "samples\workbook_manifest.tsv"
    $headers = @{}
    for ($columnIndex = 0; $columnIndex -lt $rows[0].Count; $columnIndex++) {
        $headers[[string]$rows[0][$columnIndex]] = $columnIndex
    }
    $visibleSheetOrder = @()
    $startSheet = Get-WorksheetOrNull -Workbook $Workbook -Name "Start Here"
    if ($null -ne $startSheet) {
        $startSheet.Visible = $visibleValue
        [void]$startSheet.Activate()
    }

    for ($index = 1; $index -lt $rows.Count; $index++) {
        $sheetName = [string]$rows[$index][$headers["SheetName"]]
        $visibility = ([string]$rows[$index][$headers["Visibility"]]).Trim().ToLowerInvariant()
        $editionList = if ($headers.ContainsKey("Edition")) { [string]$rows[$index][$headers["Edition"]] } else { "Planning;AssetsLite;AssetsFull" }
        $editionTokens = @($editionList -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        $includedInEdition = $editionTokens.Count -eq 0 -or $editionTokens -contains "All" -or $editionTokens -contains $Edition
        if ([string]::IsNullOrWhiteSpace($sheetName)) { continue }

        $sheet = Get-WorksheetOrNull -Workbook $Workbook -Name $sheetName
        if ($null -eq $sheet) { continue }

        if ($visibility -eq "visible" -and $includedInEdition) {
            $sheet.Visible = $visibleValue
            $visibleSheetOrder += $sheetName
        } elseif ($visibility -eq "hidden" -or -not $includedInEdition) {
            $sheet.Visible = $hiddenValue
        } else {
            throw "Unsupported workbook manifest visibility '$visibility' for sheet '$sheetName'."
        }
    }

    for ($index = $visibleSheetOrder.Count - 1; $index -ge 0; $index--) {
        $sheet = Get-WorksheetOrNull -Workbook $Workbook -Name $visibleSheetOrder[$index]
        if ($null -ne $sheet) {
            [void]$sheet.Move($Workbook.Worksheets.Item(1))
        }
    }
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
    Format-PageHeader `
        -Worksheet $assetRegisterSheet `
        -Title "Asset Register" `
        -Subtitle "Start with Asset Register to enter a simple asset. Minimum fields are AssetID, AssetName, AssetType, and Status." `
        -BandRange "A1:Q3"
    $assetRegisterSheet.Range("A4").Value2 = "You do not need Asset Evidence, Asset Finance, Semantic Map, Asset State History, or PQ Asset Evidence to enter a simple asset."
    $assetRegisterSheet.Range("A4:Q4").Interior.Color = 13431551
    $assetRegisterSheet.Range("A4:Q4").WrapText = $true
    $assetRegister = Add-TableFromMatrix -Worksheet $assetRegisterSheet -TableName "tblAssets" -TopLeft "A6" -Rows (Read-TsvMatrix "samples\asset_register_starter.tsv")
    [void](Apply-AssetRegisterValidation -Table $assetRegister)

    $semanticSheet = Add-Worksheet -Workbook $Workbook -Name "Semantic Assets"
    [void](Add-TableFromMatrix -Worksheet $semanticSheet -TableName "tblSemanticAssets" -TopLeft "A1" -Rows (Read-TsvMatrix "samples\semantic_assets_starter.tsv"))

    $assetSetupSheet = Add-Worksheet -Workbook $Workbook -Name "Asset Setup"
    $assetSetupRows = Read-TsvMatrix "samples\asset_setup_starter.tsv"
    if ($assetSetupRows.Count -ge 4) {
        $promotionRows = @($assetSetupRows[0], $assetSetupRows[1])
        $mappingRows = @($assetSetupRows[2], $assetSetupRows[3])
    } else {
        $promotionRows = New-Object 'object[][]' 2
        $promotionRows[0] = $assetSetupRows[0]
        $promotionRows[1] = New-BlankRowLike -HeaderRow $assetSetupRows[0]
        $mappingRows = New-Object 'object[][]' 2
        $mappingRows[0] = $assetSetupRows[1]
        $mappingRows[1] = New-BlankRowLike -HeaderRow $assetSetupRows[1]
    }
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
        [void](Apply-TableListValidation -Table $table -Header "ApplyReady" -ListKey "booleanFlags")
    }

    [void](Set-FreezeRows -Excel $Excel -Worksheet $assetRegisterSheet -Rows 6)
    [void]$assetRegisterSheet.Columns.AutoFit()
    $assetRegisterSheet.Columns.Item(1).ColumnWidth = 18
    $assetRegisterSheet.Columns.Item(2).ColumnWidth = 28
    $assetRegisterSheet.Columns.Item(3).ColumnWidth = 18
    $assetRegisterSheet.Columns.Item(8).ColumnWidth = 16
    $assetRegisterSheet.Columns.Item(12).ColumnWidth = 16
    $assetRegisterSheet.Columns.Item(13).ColumnWidth = 16
    $assetRegisterSheet.Columns.Item(16).ColumnWidth = 22

    foreach ($sheet in @($semanticSheet, $assetSetupSheet, $projectMapSheet, $assetChangesSheet, $assetHistorySheet)) {
        [void](Set-FreezeRows -Excel $Excel -Worksheet $sheet -Rows 1)
        [void]$sheet.Columns.AutoFit()
    }
}

function Build-DemoOutputs {
    param([object]$Workbook)

    $reviewSheet = Add-Worksheet -Workbook $Workbook -Name "Planning Review"
    [void](Set-RangeFormula -Range $reviewSheet.Range("A4") -Formula "=CapitalPlanning.CAPITAL_PLANNING_REPORT()")

    Build-StartHere -Worksheet (Add-Worksheet -Workbook $Workbook -Name "Start Here")
    Build-SourceStatus -Worksheet (Add-Worksheet -Workbook $Workbook -Name "Source Status")
    Build-AnalysisHub -Worksheet (Add-Worksheet -Workbook $Workbook -Name "Analysis Hub")
    Build-AssetHub -Worksheet (Add-Worksheet -Workbook $Workbook -Name "Asset Hub")
    Build-AssetFinanceHub -Worksheet (Add-Worksheet -Workbook $Workbook -Name "Asset Finance Hub")
    Build-SemanticMapHub -Worksheet (Add-Worksheet -Workbook $Workbook -Name "Semantic Map Hub")
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

    $startSheet = $workbook.Worksheets.Item(1)
    $startSheet.Name = "Start Here"
    [void](Build-StartHere -Worksheet $startSheet)

    $planningSheet = Add-Worksheet -Workbook $workbook -Name "Planning Table"
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

    $dataImportSheet = Add-Worksheet -Workbook $workbook -Name "Data Import Setup"
    [void](Build-DataImportSetup -Worksheet $dataImportSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $dataImportSheet -Rows 16)

    $budgetInputSheet = Add-Worksheet -Workbook $workbook -Name "PQ Budget Input"
    [void](Build-BudgetInput -Worksheet $budgetInputSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $budgetInputSheet -Rows 1)

    $budgetQaSheet = Add-Worksheet -Workbook $workbook -Name "PQ Budget QA"
    [void](Build-BudgetQA -Worksheet $budgetQaSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $budgetQaSheet -Rows 3)

    $integrationBridgeSheet = Add-Worksheet -Workbook $workbook -Name "Integration Bridge"
    [void](Build-IntegrationBridge -Worksheet $integrationBridgeSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $integrationBridgeSheet -Rows 7)

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

    $assetFinanceSheet = Add-Worksheet -Workbook $workbook -Name "Asset Finance Setup"
    [void](Build-AssetFinanceSetup -Worksheet $assetFinanceSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $assetFinanceSheet -Rows 4)

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
    Apply-AssetRegisterRelationshipValidation -Workbook $workbook | Out-Null

    $semanticSetupSheet = Add-Worksheet -Workbook $workbook -Name "Semantic Map Setup"
    [void](Build-SemanticMapSetup -Worksheet $semanticSetupSheet)
    [void](Set-FreezeRows -Excel $excel -Worksheet $semanticSetupSheet -Rows 3)

    $manifestSheet = Add-Worksheet -Workbook $workbook -Name "Workbook Manifest"
    [void](Build-WorkbookManifest -Worksheet $manifestSheet)

    Set-PublicWorkbookProperties -Workbook $workbook
    [void]$startSheet.Activate()
    [void]$workbook.SaveAs($coreWorkbookPath, 51)
    Write-Host "Built core governance workbook: $coreWorkbookPath"
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
    $installedNames = Install-FormulaModules -Workbook $workbook
    Build-DemoOutputs -Workbook $workbook | Out-Null
    [void]$workbook.Worksheets.Item("Start Here").Activate()
    Apply-WorkbookManifestVisibility -Workbook $workbook -Edition $Edition
    [void]$workbook.Worksheets.Item("Start Here").Activate()
    Set-PublicWorkbookProperties -Workbook $workbook
    [void]$excel.CalculateFull()
    Assert-NoVisibleWorkbookErrors -Workbook $workbook
    [void]$workbook.Save()
    [void]$workbook.SaveAs($templateWorkbookPath, 54)
    Write-Host "Built governance starter workbook with $installedNames defined names: $starterWorkbookPath"
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

foreach ($releasePath in @($starterWorkbookPath, $templateWorkbookPath)) {
    Sanitize-WorkbookPackage -Path $releasePath
    Assert-WorkbookPackagePublic -Path $releasePath
}

if (Test-Path -LiteralPath $scratchDirectory) {
    Remove-Item -LiteralPath $scratchDirectory -Recurse -Force
}
