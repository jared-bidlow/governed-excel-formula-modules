[CmdletBinding()]
param(
    [string]$OutputPath = "release_artifacts\asset-evidence-pq\Asset_Evidence_PQ_Seed.xlsx"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$templateDir = Join-Path $repoRoot "samples\power-query\asset-evidence"
$artifactRoot = [System.IO.Path]::GetFullPath((Join-Path $repoRoot "release_artifacts"))

if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    $resolvedOutputPath = [System.IO.Path]::GetFullPath($OutputPath)
} else {
    $resolvedOutputPath = [System.IO.Path]::GetFullPath((Join-Path $repoRoot $OutputPath))
}

if (-not $resolvedOutputPath.StartsWith($artifactRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "OutputPath must stay under release_artifacts: $resolvedOutputPath"
}

$expectedQueries = @(
    @{ Name = "qAssetEvidence_Normalized"; Sheet = "PQ Asset Evidence Normalized"; Table = "tblAssetEvidence_Normalized" },
    @{ Name = "qAssetEvidence_Classified"; Sheet = "PQ Asset Evidence Classified"; Table = "tblAssetEvidence_Classified" },
    @{ Name = "qAssetEvidence_Linked"; Sheet = "PQ Asset Evidence Linked"; Table = "tblAssetEvidence_Linked" },
    @{ Name = "qAssetEvidence_Status"; Sheet = "PQ Asset Evidence Status"; Table = "tblAssetEvidence_Status" },
    @{ Name = "qAssetEvidence_ModelInputs"; Sheet = "PQ Asset Evidence Model Inputs"; Table = "tblAssetEvidence_ModelInputs" },
    @{ Name = "qQA_AssetEvidence_MappingQueue"; Sheet = "PQ Asset Evidence Mapping Queue"; Table = "tblQA_AssetEvidence_MappingQueue" }
)

$templates = foreach ($query in $expectedQueries) {
    $sourceFile = "$($query.Name).m"
    $path = Join-Path $templateDir $sourceFile
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Missing Power Query template: $path"
    }

    [pscustomobject]@{
        QueryName = $query.Name
        SheetName = $query.Sheet
        TableName = $query.Table
        SourceFile = $sourceFile
        FormulaText = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    }
}

$outputDir = Split-Path -Parent $resolvedOutputPath
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

if (Test-Path -LiteralPath $resolvedOutputPath) {
    Remove-Item -LiteralPath $resolvedOutputPath -Force
}

function Add-Worksheet {
    param(
        [object]$Workbook,
        [string]$Name
    )

    $sheet = $Workbook.Worksheets.Add([System.Type]::Missing, $Workbook.Worksheets.Item($Workbook.Worksheets.Count))
    $sheet.Name = $Name
    return $sheet
}

function Add-Table {
    param(
        [object]$Worksheet,
        [string]$Name,
        [string]$TopLeft,
        [object[]]$Headers,
        [object[]]$Rows
    )

    $rowCount = $Rows.Count + 1
    $columnCount = $Headers.Count
    $range = $Worksheet.Range($TopLeft).Resize($rowCount, $columnCount)

    $data = New-Object 'object[,]' $rowCount, $columnCount
    for ($column = 1; $column -le $columnCount; $column++) {
        $columnOffset = $column - 1
        $data[0, $columnOffset] = $Headers[$columnOffset]
    }

    for ($row = 0; $row -lt $Rows.Count; $row++) {
        $values = $Rows[$row]
        for ($column = 1; $column -le $columnCount; $column++) {
            $rowOffset = $row + 1
            $columnOffset = $column - 1
            $data[$rowOffset, $columnOffset] = $values[$columnOffset]
        }
    }

    $range.Value2 = $data

    $table = $Worksheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $Name
    $table.TableStyle = "TableStyleMedium2"
    Format-TableHeader -Table $table
    return $table
}

function Format-TableHeader {
    param([object]$Table)

    try {
        $Table.HeaderRowRange.Font.Bold = $true
        $Table.HeaderRowRange.Font.Color = 0
        $Table.HeaderRowRange.Interior.Color = 16247773
    } catch {
        Write-Warning "Skipped table header formatting for $($Table.Name): $($_.Exception.Message)"
    }
}

function Add-LoadedQueryTable {
    param(
        [object]$Worksheet,
        [string]$QueryName,
        [string]$TableName
    )

    $connectionText = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=`$Workbook`$;Location=$QueryName;Extended Properties=`"`""
    $table = $Worksheet.ListObjects.Add(0, $connectionText, $true, 1, $Worksheet.Range("A1"))
    $table.Name = $TableName
    $table.QueryTable.CommandType = 2
    $table.QueryTable.CommandText = "SELECT * FROM [$QueryName]"
    [void]$table.QueryTable.Refresh($false)
    Format-TableHeader -Table $table
    [void]$Worksheet.Columns.AutoFit()
    return $table
}

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -gt 1) {
        $workbook.Worksheets.Item($workbook.Worksheets.Count).Delete()
    }

    $setupSheet = $workbook.Worksheets.Item(1)
    $setupSheet.Name = "Asset Evidence Setup"
    $setupSheet.Range("A1").Value2 = "Asset Evidence Source"
    $setupSheet.Range("A6").Value2 = "Asset Evidence Rules"
    $setupSheet.Range("A11").Value2 = "Asset Evidence Overrides"
    $setupSheet.Range("A1,A6,A11").Font.Bold = $true

    Add-Table `
        -Worksheet $setupSheet `
        -Name "tblAssetEvidenceSource" `
        -TopLeft "A2" `
        -Headers @("EvidenceId", "SourceSystem", "SourceRecordType", "SourceRecordKey", "ProjectKey", "AssetId", "AssetLabel", "AssetType", "EvidenceDate", "Amount", "FundingSource", "DepreciationClass", "ContextCategoryId", "ContextCategoryName", "Description") `
        -Rows @(
            ,@("EV-001", "Sample source", "Asset record", "SRC-001", "SRC-001-JOB-001", "ASSET-001", "Sample asset one", "Asset", "2026-01-15", 12500, "Sample funding", "Standard life", "CTX-001", "Sample context", "Sample structural hint that still needs classifier review")
        ) | Out-Null

    Add-Table `
        -Worksheet $setupSheet `
        -Name "tblAssetEvidenceRules" `
        -TopLeft "A7" `
        -Headers @("RuleId", "MatchField", "MatchText", "CategoryId", "CategoryName", "RulePriority", "RuleStatus") `
        -Rows @(
            ,@("RULE-001", "Description", "classifier review", "CAT-001", "Reviewed asset evidence", 1, "active")
        ) | Out-Null

    Add-Table `
        -Worksheet $setupSheet `
        -Name "tblAssetEvidenceOverrides" `
        -TopLeft "A12" `
        -Headers @("EvidenceId", "CategoryId", "CategoryName", "ClassifierSourceType", "ClassifierSourceLabel", "ClassifierRuleId", "OverrideReason", "ReviewStatus") `
        -Rows @(
            ,@("", "", "", "", "", "", "", "draft")
        ) | Out-Null

    [void]$setupSheet.Columns.AutoFit()

    foreach ($template in $templates) {
        [void]$workbook.Queries.Add($template.QueryName, $template.FormulaText, "Public asset evidence Power Query template from $($template.SourceFile).")
    }

    foreach ($template in $templates) {
        $querySheet = Add-Worksheet -Workbook $workbook -Name $template.SheetName
        [void](Add-LoadedQueryTable -Worksheet $querySheet -QueryName $template.QueryName -TableName $template.TableName)
    }

    [void]$setupSheet.Activate()
    $workbook.SaveAs($resolvedOutputPath, 51)
    Write-Host "Built $resolvedOutputPath"
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
