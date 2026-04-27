[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TargetWorkbookPath,

    [string]$OutputPath = "",

    [switch]$ReplaceExisting,

    [switch]$Force
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$templateDir = Join-Path $repoRoot "samples\power-query\asset-evidence"

function Resolve-InputPath {
    param([string]$PathValue)

    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return [System.IO.Path]::GetFullPath($PathValue)
    }

    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $PathValue))
}

function Get-DefaultOutputPath {
    param([string]$InputPath)

    $directory = Split-Path -Parent $InputPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)
    $extension = [System.IO.Path]::GetExtension($InputPath)
    return Join-Path $directory "$baseName.asset-evidence-pq$extension"
}

function Worksheet-Exists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    foreach ($sheet in $Workbook.Worksheets) {
        if ([string]::Compare($sheet.Name, $Name, $true) -eq 0) {
            return $true
        }
    }
    return $false
}

function Delete-Worksheet-IfExists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    foreach ($sheet in $Workbook.Worksheets) {
        if ([string]::Compare($sheet.Name, $Name, $true) -eq 0) {
            $sheet.Delete()
            return
        }
    }
}

function Query-Exists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    try {
        [void]$Workbook.Queries.Item($Name)
        return $true
    } catch {
        return $false
    }
}

function Delete-Query-IfExists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    try {
        $Workbook.Queries.Item($Name).Delete()
    } catch {
        return
    }
}

function Delete-Connection-IfExists {
    param(
        [object]$Workbook,
        [string]$Name
    )

    try {
        $Workbook.Connections.Item($Name).Delete()
    } catch {
        return
    }
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
    return $table
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
    [void]$Worksheet.Columns.AutoFit()
    return $table
}

function Add-AssetEvidenceSetup {
    param([object]$Workbook)

    $setupSheet = Add-Worksheet -Workbook $Workbook -Name "Asset Evidence Setup"
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
}

$resolvedTargetPath = Resolve-InputPath $TargetWorkbookPath

if (-not (Test-Path -LiteralPath $resolvedTargetPath -PathType Leaf)) {
    throw "Target workbook not found: $resolvedTargetPath"
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $resolvedOutputPath = [System.IO.Path]::GetFullPath((Get-DefaultOutputPath $resolvedTargetPath))
} else {
    $resolvedOutputPath = Resolve-InputPath $OutputPath
}

if ([string]::Compare($resolvedTargetPath, $resolvedOutputPath, $true) -eq 0) {
    throw "OutputPath must be a workbook copy, not the original target workbook."
}

if (Test-Path -LiteralPath $resolvedOutputPath) {
    if (-not $Force) {
        throw "Output workbook already exists. Pass -Force to replace it: $resolvedOutputPath"
    }
    Remove-Item -LiteralPath $resolvedOutputPath -Force
}

$outputDir = Split-Path -Parent $resolvedOutputPath
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
Copy-Item -LiteralPath $resolvedTargetPath -Destination $resolvedOutputPath

$queryDefinitions = @(
    @{ Name = "qAssetEvidence_Normalized"; Sheet = "PQ Asset Evidence Normalized"; Table = "tblAssetEvidence_Normalized" },
    @{ Name = "qAssetEvidence_Classified"; Sheet = "PQ Asset Evidence Classified"; Table = "tblAssetEvidence_Classified" },
    @{ Name = "qAssetEvidence_Linked"; Sheet = "PQ Asset Evidence Linked"; Table = "tblAssetEvidence_Linked" },
    @{ Name = "qAssetEvidence_Status"; Sheet = "PQ Asset Evidence Status"; Table = "tblAssetEvidence_Status" },
    @{ Name = "qAssetEvidence_ModelInputs"; Sheet = "PQ Asset Evidence Model Inputs"; Table = "tblAssetEvidence_ModelInputs" },
    @{ Name = "qQA_AssetEvidence_MappingQueue"; Sheet = "PQ Asset Evidence Mapping Queue"; Table = "tblQA_AssetEvidence_MappingQueue" }
)

$seedSheets = @("Asset Evidence Setup") + @($queryDefinitions | ForEach-Object { $_.Sheet })

$templates = foreach ($query in $queryDefinitions) {
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

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Open($resolvedOutputPath)

    $collidingSheets = @()
    foreach ($sheetName in $seedSheets) {
        if (Worksheet-Exists -Workbook $workbook -Name $sheetName) {
            $collidingSheets += $sheetName
        }
    }

    $collidingQueries = @()
    foreach ($template in $templates) {
        if (Query-Exists -Workbook $workbook -Name $template.QueryName) {
            $collidingQueries += $template.QueryName
        }
    }

    if (($collidingSheets.Count -gt 0 -or $collidingQueries.Count -gt 0) -and -not $ReplaceExisting) {
        $detail = @()
        if ($collidingSheets.Count -gt 0) { $detail += "sheets: $($collidingSheets -join ', ')" }
        if ($collidingQueries.Count -gt 0) { $detail += "queries: $($collidingQueries -join ', ')" }
        throw "Output workbook already has asset evidence objects. Pass -ReplaceExisting to replace only these objects: $($detail -join '; ')"
    }

    if ($ReplaceExisting) {
        if ($workbook.Worksheets.Count -le $collidingSheets.Count) {
            [void](Add-Worksheet -Workbook $workbook -Name "Asset Evidence Temp")
        }
        foreach ($sheetName in $seedSheets) {
            Delete-Worksheet-IfExists -Workbook $workbook -Name $sheetName
        }
        foreach ($template in $templates) {
            Delete-Connection-IfExists -Workbook $workbook -Name "Query - $($template.QueryName)"
            Delete-Query-IfExists -Workbook $workbook -Name $template.QueryName
        }
    }

    Add-AssetEvidenceSetup -Workbook $workbook

    foreach ($template in $templates) {
        [void]$workbook.Queries.Add($template.QueryName, $template.FormulaText, "Public asset evidence Power Query template from $($template.SourceFile).")
    }

    foreach ($template in $templates) {
        $querySheet = Add-Worksheet -Workbook $workbook -Name $template.SheetName
        [void](Add-LoadedQueryTable -Worksheet $querySheet -QueryName $template.QueryName -TableName $template.TableName)
    }

    if (Worksheet-Exists -Workbook $workbook -Name "Asset Evidence Temp") {
        Delete-Worksheet-IfExists -Workbook $workbook -Name "Asset Evidence Temp"
    }

    [void]$workbook.Worksheets.Item("Asset Evidence Setup").Activate()
    $workbook.Save()
    Write-Host "Installed asset evidence Power Query sheets into $resolvedOutputPath"
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
