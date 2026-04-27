param(
    [int]$Port = 3000,
    [switch]$SkipStaticChecks,
    [switch]$Yes,
    [switch]$NoInstall
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
Set-Location $repoRoot

function Use-InstalledNodePath {
    if (Get-Command npm -ErrorAction SilentlyContinue) {
        return
    }

    $nodeInstall = Join-Path $env:ProgramFiles "nodejs"
    $npmCmd = Join-Path $nodeInstall "npm.cmd"
    if (Test-Path $npmCmd) {
        $env:Path = "$nodeInstall;$env:Path"
    }
}

function Invoke-Checked {
    param(
        [string]$Label,
        [scriptblock]$Command
    )

    Write-Host "==> $Label"
    & $Command
    if ($LASTEXITCODE -ne 0) {
        throw "$Label failed with exit code $LASTEXITCODE"
    }
}

Write-Host "Governed Excel Formula Modules - Add-In Launcher"
Write-Host ""
Write-Host "This starts the local Excel add-in and launches Excel for sideload testing."
Write-Host "It does not edit a workbook by itself."
Write-Host ""
Write-Warning "Use a workbook copy. Do not run setup or apply buttons in a production workbook."

if (-not $Yes) {
    $answer = Read-Host "Confirm you will use a workbook copy before running setup buttons (Y/N)"
    if ($answer -notin @("Y", "y", "YES", "Yes", "yes")) {
        Write-Host "Cancelled. Make a workbook copy, then run Start-AddIn.ps1 again."
        exit 1
    }
}

Use-InstalledNodePath

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Warning "npm was not found. Install Node.js LTS, then rerun Start-AddIn.ps1."
    Write-Warning "Fallback: the existing smoke helper can start the local server and print manual sideload guidance."
    & (Join-Path $repoRoot "tools\start_addin_smoke_test.ps1") -Port $Port -SkipStaticChecks:$SkipStaticChecks
    exit $LASTEXITCODE
}

if (-not $NoInstall -and -not (Test-Path (Join-Path $repoRoot "node_modules"))) {
    Invoke-Checked "installing npm dependencies" {
        npm install
    }
}

Write-Host "==> launching Excel with the sideloaded add-in"
Write-Host "==> after Excel opens, open a workbook copy before clicking setup buttons"

& (Join-Path $repoRoot "tools\start_addin_smoke_test.ps1") -Port $Port -SkipStaticChecks:$SkipStaticChecks
exit $LASTEXITCODE
