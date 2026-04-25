param(
    [int]$Port = 3000,
    [switch]$SkipStaticChecks
)

$ErrorActionPreference = "Stop"

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

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

if (-not $SkipStaticChecks) {
    Invoke-Checked "repo audit" {
        python tools\audit_capex_module.py
    }
    Invoke-Checked "formula lint" {
        python tools\lint_formulas.py modules\*.formula.txt
    }
}

$serverScript = Join-Path $repoRoot "tools\start_addin_dev_server.ps1"
$serverCommand = "powershell -NoProfile -ExecutionPolicy Bypass -File `"$serverScript`" -Port $Port"

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Warning "npm is not on PATH, so the Microsoft Excel sideload helper cannot run."
    Write-Warning "Starting the local HTTPS server instead. In Excel, sideload addin\manifest.xml manually."
    & $serverScript -Port $Port
    exit $LASTEXITCODE
}

Write-Host "==> starting Excel desktop sideload smoke test"
Write-Host "==> after Excel opens, use the task pane button: Setup + Install + Validate"

npx --yes office-addin-debugging start addin\manifest.xml desktop --app excel --no-debug --no-live-reload --dev-server "$serverCommand" --dev-server-port $Port
exit $LASTEXITCODE
