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

function Stop-StaleDevServer {
    param(
        [int]$Port,
        [string]$RepoRoot
    )

    $connections = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    if (-not $connections) {
        return
    }

    $localTaskpane = Get-Content -Path (Join-Path $RepoRoot "addin\taskpane.js") -Raw
    $servedTaskpane = ""
    try {
        $probeUrl = ("https" + "://localhost:${Port}/addin/taskpane.js")
        $servedTaskpane = (Invoke-WebRequest -UseBasicParsing $probeUrl).Content
    } catch {
        Write-Warning "Port $Port is in use, but the add-in server probe failed: $($_.Exception.Message)"
    }

    if ($servedTaskpane -eq $localTaskpane) {
        Write-Host "==> existing dev server on port $Port is serving this checkout"
        return
    }

    $processIds = $connections | Select-Object -ExpandProperty OwningProcess -Unique
    foreach ($processId in $processIds) {
        if ($processId -and $processId -ne $PID) {
            $process = Get-CimInstance Win32_Process -Filter "ProcessId=$processId" -ErrorAction SilentlyContinue
            Write-Warning "Stopping stale dev server process $processId on port ${Port}: $($process.CommandLine)"
            Stop-Process -Id $processId -Force
        }
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot
Use-InstalledNodePath

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
Stop-StaleDevServer -Port $Port -RepoRoot $repoRoot

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    Write-Warning "npm is not on PATH, so the Microsoft Excel sideload helper cannot run."
    Write-Warning "Starting the local HTTPS server instead. In Excel, sideload addin\manifest.xml manually."
    & $serverScript -Port $Port
    exit $LASTEXITCODE
}

Write-Host "==> starting Excel desktop sideload smoke test"
Write-Host "==> after Excel opens, use: Setup + Install + Validate + Outputs"

npx --yes office-addin-debugging start addin\manifest.xml desktop --app excel --no-debug --no-live-reload --dev-server "$serverCommand" --dev-server-port $Port
exit $LASTEXITCODE
