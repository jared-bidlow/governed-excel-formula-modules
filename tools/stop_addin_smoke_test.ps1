param(
    [int]$Port = 3000
)

$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

$connections = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
$processIds = $connections | Select-Object -ExpandProperty OwningProcess -Unique

foreach ($processId in $processIds) {
    if ($processId -and $processId -ne $PID) {
        Write-Host "==> stopping local add-in server process $processId on port $Port"
        Stop-Process -Id $processId -Force
    }
}

if (Get-Command npm -ErrorAction SilentlyContinue) {
    npx --yes office-addin-debugging stop addin\manifest.xml desktop
    exit $LASTEXITCODE
}

Write-Host "==> npm is not on PATH; skipped Microsoft Office debugging stop command"
exit 0
