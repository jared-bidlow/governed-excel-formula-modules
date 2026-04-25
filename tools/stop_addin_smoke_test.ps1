$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
    throw "Node.js/npm is required for the Office.js smoke-test helper."
}

npx --yes office-addin-debugging stop addin\manifest.xml desktop
exit $LASTEXITCODE
