param(
    [string]$Message = "Update public formula template",
    [string]$Branch = "main",
    [switch]$NoPush
)

$ErrorActionPreference = "Stop"

function Invoke-Step {
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

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $repoRoot

$currentBranch = (git branch --show-current).Trim()
if ($currentBranch -ne $Branch) {
    throw "Expected branch '$Branch' but current branch is '$currentBranch'."
}

Invoke-Step "audit" { python tools\audit_capex_module.py }
Invoke-Step "formula lint" { python tools\lint_formulas.py modules\*.formula.txt }
Invoke-Step "diff whitespace check" { git diff --check }

$status = git status --porcelain
if ($status) {
    Invoke-Step "stage changes" { git add -A }
    Invoke-Step "commit" { git commit -m $Message }
} else {
    Write-Host "==> no local changes to commit"
}

Invoke-Step "fetch origin" { git fetch origin }

$upstream = "origin/$Branch"
$aheadBehind = (git rev-list --left-right --count "HEAD...$upstream").Trim()
$parts = $aheadBehind -split "\s+"
$ahead = [int]$parts[0]
$behind = [int]$parts[1]

if ($behind -gt 0) {
    Invoke-Step "rebase $upstream" { git rebase $upstream }
}

if ($NoPush) {
    Write-Host "==> push skipped by -NoPush"
    exit 0
}

Invoke-Step "push $Branch" { git push origin $Branch }
Write-Host "==> public push complete"
