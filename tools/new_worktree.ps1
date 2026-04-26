param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern("^[A-Za-z0-9][A-Za-z0-9._-]*$")]
    [string]$Name,

    [string]$BranchPrefix = "codex",
    [string]$BaseBranch = "main",
    [string]$Path,
    [switch]$NoFetch
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

$repoName = Split-Path $repoRoot -Leaf
$branchName = "$BranchPrefix/$Name"
if ($BranchPrefix -eq "") {
    $branchName = $Name
}

if (-not $Path) {
    $safeName = $Name -replace "[^A-Za-z0-9._-]", "-"
    $Path = Join-Path (Split-Path $repoRoot -Parent) "$repoName-$safeName"
}

$targetPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (Test-Path $targetPath) {
    throw "Target path already exists: $targetPath"
}

if (-not $NoFetch) {
    Invoke-Step "fetch origin" { git fetch origin }
}

$baseRef = "origin/$BaseBranch"
Invoke-Step "verify base ref $baseRef" { git rev-parse --verify $baseRef }
Invoke-Step "create worktree $branchName" { git worktree add -b $branchName $targetPath $baseRef }

Write-Host ""
Write-Host "Created worktree:"
Write-Host "  path:   $targetPath"
Write-Host "  branch: $branchName"
Write-Host ""
Write-Host "Next:"
Write-Host "  cd `"$targetPath`""
Write-Host "  git status"
