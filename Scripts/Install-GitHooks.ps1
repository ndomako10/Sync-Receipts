#Requires -Version 5.0
<#
.SYNOPSIS
    Installs local git hooks from Scripts/hooks/ into .git/hooks/.

.DESCRIPTION
    Copies the pre-commit, pre-push, and commit-msg hook entry points from the
    committed Scripts/hooks/ directory into .git/hooks/ so git executes them
    automatically.

    Safe to re-run -- existing hooks are overwritten with the latest source.
    Called automatically by Initialize-SyncReceipts.ps1 during first-time setup.

.EXAMPLE
    PowerShell -NoProfile -ExecutionPolicy Bypass -File Scripts\Install-GitHooks.ps1
#>

$hooksSource = Join-Path $PSScriptRoot "hooks"
$repoRoot    = Split-Path $PSScriptRoot -Parent
$gitHooksDir = Join-Path $repoRoot ".git\hooks"

if (-not (Test-Path $gitHooksDir)) {
    Write-Host "  SKIP [hooks] .git\hooks not found -- is this a git repository?" -ForegroundColor Yellow
    return
}

foreach ($name in @('pre-commit', 'pre-push', 'commit-msg')) {
    $src = Join-Path $hooksSource $name
    $dst = Join-Path $gitHooksDir $name
    try {
        Copy-Item $src $dst -Force
        Write-Host "  OK   Installed .git\hooks\$name" -ForegroundColor Green
    } catch {
        Write-Host "  FAIL Could not install .git\hooks\$name -- $_" -ForegroundColor Red
    }
}
