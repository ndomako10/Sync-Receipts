#Requires -Version 5.0
<#
.SYNOPSIS
    Pre-push check: runs the full Pester test suite before every push.

.DESCRIPTION
    Called by .git/hooks/pre-push via Scripts/Install-GitHooks.ps1.
    Exits 0 if all tests pass; exits 1 if any test fails or Pester is missing.
#>

$pesterAvailable = Get-Module -ListAvailable Pester -ErrorAction SilentlyContinue |
    Where-Object { $_.Version -ge [version]'5.0' }

if (-not $pesterAvailable) {
    Write-Host ("pre-push: Pester 5+ is not installed. " +
        "Run: Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser") -ForegroundColor Red
    exit 1
}

Import-Module Pester -MinimumVersion 5.0 -Force

$repoRoot  = Split-Path $PSScriptRoot -Parent | Split-Path -Parent
$testsPath = Join-Path $repoRoot "Tests"

Write-Host "pre-push: running tests in $testsPath..." -ForegroundColor Cyan
$cfg                  = New-PesterConfiguration
$cfg.Run.Path         = $testsPath
$cfg.Output.Verbosity = 'Detailed'
$cfg.Run.PassThru     = $true
$result = Invoke-Pester -Configuration $cfg

if ($result.FailedCount -gt 0) {
    Write-Host ""
    Write-Host ("pre-push: $($result.FailedCount) test(s) failed. " +
        "Fix failures before pushing.") -ForegroundColor Red
    exit 1
}

Write-Host "pre-push: all $($result.PassedCount) test(s) passed." -ForegroundColor Green
exit 0
