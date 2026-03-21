#Requires -Version 5.0
<#
.SYNOPSIS
    Pre-push check: runs the full Pester test suite before every push.

.DESCRIPTION
    Called by .git/hooks/pre-push via Scripts/Install-GitHooks.ps1.
    Exits 0 if all tests pass; exits 1 if any test fails, any test file fails to
    load, or a required module (Pester, PSScriptAnalyzer) cannot be found or installed.
#>

$pesterAvailable = Get-Module -ListAvailable Pester -ErrorAction SilentlyContinue |
    Where-Object { $_.Version -ge [version]'5.0' }

if (-not $pesterAvailable) {
    Write-Host ("pre-push: Pester 5+ is not installed. " +
        "Run: Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser") -ForegroundColor Red
    exit 1
}

$analyzerAvailable = Get-Module -ListAvailable PSScriptAnalyzer -ErrorAction SilentlyContinue
if (-not $analyzerAvailable) {
    Write-Host "pre-push: PSScriptAnalyzer not installed -- installing..." -ForegroundColor Yellow
    try {
        Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "pre-push: PSScriptAnalyzer installed." -ForegroundColor Green
    } catch {
        Write-Host ("pre-push: failed to install PSScriptAnalyzer: $_. " +
            "Run: Install-Module PSScriptAnalyzer -Scope CurrentUser") -ForegroundColor Red
        exit 1
    }
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

$failedContainers = @($result.Containers | Where-Object { $_.Result -eq 'Failed' })
if ($result.FailedCount -gt 0 -or $failedContainers.Count -gt 0) {
    Write-Host ""
    if ($result.FailedCount -gt 0) {
        Write-Host ("pre-push: $($result.FailedCount) test(s) failed. " +
            "Fix failures before pushing.") -ForegroundColor Red
    }
    if ($failedContainers.Count -gt 0) {
        Write-Host ("pre-push: $($failedContainers.Count) test file(s) failed to load. " +
            "Fix discovery errors before pushing.") -ForegroundColor Red
        foreach ($c in $failedContainers) {
            Write-Host "  - $($c.Item.Path)" -ForegroundColor Red
        }
    }
    exit 1
}

Write-Host "pre-push: all $($result.PassedCount) test(s) passed." -ForegroundColor Green
exit 0
