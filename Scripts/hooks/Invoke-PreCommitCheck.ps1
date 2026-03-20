#Requires -Version 5.0
<#
.SYNOPSIS
    Pre-commit checks: PSScriptAnalyzer lint and ASCII validation on staged .ps1 files.

.DESCRIPTION
    Called by .git/hooks/pre-commit via Scripts/Install-GitHooks.ps1.
    Reads staged file content from the git index (not the working copy) so
    unstaged edits are never flagged.

    Exits 0 if all checks pass; exits 1 if any check fails.
#>

$staged = & git diff --cached --name-only --diff-filter=ACM 2>$null |
    Where-Object { $_ -like '*.ps1' }

if (-not $staged) { exit 0 }

$repoRoot          = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$settingsPath      = Join-Path $repoRoot ".config\PSScriptAnalyzerSettings.psd1"
$hasSettings       = Test-Path $settingsPath
$analyzerAvailable = Get-Module -ListAvailable PSScriptAnalyzer -ErrorAction SilentlyContinue
$failed            = $false

foreach ($file in $staged) {
    $lines   = & git show ":$file" 2>$null
    $content = $lines -join "`n"

    # ASCII check -- no character above U+007F
    if ($content -match '[^\x00-\x7F]') {
        Write-Host "  FAIL [ASCII] ${file}: non-ASCII characters found" -ForegroundColor Red
        $failed = $true
    } else {
        Write-Host "  OK   [ASCII] $file" -ForegroundColor Green
    }

    # PSScriptAnalyzer check -- uses .config/PSScriptAnalyzerSettings.psd1 to match CI
    if ($analyzerAvailable) {
        try {
            $analyzerArgs = @{ ScriptDefinition = $content; Severity = 'Error','Warning' }
            if ($hasSettings) { $analyzerArgs['Settings'] = $settingsPath }
            $results = Invoke-ScriptAnalyzer @analyzerArgs
            if ($results) {
                foreach ($r in $results) {
                    Write-Host ("  FAIL [LINT]  ${file} line $($r.Line): " +
                        "$($r.Severity) $($r.RuleName) -- $($r.Message)") -ForegroundColor Red
                }
                $failed = $true
            } else {
                Write-Host "  OK   [LINT]  $file" -ForegroundColor Green
            }
        } catch {
            Write-Host "  WARN [LINT]  PSScriptAnalyzer error on ${file}: $_" -ForegroundColor Yellow
        }
    } else {
        Write-Host ("  SKIP [LINT]  PSScriptAnalyzer not installed -- " +
            "run: Install-Module PSScriptAnalyzer -Scope CurrentUser") -ForegroundColor Yellow
    }
}

if ($failed) {
    Write-Host ""
    Write-Host "pre-commit: checks failed. Fix the errors above and re-stage." -ForegroundColor Red
    exit 1
}

exit 0
