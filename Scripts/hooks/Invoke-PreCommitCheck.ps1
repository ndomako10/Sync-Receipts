#Requires -Version 5.0
<#
.SYNOPSIS
    Pre-commit checks: ASCII validation, PSScriptAnalyzer lint, and sensitive
    data detection on staged .ps1, .json, and .md files.

.DESCRIPTION
    Called by .git/hooks/pre-commit via Scripts/Install-GitHooks.ps1.
    Reads staged file content from the git index (not the working copy) so
    unstaged edits are never flagged.

    Sensitive data patterns are loaded from Config\SensitivePatterns.json
    (gitignored, user-customisable), falling back to
    Config\Templates\SensitivePatterns.template.json when absent.
    Additionally, Last4 values from Config\Accounts.xlsx are loaded at runtime
    and checked as dynamic patterns so real account numbers are caught before
    they reach the remote.

    Exits 0 if all checks pass; exits 1 if any check fails.
#>

. (Join-Path $PSScriptRoot 'Invoke-SensitiveDataCheck.ps1')

$stagedAll = & git diff --cached --name-only --diff-filter=ACM 2>$null |
    Where-Object { $_ -like '*.ps1' -or $_ -like '*.json' -or $_ -like '*.md' }

if (-not $stagedAll) { exit 0 }

$repoRoot          = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$settingsPath      = Join-Path $repoRoot ".config\PSScriptAnalyzerSettings.psd1"
$hasSettings       = Test-Path $settingsPath
$analyzerAvailable = Get-Module -ListAvailable PSScriptAnalyzer -ErrorAction SilentlyContinue
$failed            = $false

# --- Load sensitive data patterns -------------------------------------------

$patternsPath  = Join-Path $repoRoot "Config\SensitivePatterns.json"
$templatePath  = Join-Path $repoRoot "Config\Templates\SensitivePatterns.template.json"
$patternSource = if (Test-Path $patternsPath) { $patternsPath } elseif (Test-Path $templatePath) { $templatePath } else { $null }

$sensitivePatterns = @()
if ($patternSource) {
    try {
        $sensitivePatterns = @((Get-Content $patternSource -Raw | ConvertFrom-Json).patterns)
    } catch {
        Write-Host "  WARN [SENSITIVE] Could not load patterns from ${patternSource}: $_" -ForegroundColor Yellow
    }
}

# Load dynamic account patterns from Config\Accounts.xlsx (gitignored)
$accountsPath = Join-Path $repoRoot "Config\Accounts.xlsx"
if (Test-Path $accountsPath) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($accountsPath)
        try {
            $ssEntry = $zip.GetEntry('xl/sharedStrings.xml')
            $wsEntry = $zip.GetEntry('xl/worksheets/sheet1.xml')
            if ($ssEntry -and $wsEntry) {
                $ssReader = [System.IO.StreamReader]::new($ssEntry.Open())
                $wsReader = [System.IO.StreamReader]::new($wsEntry.Open())
                $ssXml = $ssReader.ReadToEnd()
                $wsXml = $wsReader.ReadToEnd()
                $ssReader.Dispose()
                $wsReader.Dispose()
                $last4s = @(Get-AccountsLast4 -SharedStringsXml $ssXml -WorksheetXml $wsXml)
                foreach ($last4 in $last4s) {
                    $sensitivePatterns += [PSCustomObject]@{
                        name      = "account-last4-${last4}"
                        regex     = "\b${last4}\b"
                        fileTypes = @('.ps1', '.md', '.json', '.txt')
                        message   = "Possible account number (Last4: ${last4} from Accounts.xlsx)"
                        allowlist = @()
                    }
                }
            }
        } finally {
            $zip.Dispose()
        }
    } catch {
        Write-Host "  WARN [SENSITIVE] Could not read Accounts.xlsx for account patterns: $_" -ForegroundColor Yellow
    }
}

# --- Per-file checks ---------------------------------------------------------

foreach ($file in $stagedAll) {
    $lines   = & git show ":$file" 2>$null
    $content = $lines -join "`n"

    # ASCII check -- no character above U+007F
    if ($content -match '[^\x00-\x7F]') {
        Write-Host "  FAIL [ASCII] ${file}: non-ASCII characters found" -ForegroundColor Red
        $failed = $true
    } else {
        Write-Host "  OK   [ASCII] $file" -ForegroundColor Green
    }

    # PSScriptAnalyzer check (.ps1 only) -- uses .config/PSScriptAnalyzerSettings.psd1 to match CI
    if ($file -like '*.ps1') {
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

    # Sensitive data check -- patterns from SensitivePatterns.json + dynamic Accounts.xlsx Last4s
    if ($sensitivePatterns.Count -gt 0) {
        $findings = @(Invoke-SensitiveDataCheck -Content $content -FileName $file -Patterns $sensitivePatterns)
        if ($findings.Count -gt 0) {
            foreach ($finding in $findings) {
                Write-Host ("  FAIL [SENSITIVE] ${file} line $($finding.LineNumber): " +
                    "$($finding.Message)") -ForegroundColor Red
                Write-Host "       $($finding.Line.Trim())" -ForegroundColor DarkRed
            }
            $failed = $true
        } else {
            Write-Host "  OK   [SENSITIVE] $file" -ForegroundColor Green
        }
    }
}

if ($failed) {
    Write-Host ""
    Write-Host "pre-commit: checks failed. Fix the errors above and re-stage." -ForegroundColor Red
    exit 1
}

exit 0
