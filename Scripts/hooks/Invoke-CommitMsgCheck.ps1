#Requires -Version 5.0
<#
.SYNOPSIS
    Validates that a commit message follows Conventional Commits format.

.DESCRIPTION
    Called by .git/hooks/commit-msg via Scripts/Install-GitHooks.ps1.
    Reads the commit message from the file passed by git as $1 and validates
    it against the Conventional Commits pattern.

    Merge commits, revert commits, and fixup/squash commits are exempt.

    Exits 0 if the message is valid; exits 1 otherwise.

.PARAMETER MessageFile
    Path to the file containing the commit message (passed by git as $1).
#>

param([string]$MessageFile)

$msg = (Get-Content $MessageFile -Raw).Trim()

# Strip comment lines that git injects (lines starting with #)
$msg = ($msg -split "`n" | Where-Object { $_ -notmatch '^\s*#' }) -join "`n"
$msg = $msg.Trim()

# Exempt: Merge, Revert, fixup!, squash!
if ($msg -match '^(Merge |Revert |fixup!|squash!)') { exit 0 }

# Conventional Commits pattern
$types   = 'feat|fix|docs|test|refactor|chore|ci|perf|style|build'
$pattern = "^($types)(\([a-z0-9._-]+\))?(!)?: .+"

if ($msg -match $pattern) { exit 0 }

Write-Host ""
Write-Host "commit-msg: commit message does not follow Conventional Commits." -ForegroundColor Red
Write-Host ""
Write-Host "  Expected format:  type(scope): description" -ForegroundColor Yellow
Write-Host "  Got:              $($msg -split '`n' | Select-Object -First 1)" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Valid types: feat, fix, docs, test, refactor, chore, ci, perf, style, build" -ForegroundColor Yellow
Write-Host "  Example:     feat(categories): support nested subcategories" -ForegroundColor Yellow
Write-Host ""

exit 1
