#Requires -Version 5.0
<#
.SYNOPSIS
    Integration test: idempotency smoke test for Initialize-SyncReceipts.ps1.

.DESCRIPTION
    COM-dependent integration test. Requires Microsoft Excel (for shortcut creation
    via WScript.Shell). Not run in CI.

    Pre-creates Config\Config.ini (pointing at temp directories) so that all
    Read-Host prompts are bypassed, then runs Initialize-SyncReceipts.ps1 and
    asserts that it exits 0 and creates the 4 expected shortcut files in WORKBOOKS_ROOT.

    Config\Config.ini is backed up before the test and restored afterwards.

    Exit code 0 when all assertions pass; exit code 1 on any failure.

.NOTES
    Run from the repo root:
        PowerShell -NoProfile -ExecutionPolicy Bypass -File Tests\Integration\Invoke-InitializeSyncReceiptsTest.ps1

    Precondition: Config\Accounts.xlsx, Config\Categories.json, and Config\Methods.json
    must already exist (setup must have been run at least once). The test does not modify
    these files; it only exercises the shortcut-creation and skip-copy paths.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$integrationRoot = $PSScriptRoot
$repoRoot        = Resolve-Path (Join-Path $integrationRoot '..\..')
$initScript      = Join-Path $repoRoot 'Scripts\Initialize-SyncReceipts.ps1'
$configIni       = Join-Path $repoRoot 'Config\Config.ini'
$configIniBak    = Join-Path $env:TEMP "Config_bak_$(Get-Random).ini"

$tempReceipts  = Join-Path $env:TEMP "SyncReceiptsInitTest_$(Get-Random)"
$tempWorkbooks = Join-Path $env:TEMP "SyncReceiptsInitTest_wb_$(Get-Random)"
$failures      = 0

# Verify preconditions
$preconditions = @(
    (Join-Path $repoRoot 'Config\Accounts.xlsx'),
    (Join-Path $repoRoot 'Config\Categories.json'),
    (Join-Path $repoRoot 'Config\Methods.json')
)
foreach ($f in $preconditions) {
    if (-not (Test-Path $f)) {
        Write-Host "[FAIL] Precondition not met: $f does not exist." -ForegroundColor Red
        Write-Host "       Run Scripts\Initialize-SyncReceipts.ps1 at least once before this test." -ForegroundColor Yellow
        exit 1
    }
}

try {
    # Create temp directories
    New-Item -ItemType Directory -Path $tempReceipts  | Out-Null
    New-Item -ItemType Directory -Path $tempWorkbooks | Out-Null

    # Back up existing Config.ini
    $hadConfig = Test-Path $configIni
    if ($hadConfig) {
        Copy-Item $configIni $configIniBak
    }

    # Write test Config.ini
    @"
RECEIPTS_ROOT=$tempReceipts
WORKBOOKS_ROOT=$tempWorkbooks
"@ | Set-Content $configIni

    Write-Host ''
    Write-Host 'Running Initialize-SyncReceipts.ps1 (idempotency pass)...' -ForegroundColor Cyan

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $initScript

    $exitCode = $LASTEXITCODE
    if ($exitCode -eq 0) {
        Write-Host '  [PASS] Exit code 0' -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Exit code $exitCode (expected 0)" -ForegroundColor Red
        $failures++
    }

    Write-Host ''
    Write-Host 'Asserting shortcuts...' -ForegroundColor Cyan

    $expectedLnk = @(
        'Run Sync Receipts.lnk',
        'Run Sync Month Receipts.lnk',
        'Run Sync Year Receipts.lnk',
        'Run Sync All Receipts.lnk'
    )
    foreach ($lnk in $expectedLnk) {
        $lnkPath = Join-Path $tempWorkbooks $lnk
        if (Test-Path $lnkPath) {
            Write-Host "  [PASS] $lnk" -ForegroundColor Green
        } else {
            Write-Host "  [FAIL] $lnk not found in WORKBOOKS_ROOT" -ForegroundColor Red
            $failures++
        }
    }

    $total  = 1 + $expectedLnk.Count
    $passed = $total - $failures
    Write-Host ''
    Write-Host "Integration test: $passed/$total assertions passed." -ForegroundColor $(if ($failures -eq 0) { 'Green' } else { 'Red' })

} finally {
    # Restore Config.ini
    if ($hadConfig) {
        Copy-Item $configIniBak $configIni -Force
        Remove-Item $configIniBak -Force -ErrorAction SilentlyContinue
    } else {
        Remove-Item $configIni -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path $tempReceipts)  { Remove-Item $tempReceipts  -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $tempWorkbooks) { Remove-Item $tempWorkbooks -Recurse -Force -ErrorAction SilentlyContinue }
}

if ($failures -gt 0) { exit 1 } else { exit 0 }
