#Requires -Version 5.0
<#
.SYNOPSIS
    Integration test: syncs the fixture receipts folder and validates the Flag column.

.DESCRIPTION
    COM-dependent integration test. Requires Microsoft Excel. Not run in CI.
    Runs Sync-Receipts.ps1 against a controlled fixture folder containing 7 zero-byte
    placeholder receipt files, then opens the output workbook and asserts the Flag
    column value for each row.

    Exit code 0 when all assertions pass; exit code 1 on any failure.

.NOTES
    Run from the repo root:
        PowerShell -NoProfile -ExecutionPolicy Bypass -File Tests\Integration\Invoke-SyncReceiptsTest.ps1

    Precondition: none -- all config files are self-contained in Tests\Integration\Config\.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$integrationRoot = $PSScriptRoot
$fixtureRoot     = Join-Path $integrationRoot 'Fixture'
$configRoot      = Join-Path $integrationRoot 'Config'
$syncScript      = Resolve-Path (Join-Path $integrationRoot '..\..\Scripts\Sync-Receipts.ps1')

$tempDir    = Join-Path $env:TEMP "SyncReceiptsIntTest_$(Get-Random)"
$failures   = 0
$excel      = $null
$workbook   = $null

$expected = @{
    '260101 Target -$20.00 Card 1234.pdf'   = ''
    '260102 Walmart -$15.00 Card 9999.pdf'  = 'Account inactive'
    '260103 Amazon -$30.00 Card 8888.pdf'   = 'Account not in Accounts.xlsx'
    '260104 Shell -$50.00 Gas.pdf'          = 'Unrecognised method'
    '260105 Costco -$80.00 Card.pdf'        = 'Could not parse filename'
    '260106 Starbucks -$5.00 Cash.pdf'      = ''
    '260107 BADFILENAME.pdf'                = 'Could not parse filename'
}

try {
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    Write-Host ''
    Write-Host 'Running Sync-Receipts.ps1 against integration fixture...' -ForegroundColor Cyan

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript `
        -ReceiptsRoot $fixtureRoot `
        -ConfigRoot   $configRoot `
        -YearMonth    '2601' `
        -WorkbookPath (Join-Path $tempDir '2026.xlsx')

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[FAIL] Sync-Receipts.ps1 exited with code $LASTEXITCODE" -ForegroundColor Red
        exit 1
    }

    $workbookPath = Join-Path $tempDir '2026.xlsx'
    if (-not (Test-Path $workbookPath)) {
        Write-Host '[FAIL] Output workbook was not created' -ForegroundColor Red
        exit 1
    }

    Write-Host 'Opening output workbook...' -ForegroundColor Cyan
    $excel    = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false
    $workbook = $excel.Workbooks.Open($workbookPath, 0, $true)
    $sheet    = $workbook.Sheets.Item('2601')

    # Read Flag column (col 9) keyed by File Name (col 1) display text
    $actual = @{}
    $row    = 2
    while ($true) {
        $cellVal = $sheet.Cells.Item($row, 1).Value2
        if ([string]::IsNullOrEmpty($cellVal)) { break }
        $flagVal = $sheet.Cells.Item($row, 9).Value2
        if ($null -eq $flagVal) { $flagVal = '' }
        $actual[$cellVal.ToString().Trim()] = $flagVal.ToString().Trim()
        $row++
    }

    Write-Host ''
    Write-Host 'Assertions:' -ForegroundColor Cyan
    foreach ($file in ($expected.Keys | Sort-Object)) {
        $exp = $expected[$file]
        $got = if ($actual.ContainsKey($file)) { $actual[$file] } else { '<row not found>' }
        if ($got -eq $exp) {
            Write-Host "  [PASS] $file  flag=""$got""" -ForegroundColor Green
        } else {
            Write-Host "  [FAIL] $file  got=""$got""  expected=""$exp""" -ForegroundColor Red
            $failures++
        }
    }

    $total = $expected.Count
    $passed = $total - $failures
    Write-Host ''
    Write-Host "Integration test: $passed/$total assertions passed." -ForegroundColor $(if ($failures -eq 0) { 'Green' } else { 'Red' })

} finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if ($failures -gt 0) { exit 1 } else { exit 0 }
