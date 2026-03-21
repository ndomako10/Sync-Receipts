#Requires -Version 5.0
<#
.SYNOPSIS
    Integration test: validates that New-AccountsTemplate.ps1 produces a correctly structured workbook.

.DESCRIPTION
    COM-dependent integration test. Requires Microsoft Excel. Not run in CI.
    Runs New-AccountsTemplate.ps1 against a temporary output path, then opens the
    workbook and asserts that the sheet name, table name, column headers, row count,
    and first-row values match the expected schema.

    Exit code 0 when all assertions pass; exit code 1 on any failure.

.NOTES
    Run from the repo root:
        PowerShell -NoProfile -ExecutionPolicy Bypass -File Tests\Integration\Invoke-NewAccountsTemplateTest.ps1
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$integrationRoot = $PSScriptRoot
$script          = Resolve-Path (Join-Path $integrationRoot '..\..\Scripts\New-AccountsTemplate.ps1')
$tempFile        = Join-Path $env:TEMP "AccountsTemplate_$(Get-Random).xlsx"

$failures = 0
$excel    = $null
$workbook = $null

function Assert-Equal {
    param([string]$Label, $Got, $Expected)
    if ($Got -eq $Expected) {
        Write-Host "  [PASS] $Label = ""$Got""" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] $Label: got ""$Got"" expected ""$Expected""" -ForegroundColor Red
        $script:failures++
    }
}

try {
    Write-Host ''
    Write-Host 'Running New-AccountsTemplate.ps1...' -ForegroundColor Cyan

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script -OutputPath $tempFile

    if ($LASTEXITCODE -ne 0) {
        Write-Host "[FAIL] New-AccountsTemplate.ps1 exited with code $LASTEXITCODE" -ForegroundColor Red
        exit 1
    }
    if (-not (Test-Path $tempFile)) {
        Write-Host '[FAIL] Output file was not created' -ForegroundColor Red
        exit 1
    }

    Write-Host 'Validating schema...' -ForegroundColor Cyan
    $excel    = New-Object -ComObject Excel.Application
    $excel.Visible       = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($tempFile, 0, $true)

    # Sheet exists
    $sheet = $null
    try { $sheet = $workbook.Sheets.Item('Accounts') } catch {}
    if ($null -eq $sheet) {
        Write-Host '  [FAIL] Sheet "Accounts" not found' -ForegroundColor Red
        $failures++
    } else {
        Write-Host '  [PASS] Sheet "Accounts" exists' -ForegroundColor Green

        # Table exists
        $lo = $null
        try { $lo = $sheet.ListObjects.Item('Accounts') } catch {}
        if ($null -eq $lo) {
            Write-Host '  [FAIL] Table "Accounts" not found' -ForegroundColor Red
            $failures++
        } else {
            Write-Host '  [PASS] Table "Accounts" exists' -ForegroundColor Green

            # Column count
            $colCount = $lo.ListColumns.Count
            Assert-Equal 'Column count' $colCount 6

            # Headers
            $expectedHeaders = @('Last 4', 'Method', 'Holder', 'Institution', 'Account', 'Status')
            for ($i = 1; $i -le $expectedHeaders.Count; $i++) {
                $hdr = $lo.HeaderRowRange.Cells.Item(1, $i).Value2
                Assert-Equal "Header col $i" $hdr $expectedHeaders[$i - 1]
            }

            # Row count (data body rows, excluding header and total)
            $rowCount = $lo.DataBodyRange.Rows.Count
            Assert-Equal 'Example row count' $rowCount 4

            # First row Last4
            $firstLast4 = $lo.DataBodyRange.Cells.Item(1, 1).Value2
            Assert-Equal 'First row Last 4' $firstLast4 '1234'

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($lo) | Out-Null
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
    }

    $total  = 11
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
    if (Test-Path $tempFile) {
        Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
    }
}

if ($failures -gt 0) { exit 1 } else { exit 0 }
