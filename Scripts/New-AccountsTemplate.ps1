<#
.SYNOPSIS
    Generates Config/Accounts.template.xlsx with the correct schema and data validation.

.DESCRIPTION
    Creates a formatted Excel workbook at Config\Accounts.template.xlsx containing:

      - A structured table (Accounts) with the following columns:
            Last 4       -- account number as it appears in receipt filenames
            Method       -- only populated when Last 4 is ambiguous across methods
            Holder       -- person the account belongs to
            Institution  -- bank, card issuer, or wallet provider
            Account      -- name of the balance-holding account
            Status       -- Active or Inactive (dropdown validation)

      - Dropdown validation on Method (loaded from Config\Methods.template.json, with
        Cash prepended; blank allowed) and Status (Active, Inactive; required).

      - Four example rows to illustrate the schema, covering common disambiguation
        scenarios (same Last 4 across Card/Checking/Check, and a blank-Method row);
        replace with real account data.

    Run this script whenever the Accounts schema changes and commit the result.
    Initialize-SyncReceipts.ps1 copies the template to Config\Accounts.xlsx on
    first setup.

    Requires Microsoft Excel (COM automation).

.NOTES
    Requires: PowerShell 5.0+, Microsoft Excel (COM automation)
    Tested on: Windows 10/11

.PARAMETER OutputPath
    Full path to write the accounts template workbook to. Defaults to
    Config\Accounts.template.xlsx in the repo root. Override to write the
    template to a different location (e.g. a test fixture directory).

.EXAMPLE
    PowerShell -NoProfile -ExecutionPolicy Bypass -File Scripts\New-AccountsTemplate.ps1

.EXAMPLE
    PowerShell -NoProfile -ExecutionPolicy Bypass -File Scripts\New-AccountsTemplate.ps1 -OutputPath C:\Temp\Accounts.template.xlsx
#>

param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) "Config\Accounts.template.xlsx")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot    = Split-Path $PSScriptRoot -Parent
$outPath     = $OutputPath
$methodsJson = Join-Path $repoRoot "Config\Methods.template.json"

# Load method tokens from Methods.template.json; Cash is always prepended.
$defaultMethods = @('Card', 'Check', 'Checking', 'Savings', 'Transfer', 'Wire')
$loadedMethods  = $defaultMethods
if (Test-Path $methodsJson) {
    try {
        $loadedMethods = Get-Content $methodsJson -Raw | ConvertFrom-Json
        Write-Host "  Loaded method tokens from Methods.template.json."
    } catch {
        Write-Host "  WARNING: Could not parse Methods.template.json -- using built-in defaults." -ForegroundColor Yellow
    }
} else {
    Write-Host "  WARNING: Methods.template.json not found -- using built-in defaults." -ForegroundColor Yellow
}
# Cash is always first; exclude it from the loaded list in case it was added manually.
$methodList = (@('Cash') + ($loadedMethods | Where-Object { $_ -ne 'Cash' })) -join ','
$statusList = "Active,Inactive"

# Column index constants
$COL_LAST4       = 1
$COL_METHOD      = 2
$COL_HOLDER      = 3
$COL_INSTITUTION = 4
$COL_ACCOUNT = 5
$COL_STATUS      = 6
$NUM_COLS        = 6

$excel    = $null
$workbook = $null

try {
    Write-Host "Starting Excel..."
    $excel                = New-Object -ComObject Excel.Application
    $excel.Visible        = $false
    $excel.DisplayAlerts  = $false
    $excel.ScreenUpdating = $false

    Write-Host "Building workbook..."
    $workbook = $excel.Workbooks.Add()

    # Remove extra sheets; keep only one
    while ($workbook.Sheets.Count -gt 1) {
        $workbook.Sheets.Item($workbook.Sheets.Count).Delete()
    }
    $sheet      = $workbook.Sheets.Item(1)
    $sheet.Name = "Accounts"

    # ---------------------------------------------------------------------------
    # Headers
    # ---------------------------------------------------------------------------
    $sheet.Cells.Item(1, $COL_LAST4      ).Value2 = "Last 4"
    $sheet.Cells.Item(1, $COL_METHOD     ).Value2 = "Method"
    $sheet.Cells.Item(1, $COL_HOLDER     ).Value2 = "Holder"
    $sheet.Cells.Item(1, $COL_INSTITUTION).Value2 = "Institution"
    $sheet.Cells.Item(1, $COL_ACCOUNT).Value2 = "Account"
    $sheet.Cells.Item(1, $COL_STATUS     ).Value2 = "Status"

    # ---------------------------------------------------------------------------
    # Example rows
    # ---------------------------------------------------------------------------
    $examples = @(
        @("1234", "Card",     "Account Holder", "Institution Name", "Checking Account Name", "Active"),
        @("5678", "",         "Account Holder", "Institution Name", "Checking Account Name", "Active"),
        @("9012", "Checking", "Account Holder", "Institution Name", "Checking Account Name", "Active"),
        @("9012", "Check",    "Account Holder", "Institution Name", "Checking Account Name", "Active")
    )
    $dataRows = $examples.Count
    for ($r = 0; $r -lt $dataRows; $r++) {
        for ($c = 0; $c -lt $NUM_COLS; $c++) {
            $sheet.Cells.Item($r + 2, $c + 1).Value2 = $examples[$r][$c]
        }
    }

    # ---------------------------------------------------------------------------
    # Structured table
    # ---------------------------------------------------------------------------
    $tableRange = $sheet.Range(
        $sheet.Cells.Item(1, 1),
        $sheet.Cells.Item($dataRows + 1, $NUM_COLS)
    )
    $table             = $sheet.ListObjects.Add(1, $tableRange, [System.Reflection.Missing]::Value, 1)
    $table.Name        = "Accounts"
    $table.TableStyle  = "TableStyleMedium2"
    Write-Host "  Table 'Accounts' created."

    # ---------------------------------------------------------------------------
    # Column widths
    # ---------------------------------------------------------------------------
    $sheet.Columns.Item($COL_LAST4      ).ColumnWidth = 10
    $sheet.Columns.Item($COL_METHOD     ).ColumnWidth = 12
    $sheet.Columns.Item($COL_HOLDER     ).ColumnWidth = 20
    $sheet.Columns.Item($COL_INSTITUTION).ColumnWidth = 20
    $sheet.Columns.Item($COL_ACCOUNT).ColumnWidth = 25
    $sheet.Columns.Item($COL_STATUS     ).ColumnWidth = 12

    # ---------------------------------------------------------------------------
    # Data validation
    # Applies to the data range plus a generous buffer for future rows.
    # xlValidateList=3  xlValidAlertStop=1  xlValidAlertInformation=3  xlBetween=1
    # ---------------------------------------------------------------------------

    # Method -- blank allowed; informational alert so missing values are permitted
    $methodRange = $sheet.Range(
        $sheet.Cells.Item(2, $COL_METHOD),
        $sheet.Cells.Item(1048576, $COL_METHOD)
    )
    $methodRange.Validation.Delete()
    $methodRange.Validation.Add(3, 3, 1, $methodList)
    $methodRange.Validation.IgnoreBlank    = $true
    $methodRange.Validation.InCellDropdown = $true
    Write-Host "  Validation added: Method ($methodList)."

    # Status -- required; stop alert enforces valid values
    $statusRange = $sheet.Range(
        $sheet.Cells.Item(2, $COL_STATUS),
        $sheet.Cells.Item(1048576, $COL_STATUS)
    )
    $statusRange.Validation.Delete()
    $statusRange.Validation.Add(3, 1, 1, $statusList)
    $statusRange.Validation.IgnoreBlank    = $false
    $statusRange.Validation.InCellDropdown = $true
    Write-Host "  Validation added: Status ($statusList)."

    # ---------------------------------------------------------------------------
    # Save
    # ---------------------------------------------------------------------------
    if (Test-Path $outPath) {
        Remove-Item $outPath -Force
        Write-Host "  Removed existing Accounts.template.xlsx."
    }

    # 51 = xlOpenXMLWorkbook (.xlsx)
    $workbook.SaveAs($outPath, 51)
    Write-Host "  [OK] Saved to Config\Accounts.template.xlsx"

} catch {
    Write-Host "  [FAIL] $_" -ForegroundColor Red
    exit 1
} finally {
    if ($workbook) { try { $workbook.Close($false) } catch {} }
    if ($excel)    { try { $excel.Quit()            } catch {} }
    if ($workbook) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    }
    if ($excel) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
