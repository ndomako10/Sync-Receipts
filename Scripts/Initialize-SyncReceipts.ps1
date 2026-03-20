<#
.SYNOPSIS
    One-time setup script for Sync-Receipts on a new machine.

.DESCRIPTION
    Performs all first-time configuration steps needed before running Sync-Receipts:

      1. Checks that PowerShell 5.0+ and Microsoft Excel are installed.
      2. Reads RECEIPTS_ROOT from Config\Config.env, or prompts for the path and creates
         Config\Config.env from Config\Config.template.env.
      3. Creates the RECEIPTS_ROOT folder if it does not already exist.
      4. Copies Accounts.template.xlsx to RECEIPTS_ROOT\Accounts.xlsx (skipped if present).
      5. Creates Windows shortcut (.lnk) files in RECEIPTS_ROOT that point to the batch
         launchers in the script directory, with WorkingDirectory set so UNC-path shortcuts
         open without the "UNC paths are not supported" CMD error.

    Run via Setup.bat (recommended), or directly:
        PowerShell -NoProfile -ExecutionPolicy Bypass -File Initialize-SyncReceipts.ps1

    The script is safe to re-run -- steps that are already complete are skipped.

.NOTES
    Requires: PowerShell 5.0+, Microsoft Excel (COM automation)
    Tested on: Windows 10/11

    Steps performed:
      1. Checks prerequisites (PowerShell 5.0+, Excel)
      2. Creates Config\Config.env from template (prompts for RECEIPTS_ROOT, WORKBOOKS_ROOT)
      3. Creates RECEIPTS_ROOT folder if absent
      4. Creates WORKBOOKS_ROOT folder if absent and different from RECEIPTS_ROOT
      5. Copies Accounts.template.xlsx to Config\Accounts.xlsx (skipped if present)
      6. Creates .lnk shortcuts in RECEIPTS_ROOT pointing to Launchers\
      7. Installs local git hooks from Scripts\hooks\ into .git\hooks\
#>

$scriptDir    = $PSScriptRoot
$repoRoot     = Split-Path $scriptDir -Parent
$launchersDir = Join-Path $repoRoot "Launchers"

function Write-Step {
    <#
    .SYNOPSIS
        Writes a step-header message to the console in cyan.

    .DESCRIPTION
        Used internally by Initialize-SyncReceipts.ps1 to announce the start of each setup phase.
        Output is indented by two spaces and coloured cyan to stand out from
        sub-step OK/SKIP/FAIL lines.

    .PARAMETER msg
        The step description to display.

    .EXAMPLE
        Write-Step "Checking prerequisites..."
        # Writes "  Checking prerequisites..." in cyan.
    #>
    param([string]$msg)
    Write-Host "  $msg" -ForegroundColor Cyan
}

function Write-OK {
    <#
    .SYNOPSIS
        Writes a success line to the console in green.

    .DESCRIPTION
        Used internally by Initialize-SyncReceipts.ps1 to confirm that a setup action completed
        successfully. Output is prefixed with "[OK]" and coloured green.

    .PARAMETER msg
        The success message to display.

    .EXAMPLE
        Write-OK "PowerShell 5.1.26100.2161"
        # Writes "  [OK]   PowerShell 5.1.26100.2161" in green.
    #>
    param([string]$msg)
    Write-Host "  [OK]   $msg" -ForegroundColor Green
}

function Write-Skip {
    <#
    .SYNOPSIS
        Writes a skipped-step line to the console in yellow.

    .DESCRIPTION
        Used internally by Initialize-SyncReceipts.ps1 when a setup action is intentionally
        bypassed because the target already exists or the step is not needed.
        Output is prefixed with "[SKIP]" and coloured yellow.

    .PARAMETER msg
        The skip-reason message to display.

    .EXAMPLE
        Write-Skip "Accounts.xlsx already exists -- edit it directly to update your accounts"
        # Writes "  [SKIP] Accounts.xlsx already exists ..." in yellow.
    #>
    param([string]$msg)
    Write-Host "  [SKIP] $msg" -ForegroundColor Yellow
}

function Write-Fail {
    <#
    .SYNOPSIS
        Writes a failure line to the console in red.

    .DESCRIPTION
        Used internally by Initialize-SyncReceipts.ps1 when a setup action fails and the script
        cannot continue. Output is prefixed with "[FAIL]" and coloured red.
        The caller is responsible for calling exit after Write-Fail when the
        error is fatal.

    .PARAMETER msg
        The failure message to display.

    .EXAMPLE
        Write-Fail "Microsoft Excel does not appear to be installed."
        # Writes "  [FAIL] Microsoft Excel does not appear to be installed." in red.
    #>
    param([string]$msg)
    Write-Host "  [FAIL] $msg" -ForegroundColor Red
}

Write-Host ""
Write-Host "Sync-Receipts Setup" -ForegroundColor Cyan
Write-Host "-------------------" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# 1. Check prerequisites
# ---------------------------------------------------------------------------

Write-Step "Checking prerequisites..."

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Fail "PowerShell 5.0 or later is required (found $($PSVersionTable.PSVersion))."
    exit 1
}
Write-OK "PowerShell $($PSVersionTable.PSVersion)"

$excelInstalled = $false
try {
    $testExcel = New-Object -ComObject Excel.Application
    $testExcel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($testExcel) | Out-Null
    $excelInstalled = $true
} catch {}
if (-not $excelInstalled) {
    Write-Fail "Microsoft Excel does not appear to be installed."
    exit 1
}
Write-OK "Microsoft Excel"

# ---------------------------------------------------------------------------
# 2. Resolve RECEIPTS_ROOT -- read Config.env or prompt
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Configuring RECEIPTS_ROOT..."

$configPath = Join-Path $repoRoot "Config\Config.env"
$receiptsRoot  = $null
$workbooksRoot = $null

if (Test-Path $configPath) {
    foreach ($line in (Get-Content $configPath)) {
        if ($line -match '^RECEIPTS_ROOT=(.+)$') {
            $receiptsRoot = $Matches[1].Trim().TrimEnd('\')
        }
        if ($line -match '^WORKBOOKS_ROOT=(.+)$') {
            $workbooksRoot = $Matches[1].Trim().TrimEnd('\')
        }
    }
    if (-not $receiptsRoot) {
        Write-Fail "Config.env exists but RECEIPTS_ROOT could not be parsed."
        exit 1
    }
    if (-not $workbooksRoot) { $workbooksRoot = $receiptsRoot }
    Write-OK "Config.env found -- RECEIPTS_ROOT: $receiptsRoot"
} else {
    Write-Host ""
    Write-Host "  Config.env not found. Enter the path to your receipts root folder." -ForegroundColor Yellow
    Write-Host "  Example: \\Server\Share\Receipts  or  C:\Users\You\Documents\Receipts" -ForegroundColor Yellow
    Write-Host ""
    $userInput = Read-Host "  RECEIPTS_ROOT"
    $receiptsRoot = $userInput.Trim().TrimEnd('\')

    if (-not $receiptsRoot) {
        Write-Fail "No path entered. Aborting."
        exit 1
    }

    Write-Host ""
    Write-Host "  Enter the folder where per-year workbooks will be written." -ForegroundColor Yellow
    Write-Host "  Press Enter to use the same location as your receipts root." -ForegroundColor Yellow
    Write-Host ""
    $wbInput = Read-Host "  WORKBOOKS_ROOT [$receiptsRoot]"
    $workbooksRoot = if ($wbInput.Trim()) { $wbInput.Trim().TrimEnd('\') } else { $receiptsRoot }

    try {
        $template = Get-Content (Join-Path $repoRoot "Config\Config.template.env") -Raw
        $config   = $template `
            -replace 'RECEIPTS_ROOT=.*',       "RECEIPTS_ROOT=$receiptsRoot" `
            -replace 'RECEIPTS_ROOT_LOCAL=.*', "RECEIPTS_ROOT_LOCAL=$receiptsRoot" `
            -replace 'WORKBOOKS_ROOT=.*',      "WORKBOOKS_ROOT=$workbooksRoot"
        Set-Content $configPath $config -Encoding ASCII
        Write-OK "Created Config.env (set RECEIPTS_ROOT=$receiptsRoot)"
        if ($workbooksRoot -ne $receiptsRoot) {
            Write-OK "WORKBOOKS_ROOT: $workbooksRoot"
        }
        Write-Host ""
        Write-Host "  NOTE: If your receipts root has a local drive-letter equivalent, edit" -ForegroundColor Yellow
        Write-Host "  Config.env and update RECEIPTS_ROOT_LOCAL. This is only used by tests." -ForegroundColor Yellow
    } catch {
        Write-Fail "Could not create Config.env -- $_"
        exit 1
    }
}

# ---------------------------------------------------------------------------
# 3. Create RECEIPTS_ROOT if it does not exist
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Checking receipts root folder..."

if (Test-Path $receiptsRoot) {
    Write-OK "Folder exists: $receiptsRoot"
} else {
    $confirm = Read-Host "  Folder does not exist. Create '$receiptsRoot'? [Y/n]"
    if ($confirm -ne '' -and $confirm -notmatch '^[Yy]') {
        Write-Fail "Folder not created. Aborting."
        exit 1
    }
    try {
        New-Item -ItemType Directory -Path $receiptsRoot -Force | Out-Null
        Write-OK "Created folder: $receiptsRoot"
    } catch {
        Write-Fail "Could not create $receiptsRoot -- $_"
        exit 1
    }
}

# ---------------------------------------------------------------------------
# 3.5 Create WORKBOOKS_ROOT if it does not exist and differs from RECEIPTS_ROOT
# ---------------------------------------------------------------------------

if ($workbooksRoot -ne $receiptsRoot) {
    Write-Host ""
    Write-Step "Checking workbooks root folder..."

    if (Test-Path $workbooksRoot) {
        Write-OK "Folder exists: $workbooksRoot"
    } else {
        $confirm = Read-Host "  Folder does not exist. Create '$workbooksRoot'? [Y/n]"
        if ($confirm -ne '' -and $confirm -notmatch '^[Yy]') {
            Write-Fail "Folder not created. Aborting."
            exit 1
        }
        try {
            New-Item -ItemType Directory -Path $workbooksRoot -Force | Out-Null
            Write-OK "Created folder: $workbooksRoot"
        } catch {
            Write-Fail "Could not create $workbooksRoot -- $_"
            exit 1
        }
    }
}

# ---------------------------------------------------------------------------
# 4. Copy Accounts.template.xlsx -> Config\Accounts.xlsx
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Setting up Accounts.xlsx..."

$templateXlsx = Join-Path $repoRoot "Config\Accounts.template.xlsx"
$accountsXlsx = Join-Path $repoRoot "Config\Accounts.xlsx"

if (Test-Path $accountsXlsx) {
    Write-Skip "Accounts.xlsx already exists -- edit it directly to update your accounts"
} else {
    try {
        Copy-Item $templateXlsx $accountsXlsx
        Write-OK "Copied Accounts.template.xlsx -> Config\Accounts.xlsx"
        Write-Host ""
        Write-Host "  Open Config\Accounts.xlsx and replace the" -ForegroundColor Yellow
        Write-Host "  example rows with your own accounts before running the script." -ForegroundColor Yellow
    } catch {
        Write-Fail "Could not copy Accounts.template.xlsx -- $_"
        exit 1
    }
}

# ---------------------------------------------------------------------------
# 5. Create shortcuts in RECEIPTS_ROOT
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Creating shortcuts in $receiptsRoot..."

$shortcuts = @(
    @{
        Name        = "Run Sync Receipts.lnk"
        Target      = Join-Path $launchersDir "Run-SyncReceipts.bat"
        Description = "Sync the current month into the receipts workbook"
    },
    @{
        Name        = "Run Sync Month Receipts.lnk"
        Target      = Join-Path $launchersDir "Run-SyncMonthReceipts.bat"
        Description = "Sync a specific month into the receipts workbook"
    },
    @{
        Name        = "Run Sync Year Receipts.lnk"
        Target      = Join-Path $launchersDir "Run-SyncYearReceipts.bat"
        Description = "Sync all months for a specific year into the receipts workbook"
    },
    @{
        Name        = "Run Sync All Receipts.lnk"
        Target      = Join-Path $launchersDir "Run-SyncAllReceipts.bat"
        Description = "Sync all month folders across all years"
    }
)

$wsh = New-Object -ComObject WScript.Shell
foreach ($s in $shortcuts) {
    try {
        $lnkPath          = Join-Path $receiptsRoot $s.Name
        $verb             = if (Test-Path $lnkPath) { "Updated" } else { "Created" }
        $shortcut         = $wsh.CreateShortcut($lnkPath)
        $shortcut.TargetPath       = $s.Target
        $shortcut.WorkingDirectory = $launchersDir
        $shortcut.Description      = $s.Description
        $shortcut.Save()
        Write-OK "${verb} shortcut: $($s.Name)"
    } catch {
        Write-Fail "Could not create shortcut '$($s.Name)' -- $_"
    }
}

# ---------------------------------------------------------------------------
# 6. Install local git hooks
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Installing git hooks..."

$installHooks = Join-Path $scriptDir "Install-GitHooks.ps1"
if (Test-Path $installHooks) {
    try {
        & $installHooks
    } catch {
        Write-Fail "Could not install git hooks -- $_"
    }
} else {
    Write-Skip "Install-GitHooks.ps1 not found -- skipping"
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "Setup complete." -ForegroundColor Cyan
Write-Host ""
