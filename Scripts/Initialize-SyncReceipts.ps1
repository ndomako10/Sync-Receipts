<#
.SYNOPSIS
    One-time setup script for Sync-Receipts on a new machine.

.DESCRIPTION
    Performs all first-time configuration steps needed before running Sync-Receipts:

      1. Checks that PowerShell 5.0+ and Microsoft Excel are installed.
      2. Reads RECEIPTS_ROOT from Config\Config.ini, or prompts for the path and creates
         Config\Config.ini from Config\Templates\Config.template.ini.
      3. Creates the RECEIPTS_ROOT folder if it does not already exist.
      4. Creates the WORKBOOKS_ROOT folder if absent and different from RECEIPTS_ROOT.
      5. Copies Accounts.template.xlsx to Config\Accounts.xlsx (skipped if present).
      6. Copies Categories.template.json to Config\Categories.json (skipped if present).
      7. Copies Methods.template.json to Config\Methods.json (skipped if present).
      8. Creates Windows shortcut (.lnk) files in WORKBOOKS_ROOT that point to the batch
         launchers in the script directory, with WorkingDirectory set so UNC-path shortcuts
         open without the "UNC paths are not supported" CMD error.
      9. Installs Pester and PSScriptAnalyzer (required by the pre-commit and pre-push hooks).
     10. Installs local git hooks from Scripts\hooks\ into .git\hooks\.

    Run via Setup.bat (recommended), or directly:
        PowerShell -NoProfile -ExecutionPolicy Bypass -File Initialize-SyncReceipts.ps1

    The script is safe to re-run -- steps that are already complete are skipped.

.EXAMPLE
    PowerShell -NoProfile -ExecutionPolicy Bypass -File Scripts\Initialize-SyncReceipts.ps1

    Runs the setup script from the repo root. Equivalent to running Setup.bat.

.NOTES
    Requires: PowerShell 5.0+, Microsoft Excel (COM automation)
    Tested on: Windows 10/11
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

function Copy-ConfigTemplate {
    <#
    .SYNOPSIS
        Copies a config template file to its destination if the destination is absent.

    .DESCRIPTION
        If the destination file does not exist, copies the source file and writes
        an OK confirmation. If the destination already exists, writes a SKIP notice.
        On copy failure, writes a FAIL message and exits the script with code 1.

    .PARAMETER Source
        Full path to the template file to copy from.

    .PARAMETER Destination
        Full path to write the file to.

    .PARAMETER Label
        Human-readable name for the file, used in log messages.

    .OUTPUTS
        [bool] $true if the file was copied; $false if it was skipped.

    .EXAMPLE
        Copy-ConfigTemplate -Source $templateXlsx -Destination $accountsXlsx -Label 'Accounts.xlsx'
    #>
    param(
        [string]$Source,
        [string]$Destination,
        [string]$Label
    )
    if (Test-Path $Destination) {
        Write-Skip "$Label already exists -- edit it directly to update your settings"
        return $false
    }
    try {
        Copy-Item $Source $Destination
        Write-OK "Copied $Label"
        return $true
    } catch {
        Write-Fail "Could not copy $Label -- $_"
        exit 1
    }
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
# 1.5 Migration: move template files from Config\ to Config\Templates\ (pre-v4.0.0)
# ---------------------------------------------------------------------------

$templatesDir = Join-Path $repoRoot "Config\Templates"
$templateMoves = @(
    @{ From = "Config\Config.template.env"; To = "Config\Templates\Config.template.ini" }
    @{ From = "Config\Config.template.ini"; To = "Config\Templates\Config.template.ini" }
    @{ From = "Config\Accounts.template.xlsx";    To = "Config\Templates\Accounts.template.xlsx" }
    @{ From = "Config\Categories.template.json";  To = "Config\Templates\Categories.template.json" }
    @{ From = "Config\Methods.template.json";     To = "Config\Templates\Methods.template.json" }
)
foreach ($m in $templateMoves) {
    $src = Join-Path $repoRoot $m.From
    $dst = Join-Path $repoRoot $m.To
    if ((Test-Path $src) -and -not (Test-Path $dst)) {
        try {
            if (-not (Test-Path $templatesDir)) { New-Item -ItemType Directory -Path $templatesDir | Out-Null }
            Move-Item -Path $src -Destination $dst -ErrorAction Stop
            Write-OK "Migrated $($m.From) -> $($m.To)"
        } catch {
            Write-Fail "Could not migrate $($m.From) -- $_"
        }
    }
}

# ---------------------------------------------------------------------------
# 2. Resolve RECEIPTS_ROOT -- read Config.ini or prompt
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Configuring RECEIPTS_ROOT..."

$configPath = Join-Path $repoRoot "Config\Config.ini"
$receiptsRoot  = $null
$workbooksRoot = $null

# Migration: rename legacy Config.env -> Config.ini (pre-v4.0.0 upgrade path)
$legacyConfigPath = Join-Path $repoRoot "Config\Config.env"
if ((Test-Path $legacyConfigPath) -and -not (Test-Path $configPath)) {
    Write-Host ""
    Write-Host "  Config\Config.env found (pre-v4.0.0 name)." -ForegroundColor Yellow
    Write-Host "  From v4.0.0, the config file is named Config.ini." -ForegroundColor Yellow
    $r = Read-Host "  Rename it now? [Y/N]"
    if ($r -match '^[Yy]') {
        try {
            Rename-Item -Path $legacyConfigPath -NewName "Config.ini" -ErrorAction Stop
            Write-OK "Renamed Config\Config.env -> Config\Config.ini"
        } catch {
            Write-Fail "Could not rename Config\Config.env -- $_"
        }
    } else {
        Write-Host "  Skipped. Rename manually before running the launchers." -ForegroundColor Yellow
    }
}

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
        Write-Fail "Config.ini exists but RECEIPTS_ROOT could not be parsed."
        exit 1
    }
    if (-not $workbooksRoot) { $workbooksRoot = $receiptsRoot }
    Write-OK "Config.ini found -- RECEIPTS_ROOT: $receiptsRoot"
} else {
    Write-Host ""
    Write-Host "  Config.ini not found. Enter the path to your receipts root folder." -ForegroundColor Yellow
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
        $template = Get-Content (Join-Path $repoRoot "Config\Templates\Config.template.ini") -Raw
        $config   = $template `
            -replace 'RECEIPTS_ROOT=.*',       "RECEIPTS_ROOT=$receiptsRoot" `
            -replace 'RECEIPTS_ROOT_LOCAL=.*', "RECEIPTS_ROOT_LOCAL=$receiptsRoot" `
            -replace 'WORKBOOKS_ROOT=.*',      "WORKBOOKS_ROOT=$workbooksRoot"
        Set-Content $configPath $config -Encoding ASCII
        Write-OK "Created Config.ini (set RECEIPTS_ROOT=$receiptsRoot)"
        if ($workbooksRoot -ne $receiptsRoot) {
            Write-OK "WORKBOOKS_ROOT: $workbooksRoot"
        }
        Write-Host ""
        Write-Host "  NOTE: If your receipts root has a local drive-letter equivalent, edit" -ForegroundColor Yellow
        Write-Host "  Config.ini and update RECEIPTS_ROOT_LOCAL. This is only used by tests." -ForegroundColor Yellow
    } catch {
        Write-Fail "Could not create Config.ini -- $_"
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

$templateXlsx = Join-Path $repoRoot "Config\Templates\Accounts.template.xlsx"
$accountsXlsx = Join-Path $repoRoot "Config\Accounts.xlsx"

if (Copy-ConfigTemplate -Source $templateXlsx -Destination $accountsXlsx -Label 'Accounts.xlsx') {
    Write-Host ""
    Write-Host "  Open Config\Accounts.xlsx and replace the" -ForegroundColor Yellow
    Write-Host "  example rows with your own accounts before running the script." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# 4.5 Copy Categories.template.json -> Config\Categories.json
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Setting up Categories.json..."

$templateJson  = Join-Path $repoRoot "Config\Templates\Categories.template.json"
$categoriesJson = Join-Path $repoRoot "Config\Categories.json"

if (Copy-ConfigTemplate -Source $templateJson -Destination $categoriesJson -Label 'Categories.json') {
    Write-Host ""
    Write-Host "  Edit Config\Categories.json to add or remove categories" -ForegroundColor Yellow
    Write-Host "  before running the script." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# 4.6 Copy Methods.template.json -> Config\Methods.json
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Setting up Methods.json..."

$templateMethods = Join-Path $repoRoot "Config\Templates\Methods.template.json"
$methodsJson     = Join-Path $repoRoot "Config\Methods.json"

if (Copy-ConfigTemplate -Source $templateMethods -Destination $methodsJson -Label 'Methods.json') {
    Write-Host ""
    Write-Host "  Edit Config\Methods.json to add or remove payment method tokens" -ForegroundColor Yellow
    Write-Host "  before running the script. Cash is always valid and must not be added." -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# 5. Create shortcuts in WORKBOOKS_ROOT
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Creating shortcuts in $workbooksRoot..."

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
        $lnkPath          = Join-Path $workbooksRoot $s.Name
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
# 7. Install Pester and PSScriptAnalyzer
# ---------------------------------------------------------------------------

Write-Host ""
Write-Step "Checking test dependencies (Pester, PSScriptAnalyzer)..."

foreach ($module in @(
    @{ Name = 'Pester';            MinVersion = [version]'5.0'; InstallArgs = @('-MinimumVersion', '5.0') },
    @{ Name = 'PSScriptAnalyzer'; MinVersion = [version]'1.0'; InstallArgs = @() }
)) {
    $found = Get-Module -ListAvailable $module.Name -ErrorAction SilentlyContinue |
        Where-Object { $_.Version -ge $module.MinVersion } |
        Select-Object -First 1
    if ($found) {
        Write-Skip "$($module.Name) $($found.Version) already installed"
    } else {
        try {
            Write-Host "  Installing $($module.Name)..." -ForegroundColor Cyan
            $iArgs = $module.InstallArgs
            Install-Module $module.Name @iArgs -Scope CurrentUser -Force -ErrorAction Stop
            Write-OK "Installed $($module.Name)"
        } catch {
            Write-Fail "Could not install $($module.Name) -- $_"
        }
    }
}

# ---------------------------------------------------------------------------
# 8. Install local git hooks
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
