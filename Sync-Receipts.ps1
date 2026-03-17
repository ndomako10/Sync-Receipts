# Sync-Receipts.ps1  v0.5.0

<#
.SYNOPSIS
    Syncs receipt filenames into a formatted Excel workbook, one sheet per month.

.DESCRIPTION
    Parses every receipt file in a month folder and writes a formatted table into a
    per-year Excel workbook (e.g. 2026.xlsx) stored in ReceiptsRoot. Each year gets
    its own workbook, created automatically on first sync. Each month gets its own
    sheet named in YYMM format (e.g. 2603 for March 2026).

    Receipt filenames encode all metadata:

        YYMMDD Vendor $Amount Method [Account].ext

    The Account field accepts a 4-digit last-4, or one of two placeholder tokens:
        xxxx  - last 4 is known but intentionally omitted for privacy
        ----  - last 4 is unknown (e.g. receipt did not print it)

    Cash receipts carry no account number.

    On each run the script:
      - Reads category/subcategory definitions from Categories.json in the script folder
      - Reads valid account numbers from Accounts.xlsx in ReceiptsRoot
      - Creates or overwrites the month sheet(s) in the year workbook
      - Preserves any Category and Subcategory values previously entered by the user
      - Sorts month sheet tabs into chronological order
      - Injects dropdown validation for Category and Subcategory columns via XML patch

    Normally invoked via Run-SyncReceipts.bat or Run-SyncAllReceipts.bat, which read
    RECEIPTS_ROOT from config.bat and pass it as -ReceiptsRoot.

.PARAMETER ReceiptsRoot
    Path to the folder containing year subfolders and per-year workbooks.
    Must be provided explicitly; the script's own folder is not the data root.
    Typically set via RECEIPTS_ROOT in config.bat.

.PARAMETER YearMonth
    YYMM string identifying the month to sync (e.g. "2603" for March 2026).
    Defaults to the current month. Mutually exclusive with -Year and -All.

.PARAMETER Year
    4-digit year string (e.g. "2026"). Syncs all month folders found under
    ReceiptsRoot\{Year}\ into a single year workbook. Mutually exclusive with
    -YearMonth and -All.

.PARAMETER WorkbookPath
    Full path to a specific .xlsx file to write into. Overrides the default
    per-year workbook path ({year}.xlsx in ReceiptsRoot). Intended for testing.

.PARAMETER All
    Syncs every month folder found under every year folder in ReceiptsRoot.
    Each year is written to its own workbook. Mutually exclusive with -YearMonth
    and -Year.

.PARAMETER KillExcel
    Force-terminates any running EXCEL.EXE processes before starting.
    Use when a previous run crashed and left Excel holding the file locked.

.EXAMPLE
    .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts"

    Syncs the current month into \\Server\Share\Receipts\{year}.xlsx.

.EXAMPLE
    .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -YearMonth "2603"

    Syncs March 2026 into \\Server\Share\Receipts\2026.xlsx.

.EXAMPLE
    .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -Year 2026

    Syncs all month folders under \\Server\Share\Receipts\2026\ into 2026.xlsx.

.EXAMPLE
    .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -All

    Syncs every month folder across all years, writing each year to its own workbook.

.EXAMPLE
    .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -KillExcel

    Terminates any running Excel processes, then syncs the current month.

.NOTES
    Version : 0.5.0
    Requires: Windows PowerShell 5.0+, Microsoft Excel 2016+
    License : GNU General Public License v3.0
#>


param (
    [string]$ReceiptsRoot = $PSScriptRoot,
    [string]$YearMonth    = (Get-Date -Format "yyMM"),
    [string]$Year         = "",
    [string]$WorkbookPath = "",
    [switch]$All,
    [switch]$KillExcel
)

# ---------------------------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------------------------

function Write-Log {
<#
.SYNOPSIS
    Writes a timestamped, tagged log line to the console.

.DESCRIPTION
    All script output passes through Write-Log so that every line carries a
    consistent [HH:mm:ss] timestamp and a fixed-width type tag. The tag
    controls both the colour and the output stream:

        [STEP ] Cyan       -- major phase boundary (opening workbook, syncing month, XML patch)
        [INFO ] White      -- normal operational output (counts, saved, done)
        [WARN ] Yellow     -- non-fatal warning; script continues
        [ERROR] Red        -- error; script may abort after this
        [VERB ] (hidden)   -- verbose diagnostic detail routed to Write-Verbose;
                              shown only when the script is invoked with -Verbose

    Output format:
        [10:23:45] [INFO ] Files: 18 receipt(s) found

.PARAMETER Message
    The log message text. Should not include a leading tag or timestamp --
    those are added automatically.

.PARAMETER Tag
    Message type. One of: INFO, STEP, WARN, ERROR, VERB. Defaults to INFO.

.EXAMPLE
    Write-Log "Workbook: $wbPath" -Tag STEP

.EXAMPLE
    Write-Log "Accounts: $($accounts.Count) account(s) loaded" -Tag INFO

.EXAMPLE
    Write-Log "Sheet '$name': could not clear -- $_" -Tag WARN

.EXAMPLE
    Write-Log "Named range '$n' -> $ref" -Tag VERB
#>
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO', 'STEP', 'WARN', 'ERROR', 'VERB')]
        [string]$Tag = 'INFO'
    )
    $ts   = Get-Date -Format 'HH:mm:ss'
    $line = "[$ts] [$($Tag.PadRight(5))] $Message"
    switch ($Tag) {
        'STEP'  { Write-Host $line -ForegroundColor Cyan }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'VERB'  { Write-Verbose $line }
        default { Write-Host $line }
    }
}

function Parse-Receipt {
    param([string]$Stem)
    $pattern = '^(\d{6})\s+(.+?)\s+(-?\$[\d]+\.[\d]{2})\s+(Card|Cash|Checking|Savings)(?:\s+(\d{4}|xxxx|----))?$'
    if ($Stem -match $pattern) {
        $date = $null
        try {
            $date = [datetime]::ParseExact($Matches[1], "yyMMdd", $null)
        } catch {
            Write-Log "Parse: could not parse date '$($Matches[1])' in '$Stem' -- $_" -Tag WARN
            return @{ OK=$false; Date=$null; Vendor=""; Amount=""; Method=""; Account="" }
        }
        $vendor  = $Matches[2].Trim()
        $amount  = $Matches[3] -replace '[^0-9.\-]', ''
        $method  = $Matches[4]
        $account = if ($Matches[5]) { $Matches[5] } else { "" }
        return @{ OK=$true; Date=$date; Vendor=$vendor; Amount=$amount; Method=$method; Account=$account }
    }
    return @{ OK=$false; Date=$null; Vendor=""; Amount=""; Method=""; Account="" }
}

function Get-ValidAccounts {
    param(
        [string]$ReceiptsRoot,
        [object]$Excel    = $null,
        [object]$Workbook = $null
    )
    $accounts = @()
    $xlsxPath = Join-Path $ReceiptsRoot "Accounts.xlsx"
    if (Test-Path $xlsxPath) {
        Write-Log "Accounts: reading from Accounts.xlsx" -Tag VERB
        $accWorkbook = $null
        try {
            $accWorkbook = $Excel.Workbooks.Open($xlsxPath, 0, $true)
            $accSheet = $accWorkbook.Sheets.Item(1)
            $row = 2
            while ($true) {
                $val = $null
                try { $val = $accSheet.Cells.Item($row, 1).Value2 } catch { break }
                if ($null -eq $val -or "$val" -eq "") { break }
                $acct = $val.ToString().Trim().PadLeft(4, "0")
                if ($accounts -notcontains $acct) { $accounts += $acct }
                $row++
            }
        } catch {
            Write-Log "Accounts: failed to read Accounts.xlsx -- $_" -Tag WARN
        } finally {
            if ($accWorkbook) { try { $accWorkbook.Close($false) } catch {} }
        }
        Write-Log "Accounts: $($accounts.Count) account(s) loaded from Accounts.xlsx" -Tag VERB
        return $accounts
    }
    if ($Workbook) {
        Write-Log "Accounts: Accounts.xlsx not found in '${ReceiptsRoot}' -- falling back to Account sheet (deprecated)" -Tag WARN
        $accSheet = $null
        try { $accSheet = $Workbook.Sheets.Item("Account") } catch {}
        if (-not $accSheet) {
            Write-Log "Accounts: Account sheet not found -- account validation skipped" -Tag WARN
            return $accounts
        }
        Write-Log "Accounts: reading from Account sheet (deprecated)" -Tag VERB
        $row = 2
        while ($true) {
            $val = $null
            try { $val = $accSheet.Cells.Item($row, 1).Value2 } catch { break }
            if ($null -eq $val -or "$val" -eq "") { break }
            $acct = $val.ToString().Trim().PadLeft(4, "0")
            if ($accounts -notcontains $acct) { $accounts += $acct }
            $row++
        }
        Write-Log "Accounts: $($accounts.Count) account(s) loaded from Account sheet" -Tag VERB
    } else {
        Write-Log "Accounts: Accounts.xlsx not found in '${ReceiptsRoot}' -- account validation skipped" -Tag WARN
    }
    return $accounts
}

function Get-Categories {
    param([string]$ReceiptsRoot = $PSScriptRoot)
    $jsonPath = Join-Path $ReceiptsRoot "Categories.json"
    if (-not (Test-Path $jsonPath)) {
        Write-Log "Categories: Categories.json not found in '$ReceiptsRoot' -- dropdowns skipped" -Tag WARN
        return $null
    }
    Write-Log "Categories: reading from $jsonPath" -Tag VERB
    try {
        $json = Get-Content $jsonPath -Raw | ConvertFrom-Json
        $categories = [ordered]@{}
        $json.PSObject.Properties | ForEach-Object {
            $categories[$_.Name] = @($_.Value)
        }
        Write-Log "Categories: $($categories.Count) group(s) loaded" -Tag VERB
        return $categories
    } catch {
        Write-Log "Categories: failed to parse Categories.json -- $_" -Tag WARN
        return $null
    }
}

function Read-PreservedCategoryValues {
    param([array]$SheetData)
    $preserved = @{}
    if ($SheetData.Count -lt 2) { return $preserved }
    $headers = $SheetData[0]
    $colFN = $null; $colCat = $null; $colSub = $null
    for ($c = 0; $c -lt $headers.Count; $c++) {
        switch ($headers[$c]) {
            "File Name"   { $colFN  = $c }
            "Category"    { $colCat = $c }
            "Subcategory" { $colSub = $c }
        }
    }
    if ($null -eq $colFN -or ($null -eq $colCat -and $null -eq $colSub)) { return $preserved }
    for ($r = 1; $r -lt $SheetData.Count; $r++) {
        $row = $SheetData[$r]
        $fn  = if ($colFN  -lt $row.Count) { "$($row[$colFN])".Trim()  } else { "" }
        $cat = if ($null -ne $colCat -and $colCat -lt $row.Count) { "$($row[$colCat])".Trim() } else { "" }
        $sub = if ($null -ne $colSub -and $colSub -lt $row.Count) { "$($row[$colSub])".Trim() } else { "" }
        if ($fn -ne "" -and ($cat -ne "" -or $sub -ne "")) {
            $preserved[$fn] = @{ Category = $cat; Subcategory = $sub }
        }
    }
    return $preserved
}

function Sync-CategorySheet {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [object]$Workbook,
        [object]$Categories
    )
    Write-Log "Category sheet: syncing" -Tag VERB
    $catSheet = $null
    try { $catSheet = $Workbook.Sheets.Item("Category") } catch {}
    if (-not $catSheet) {
        try {
            $catSheet = $Workbook.Sheets.Add(
                [System.Reflection.Missing]::Value,
                $Workbook.Sheets.Item($Workbook.Sheets.Count)
            )
            $catSheet.Name = "Category"
            Write-Log "Category sheet: created" -Tag INFO
        } catch {
            Write-Log "Category sheet: could not create -- $_" -Tag ERROR
            return $null
        }
    } else {
        try {
            $catSheet.Cells.Clear()
            Write-Log "Category sheet: cleared existing content" -Tag VERB
        } catch {
            Write-Log "Category sheet: could not clear -- $_" -Tag WARN
        }
    }
    if ($PSCmdlet.ShouldProcess("Category sheet", "Sync")) {
        $col = 1
        foreach ($catName in $Categories.Keys) {
            $catSheet.Cells.Item(1, $col).Value2 = $catName
            $subcats = $Categories[$catName]
            for ($r = 0; $r -lt $subcats.Count; $r++) {
                $catSheet.Cells.Item($r + 2, $col).Value2 = $subcats[$r]
            }
            $col++
        }
        try {
            $catSheet.Visible = 0
            Write-Log "Category sheet: hidden" -Tag VERB
        } catch {
            Write-Log "Category sheet: could not hide (may be the only visible sheet) -- $_" -Tag WARN
        }
        Write-Log "Category sheet: $($Categories.Count) categories written" -Tag INFO
    }
    return $catSheet
}

function Get-ExcelColumnLetter {
    param([int]$Col)
    $result = ''
    while ($Col -gt 0) {
        $r      = ($Col - 1) % 26
        $result = [char](65 + $r) + $result
        $Col    = [math]::Floor(($Col - 1) / 26)
    }
    return $result
}

function Set-CategoryNamedRanges {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [object]$Workbook,
        [object]$CatSheet,
        [object]$Categories
    )

    # Named range "Category" covers the header row (all category names).
    # This is used directly as the source for the Category column dropdown.
    #
    # Address strings are built in pure PowerShell -- accessing ranges on a
    # hidden sheet via COM hangs in headless Excel on existing workbooks.
    $catCount      = $Categories.Keys.Count
    $lastColLetter = Get-ExcelColumnLetter -Col $catCount
    $headerRef     = "Category!" + '$A$1:$' + $lastColLetter + '$1'
    Write-Log "Named ranges: setting $catCount category named range(s)" -Tag VERB

    if ($PSCmdlet.ShouldProcess("named range 'Category' in workbook", "Set")) {
        try {
            $existing = $Workbook.Names.Item("Category")
            if ($existing) { $existing.Delete() }
        } catch {}
        try {
            $Workbook.Names.Add("Category", "=$headerRef") | Out-Null
            Write-Log "Named range 'Category' -> $headerRef" -Tag VERB
        } catch {
            Write-Log "Named range 'Category': could not add -- $_" -Tag ERROR
        }
    }

    # One named range per category, named exactly after the category (e.g. "Food", "Housing").
    # These are referenced by =INDIRECT(G2) for the Subcategory dropdown.
    $col = 1
    foreach ($catName in $Categories.Keys) {
        $subcats   = $Categories[$catName]
        $rowCount  = $subcats.Count
        $rangeName = $catName
        if ($rowCount -gt 0) {
            $colLetter = Get-ExcelColumnLetter -Col $col
            $rangeRef  = "Category!" + '$' + $colLetter + '$2:$' + $colLetter + '$' + ($rowCount + 1)
            if ($PSCmdlet.ShouldProcess("named range '$rangeName' in workbook", "Set")) {
                try {
                    $existing = $Workbook.Names.Item($rangeName)
                    if ($existing) { $existing.Delete() }
                } catch {}
                try {
                    $Workbook.Names.Add($rangeName, "=$rangeRef") | Out-Null
                    Write-Log "Named range '$rangeName' -> $rangeRef" -Tag VERB
                } catch {
                    Write-Log "Named range '$rangeName': could not add -- $_" -Tag ERROR
                }
            }
        } else {
            Write-Log "Named range '$rangeName': skipped (no subcategories)" -Tag WARN
        }
        $col++
    }
}


function Set-SubcategoryValidationXml {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string]$WorkbookPath,
        [array]$SyncResults
    )
    $xlsxName = [System.IO.Path]::GetFileName($WorkbookPath)
    Write-Log "XML patch: starting $xlsxName" -Tag STEP
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tempPath = $WorkbookPath + ".tmp"
    try {
        # Step 1: Read all zip entries into memory
        $entries = [System.Collections.Generic.Dictionary[string,byte[]]]::new()
        $srcZip = [System.IO.Compression.ZipFile]::OpenRead($WorkbookPath)
        foreach ($e in $srcZip.Entries) {
            $ms = New-Object System.IO.MemoryStream
            $e.Open().CopyTo($ms)
            $entries[$e.FullName] = $ms.ToArray()
        }
        $srcZip.Dispose()
        Write-Log "XML patch: read $($entries.Count) zip entries" -Tag VERB

        # Step 2: Resolve sheet name -> zip path via workbook.xml.rels (XML DOM, not regex)
        $wbDoc = [System.Xml.XmlDocument]::new()
        $wbDoc.LoadXml([System.Text.Encoding]::UTF8.GetString($entries['xl/workbook.xml']))
        $nsWb = [System.Xml.XmlNamespaceManager]::new($wbDoc.NameTable)
        $nsWb.AddNamespace("main", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $nsWb.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        $relsDoc = [System.Xml.XmlDocument]::new()
        $relsDoc.LoadXml([System.Text.Encoding]::UTF8.GetString($entries['xl/_rels/workbook.xml.rels']))
        $nsRels = [System.Xml.XmlNamespaceManager]::new($relsDoc.NameTable)
        $nsRels.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships")

        foreach ($sr in $SyncResults) {
            if (-not $sr.Result -or $sr.Result.DataEndRow -lt 2) {
                Write-Log "XML patch: skipping '$($sr.SheetName)' (no data)" -Tag VERB
                continue
            }
            $lastRow   = $sr.Result.DataEndRow
            $sheetName = $sr.SheetName

            $sheetNode = $wbDoc.SelectSingleNode("//main:sheet[@name='$sheetName']", $nsWb)
            if (-not $sheetNode) {
                Write-Log "XML patch: sheet '$sheetName' not found in workbook.xml" -Tag WARN
                continue
            }
            $rId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

            $relNode = $relsDoc.SelectSingleNode("//rel:Relationship[@Id='$rId']", $nsRels)
            if (-not $relNode) {
                Write-Log "XML patch: rId '$rId' not found in workbook.xml.rels" -Tag WARN
                continue
            }
            $zipPath = "xl/" + $relNode.GetAttribute("Target")

            if (-not $entries.ContainsKey($zipPath)) {
                Write-Log "XML patch: entry '$zipPath' not in zip" -Tag WARN
                continue
            }
            Write-Log "XML patch: patching '$sheetName' -> $zipPath (rows 2..$lastRow)" -Tag VERB

            # Step 3: Patch sheet XML -- remove existing dataValidations, inject both
            $sheetXml = [System.Text.Encoding]::UTF8.GetString($entries[$zipPath])
            $sheetXml = [System.Text.RegularExpressions.Regex]::Replace(
                $sheetXml,
                '<dataValidations[^>]*>.*?</dataValidations>',
                '',
                [System.Text.RegularExpressions.RegexOptions]::Singleline
            )
            $guid1 = [System.Guid]::NewGuid().ToString().ToUpper()
            $guid2 = [System.Guid]::NewGuid().ToString().ToUpper()
            $dvXml = '<dataValidations count="2">' +
                '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"' +
                " sqref=`"G2:G$lastRow`" xr:uid=`"{$guid1}`">" +
                '<formula1>Category</formula1></dataValidation>' +
                '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"' +
                " sqref=`"H2:H$lastRow`" xr:uid=`"{$guid2}`">" +
                '<formula1>INDIRECT(G2)</formula1></dataValidation>' +
                '</dataValidations>'
            if ($sheetXml -match '<hyperlinks') {
                $sheetXml = $sheetXml -replace '<hyperlinks', "$dvXml<hyperlinks"
            } elseif ($sheetXml -match '<tableParts') {
                $sheetXml = $sheetXml -replace '<tableParts', "$dvXml<tableParts"
            } else {
                $sheetXml = $sheetXml.Replace('</worksheet>', "$dvXml</worksheet>")
            }
            $entries[$zipPath] = [System.Text.Encoding]::UTF8.GetBytes($sheetXml)
            Write-Log "XML patch: '$sheetName' patched" -Tag VERB
        }

        # Step 4: Write new zip
        if (Test-Path $tempPath) { Remove-Item $tempPath -Force }
        $dstZip = [System.IO.Compression.ZipFile]::Open($tempPath, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($kvp in $entries.GetEnumerator()) {
            $e2 = $dstZip.CreateEntry($kvp.Key, [System.IO.Compression.CompressionLevel]::Optimal)
            $s  = $e2.Open()
            $s.Write($kvp.Value, 0, $kvp.Value.Length)
            $s.Close()
        }
        $dstZip.Dispose()
        Write-Log "XML patch: new zip written ($($entries.Count) entries)" -Tag VERB

        # Step 5: Binary-patch zip headers
        # Excel requires flag_bits=0x0006 in local file headers and
        # version_made_by=45 + flag_bits=0x0006 in central directory headers.
        # .NET ZipArchive writes 0x0000 and 20 respectively, which Excel rejects.
        $bytes = [System.IO.File]::ReadAllBytes($tempPath)
        $patchCount = 0
        for ($i = 0; $i -lt $bytes.Length - 4; $i++) {
            if ($bytes[$i] -eq 0x50 -and $bytes[$i+1] -eq 0x4B) {
                if ($bytes[$i+2] -eq 0x03 -and $bytes[$i+3] -eq 0x04) {
                    $bytes[$i+6] = 0x06; $bytes[$i+7] = 0x00
                    $patchCount++
                }
                if ($bytes[$i+2] -eq 0x01 -and $bytes[$i+3] -eq 0x02) {
                    $bytes[$i+4] = 0x2D; $bytes[$i+5] = 0x00
                    $bytes[$i+8] = 0x06; $bytes[$i+9] = 0x00
                    $patchCount++
                }
            }
        }
        Write-Log "XML patch: patched $patchCount zip header(s)" -Tag VERB

        if ($bytes.Length -lt 1024) {
            Write-Log "XML patch: temp file too small ($($bytes.Length) bytes) -- aborting to avoid corrupting workbook" -Tag ERROR
            Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            return
        }
        [System.IO.File]::WriteAllBytes($tempPath, $bytes)

        # Step 6: Replace original
        if ($PSCmdlet.ShouldProcess($WorkbookPath, "Overwrite with patched copy")) {
            Copy-Item $tempPath $WorkbookPath -Force
            Write-Log "XML patch: done -- $xlsxName updated" -Tag INFO
        } else {
            Write-Log "XML patch: skipped (WhatIf)" -Tag WARN
        }
        Remove-Item $tempPath -Force

    } catch {
        Write-Log "XML patch: ERROR -- $_" -Tag ERROR
        if (Test-Path $tempPath) { Remove-Item $tempPath -Force -ErrorAction SilentlyContinue }
    }
}


function Sync-Month {
    param(
        [string]$FolderPath,
        [string]$SheetName,
        [object]$Workbook,
        [array]$ValidAccounts
    )

    $receipts = @()
    try {
        Get-ChildItem -Path $FolderPath -File -ErrorAction Stop | Sort-Object Name | ForEach-Object {
            $parsed = Parse-Receipt -Stem $_.BaseName
            $receipts += [PSCustomObject]@{
                FilePath = $_.FullName
                FileName = $_.Name
                Date     = $parsed.Date
                Vendor   = $parsed.Vendor
                Amount   = $parsed.Amount
                Method   = $parsed.Method
                Account  = $parsed.Account
                ParseOK  = $parsed.OK
            }
        }
    } catch {
        Write-Log "Files: could not read folder '$FolderPath' -- $_" -Tag ERROR
        return $null
    }
    Write-Log "Files: $($receipts.Count) receipt(s) found" -Tag INFO

    $sheet = $null
    try { $sheet = $Workbook.Sheets.Item($SheetName) } catch {}
    if (-not $sheet) {
        try {
            $sheet      = $Workbook.Sheets.Add()
            $sheet.Name = $SheetName
            Write-Log "Sheet '$SheetName': created" -Tag INFO
        } catch {
            Write-Log "Sheet '$SheetName': could not create -- $_" -Tag ERROR
            return $null
        }
    } else {
        Write-Log "Sheet '$SheetName': found (existing)" -Tag VERB
    }

    # A=File Name  B=Date  C=Vendor  D=Amount  E=Method  F=Account  G=Category  H=Subcategory  I=Flag
    $COL_FILENAME    = 1
    $COL_DATE        = 2
    $COL_VENDOR      = 3
    $COL_AMOUNT      = 4
    $COL_METHOD      = 5
    $COL_ACCOUNT     = 6
    $COL_CATEGORY    = 7
    $COL_SUBCATEGORY = 8
    $COL_FLAG        = 9
    $NUM_COLS        = 9

    # Preserve existing Category/Subcategory values keyed by filename.
    # Raw cell values are extracted into a plain 2D array so the logic is
    # handled by Read-PreservedCategoryValues (testable without COM).
    $preserved = @{}
    try {
        $existingRows = $sheet.UsedRange.Rows.Count
        $existingCols = $sheet.UsedRange.Columns.Count
        if ($existingRows -gt 1 -and $existingCols -ge 1) {
            $sheetData = @()
            for ($r = 1; $r -le $existingRows; $r++) {
                $rowData = @()
                for ($c = 1; $c -le $existingCols; $c++) {
                    $rowData += "$($sheet.Cells.Item($r, $c).Value2)".Trim()
                }
                $sheetData += ,$rowData
            }
            $preserved = Read-PreservedCategoryValues -SheetData $sheetData
            if ($preserved.Count -gt 0) {
                Write-Log "Sheet '$SheetName': preserved $($preserved.Count) Category/Subcategory value(s)" -Tag INFO
            }
        }
    } catch {
        Write-Log "Sheet '$SheetName': could not read existing Category/Subcategory data -- $_" -Tag WARN
    }

    if ($sheet.ListObjects.Count -gt 0) {
        try {
            $sheet.ListObjects.Item(1).Unlist()
            Write-Log "Sheet '$SheetName': unlisted existing table" -Tag VERB
        } catch {
            Write-Log "Sheet '$SheetName': could not unlist existing table -- $_" -Tag WARN
        }
    }
    try {
        $lastRow = $sheet.UsedRange.Rows.Count
        if ($lastRow -gt 1) {
            $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($lastRow, $NUM_COLS)).Clear()
            Write-Log "Sheet '$SheetName': cleared $lastRow existing row(s)" -Tag VERB
        }
    } catch {
        Write-Log "Sheet '$SheetName': could not clear existing content -- $_" -Tag WARN
    }

    $headers = @("File Name", "Date", "Vendor", "Amount", "Method", "Account", "Category", "Subcategory", "Flag")
    for ($c = 1; $c -le $NUM_COLS; $c++) {
        $sheet.Cells.Item(1, $c).Value2 = $headers[$c-1]
    }

    $row = 2
    foreach ($r in $receipts) {

        $cell = $sheet.Cells.Item($row, $COL_FILENAME)
        try {
            $sheet.Hyperlinks.Add($cell, $r.FilePath, [System.Reflection.Missing]::Value, "Open receipt file", $r.FileName) | Out-Null
        } catch {
            Write-Log "Hyperlink: could not add for '$($r.FileName)' -- $_" -Tag WARN
            $cell.Value2 = $r.FileName
        }

        try {
            if ($r.Date) {
                $sheet.Cells.Item($row, $COL_DATE).Value2       = [double]$r.Date.ToOADate()
                $sheet.Cells.Item($row, $COL_DATE).NumberFormat = "d-mmm"
            }
            $sheet.Cells.Item($row, $COL_VENDOR).Value2        = $r.Vendor
            $sheet.Cells.Item($row, $COL_METHOD).Value2        = $r.Method
            $sheet.Cells.Item($row, $COL_ACCOUNT).NumberFormat = "@"
            $sheet.Cells.Item($row, $COL_ACCOUNT).Value2       = $r.Account
            if ($r.Amount -ne "") {
                $sheet.Cells.Item($row, $COL_AMOUNT).Value2 = [double]$r.Amount
            }
        } catch {
            Write-Log "Row ${row}: error writing data cells ('$($r.FileName)') -- $_" -Tag WARN
        }

        $flag = ""
        if (-not $r.ParseOK) {
            $flag = "Could not parse filename"
        } elseif ($r.Account -eq "xxxx") {
            $flag = "Account obfuscated"
        } elseif ($r.Account -eq "----") {
            $flag = "Account unknown"
        } elseif ($r.Account -ne "" -and $ValidAccounts.Count -gt 0 -and $ValidAccounts -notcontains $r.Account) {
            $flag = "Account not in Accounts.xlsx"
        }
        if ($flag -ne "") {
            try {
                $sheet.Cells.Item($row, $COL_FLAG).Value2 = $flag
            } catch {
                Write-Log "Row ${row}: could not write flag -- $_" -Tag WARN
            }
        }

        if ($preserved.ContainsKey($r.FileName)) {
            $saved = $preserved[$r.FileName]
            try {
                if ($saved.Category    -ne "") { $sheet.Cells.Item($row, $COL_CATEGORY).Value2    = $saved.Category }
                if ($saved.Subcategory -ne "") { $sheet.Cells.Item($row, $COL_SUBCATEGORY).Value2 = $saved.Subcategory }
            } catch {
                Write-Log "Row ${row}: could not restore Category/Subcategory for '$($r.FileName)' -- $_" -Tag WARN
            }
        }

        $row++
    }

    $dataEnd    = $row - 1
    Write-Log "Rows: header=1, data=2..$dataEnd ($(($dataEnd - 1)) row(s))" -Tag INFO
    $tableRange = $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($dataEnd, $NUM_COLS))
    $table = $null
    try {
        $table      = $sheet.ListObjects.Add(1, $tableRange, [System.Reflection.Missing]::Value, 1)
        $table.Name = "Receipts_$SheetName"
        Write-Log "Table 'Receipts_$SheetName': created over $($tableRange.Address())" -Tag VERB
    } catch {
        Write-Log "Table 'Receipts_$SheetName': could not create -- $_" -Tag WARN
    }

    if ($table) {
        try {
            $table.ListColumns.Item($COL_AMOUNT).DataBodyRange.NumberFormat = '$#,##0.00;[Red]($#,##0.00)'
        } catch {
            Write-Log "Amount format: could not set -- $_" -Tag WARN
        }
    }

    for ($r2 = 2; $r2 -le $dataEnd; $r2++) {
        $flagVal = $sheet.Cells.Item($r2, $COL_FLAG).Value2
        if ($flagVal -ne "" -and $null -ne $flagVal) {
            $sheet.Cells.Item($r2, $COL_FLAG).Font.ColorIndex = 3
            $sheet.Cells.Item($r2, $COL_FLAG).Font.Bold       = $true
        }
    }

    $sumRow    = $dataEnd + 2
    $flagCount = ($receipts | Where-Object {
        -not $_.ParseOK -or
        $_.Account -eq "xxxx" -or
        $_.Account -eq "----" -or
        ($_.Account -ne "" -and $ValidAccounts.Count -gt 0 -and $ValidAccounts -notcontains $_.Account)
    }).Count

    $amtColLetter = [char](64 + $COL_AMOUNT)
    try {
        $sheet.Cells.Item($sumRow, $COL_VENDOR).Value2       = "Total"
        $sheet.Cells.Item($sumRow, $COL_VENDOR).Font.Bold    = $true
        $sheet.Cells.Item($sumRow, $COL_AMOUNT).Formula      = "=SUM(${amtColLetter}2:${amtColLetter}$dataEnd)"
        $sheet.Cells.Item($sumRow, $COL_AMOUNT).NumberFormat = '$#,##0.00;[Red]($#,##0.00)'
        $sheet.Cells.Item($sumRow, $COL_AMOUNT).Font.Bold    = $true
        if ($flagCount -gt 0) {
            $sheet.Cells.Item($sumRow, $COL_FLAG).Value2          = "$flagCount flagged"
            $sheet.Cells.Item($sumRow, $COL_FLAG).Font.ColorIndex = 3
            $sheet.Cells.Item($sumRow, $COL_FLAG).Font.Bold       = $true
        }
        Write-Log "Total row: written at row $sumRow" -Tag VERB
    } catch {
        Write-Log "Total row: could not write -- $_" -Tag WARN
    }

    $sheet.Columns.Item($COL_FILENAME).ColumnWidth    = 55
    $sheet.Columns.Item($COL_DATE).ColumnWidth        = 12
    $sheet.Columns.Item($COL_VENDOR).ColumnWidth      = 28
    $sheet.Columns.Item($COL_AMOUNT).ColumnWidth      = 14
    $sheet.Columns.Item($COL_METHOD).ColumnWidth      = 12
    $sheet.Columns.Item($COL_ACCOUNT).ColumnWidth     = 12
    $sheet.Columns.Item($COL_CATEGORY).ColumnWidth    = 20
    $sheet.Columns.Item($COL_SUBCATEGORY).ColumnWidth = 20
    $sheet.Columns.Item($COL_FLAG).ColumnWidth        = 28

    try {
        $null = $sheet.Activate()
        $null = $sheet.Cells.Item(2, 1).Select()
        $excel.ActiveWindow.FreezePanes = $true
    } catch {
        Write-Log "Freeze panes: could not set -- $_" -Tag WARN
    }

    $parsed   = ($receipts | Where-Object { $_.ParseOK }).Count
    $unparsed = ($receipts | Where-Object { -not $_.ParseOK }).Count
    Write-Log "Written: $parsed receipt(s)" -Tag INFO
    if ($unparsed -gt 0) {
        Write-Log "Unparsed: $unparsed receipt(s) flagged in sheet" -Tag WARN
    }

    return [PSCustomObject]@{
        DataEndRow = $dataEnd
    }
}

function Set-MonthSheetOrder {
    param([object]$Workbook)
    # Collect all month sheets (4-digit YYMM names).
    $monthSheets = @()
    foreach ($s in $Workbook.Sheets) {
        if ($s.Name -match '^\d{4}$') { $monthSheets += $s }
    }
    if ($monthSheets.Count -lt 2) { return }

    $sorted  = $monthSheets | Sort-Object Name
    $missing = [System.Reflection.Missing]::Value

    # Move the earliest month sheet to position 1, then place each subsequent
    # sheet immediately after the one before it.
    $sorted[0].Move($Workbook.Sheets.Item(1))
    for ($i = 1; $i -lt $sorted.Count; $i++) {
        $sorted[$i].Move($missing, $sorted[$i - 1])
    }
    Write-Log "Sheets: $($sorted.Count) month sheet(s) sorted chronologically" -Tag INFO
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

if (-not (Test-Path $ReceiptsRoot)) {
    Write-Error "Receipts folder not found: '$ReceiptsRoot'"
    exit 1
}

# Validate mutually exclusive mode flags.
$modeCount = 0
if ($All)          { $modeCount++ }
if ($Year -ne "")  { $modeCount++ }
if ($modeCount -gt 1) {
    Write-Error "-All and -Year are mutually exclusive. Specify only one."
    exit 1
}
if ($Year -ne "" -and $YearMonth -ne (Get-Date -Format "yyMM")) {
    Write-Error "-Year and -YearMonth are mutually exclusive. Specify only one."
    exit 1
}

# Gather months grouped by year.
# Each year maps to an array of {SheetName, FolderPath} objects.
$yearGroups = @{}
if ($All) {
    Get-ChildItem -Path $ReceiptsRoot -Directory |
        Where-Object { $_.Name -match '^\d{4}$' } |
        ForEach-Object {
            $yearDir = $_
            $yearKey = $yearDir.Name
            Get-ChildItem -Path $yearDir.FullName -Directory |
                Where-Object { $_.Name -match '^\d{4}' } |
                ForEach-Object {
                    if (-not $yearGroups.ContainsKey($yearKey)) { $yearGroups[$yearKey] = @() }
                    $yearGroups[$yearKey] += [PSCustomObject]@{
                        SheetName  = $_.Name.Substring(0, 4)
                        FolderPath = $_.FullName
                    }
                }
        }
    $totalMonths = ($yearGroups.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    Write-Log "Scan: $($yearGroups.Count) year(s), $totalMonths month folder(s) found" -Tag INFO
} elseif ($Year -ne "") {
    $yearDir = Join-Path $ReceiptsRoot $Year
    if (-not (Test-Path $yearDir)) {
        Write-Error "Year folder not found: '$yearDir'"
        exit 1
    }
    Get-ChildItem -Path $yearDir -Directory |
        Where-Object { $_.Name -match '^\d{4}' } |
        ForEach-Object {
            if (-not $yearGroups.ContainsKey($Year)) { $yearGroups[$Year] = @() }
            $yearGroups[$Year] += [PSCustomObject]@{
                SheetName  = $_.Name.Substring(0, 4)
                FolderPath = $_.FullName
            }
        }
    if ($yearGroups.Count -eq 0 -or $yearGroups[$Year].Count -eq 0) {
        Write-Error "No month folders found under '$yearDir'"
        exit 1
    }
    Write-Log "Scan: $($yearGroups[$Year].Count) month folder(s) found for $Year" -Tag INFO
} else {
    $yearSingle  = "20" + $YearMonth.Substring(0, 2)
    $monthFolder = Get-ChildItem -Path (Join-Path $ReceiptsRoot $yearSingle) -Directory |
                   Where-Object { $_.Name -match "^$YearMonth" } |
                   Select-Object -First 1
    if (-not $monthFolder) {
        Write-Error "Could not find a folder matching '$YearMonth*' under '$ReceiptsRoot\$yearSingle'"
        exit 1
    }
    $yearGroups[$yearSingle] = @([PSCustomObject]@{ SheetName = $YearMonth; FolderPath = $monthFolder.FullName })
}

# Kill any lingering EXCEL.EXE processes before starting, if requested.
if ($KillExcel) {
    $procs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($procs) {
        Write-Log "KillExcel: stopping $($procs.Count) EXCEL.EXE process(es)" -Tag STEP
        $procs | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Log "KillExcel: done" -Tag INFO
    } else {
        Write-Log "KillExcel: no EXCEL.EXE processes found" -Tag INFO
    }
}

$excel                  = New-Object -ComObject Excel.Application
$excel.Visible          = $false
$excel.DisplayAlerts    = $false
$excel.AskToUpdateLinks = $false

foreach ($yearEntry in ($yearGroups.GetEnumerator() | Sort-Object Key)) {
    $year   = $yearEntry.Key
    $months = $yearEntry.Value

    # Resolve workbook path for this year.
    # -WorkbookPath overrides the derived path (used for testing).
    if ($WorkbookPath -ne "") {
        $wbPath = $WorkbookPath
    } else {
        $wbPath = Join-Path $ReceiptsRoot "$year.xlsx"
    }
    $xlsxName = [System.IO.Path]::GetFileName($wbPath)
    Write-Host ""
    Write-Log "Workbook: $wbPath" -Tag STEP

    # Check for file lock before attempting to open an existing workbook.
    if (Test-Path $wbPath) {
        try {
            $stream = [System.IO.File]::Open($wbPath, 'Open', 'Read', 'None')
            $stream.Close()
        } catch {
            $procs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
            if ($procs) {
                Write-Log "Workbook: '$xlsxName' is locked by $($procs.Count) EXCEL.EXE process(es)" -Tag ERROR
                $procs | ForEach-Object { Write-Log "  PID $($_.Id)  started $($_.StartTime)" -Tag ERROR }
                Write-Log "  Re-run with -KillExcel to terminate automatically" -Tag WARN
            } else {
                Write-Log "Workbook: '$xlsxName' is locked but no EXCEL.EXE found" -Tag ERROR
                Write-Log "  Another process or the OS may be holding the file" -Tag WARN
            }
            $excel.Quit()
            exit 1
        }
    }

    # Open the existing workbook, or create a new one if it does not exist yet.
    $workbook = $null
    if (Test-Path $wbPath) {
        Write-Log "Workbook: opening $xlsxName" -Tag VERB
        try {
            $workbook = $excel.Workbooks.Open($wbPath, 0)  # UpdateLinks=0: do not prompt to update external links
        } catch {
            $errMsg = $_.Exception.Message
            Write-Log "Workbook: Excel COM threw an exception opening '$xlsxName'" -Tag ERROR
            Write-Log "  Path   : $wbPath" -Tag ERROR
            Write-Log "  Detail : $errMsg" -Tag ERROR
            Write-Log "  Likely cause 1: file is corrupted -- restore from a backup and re-run" -Tag WARN
            Write-Log "  Likely cause 2: Excel blocked the file (network share) -- open manually and click Enable Editing" -Tag WARN
            Write-Log "  Likely cause 3: a previous Excel process is still running -- re-run with -KillExcel" -Tag WARN
            $excel.Quit()
            exit 1
        }
    } else {
        Write-Log "Workbook: creating $xlsxName (new)" -Tag STEP
        try {
            $workbook = $excel.Workbooks.Add()
            $workbook.SaveAs($wbPath)
        } catch {
            $errMsg = $_.Exception.Message
            Write-Log "Workbook: failed to create at '$wbPath' -- $errMsg" -Tag ERROR
            $excel.Quit()
            exit 1
        }
    }
    if (-not $workbook) {
        Write-Error "Failed to open workbook (returned null): $wbPath"
        $excel.Quit()
        exit 1
    }
    Write-Log "Workbook: $xlsxName opened successfully" -Tag VERB

    Write-Log "Accounts: loading" -Tag VERB
    $validAccounts = @()
    try {
        $validAccounts = Get-ValidAccounts -ReceiptsRoot $ReceiptsRoot -Excel $excel -Workbook $workbook
    } catch {
        Write-Log "Accounts: error in Get-ValidAccounts -- $_" -Tag WARN
    }

    Write-Log "Categories: loading" -Tag VERB
    $categories = $null
    try {
        $categories = Get-Categories
    } catch {
        Write-Log "Categories: error in Get-Categories -- $_" -Tag WARN
    }

    if ($categories) {
        $catSheet = $null
        try {
            $catSheet = Sync-CategorySheet -Workbook $workbook -Categories $categories
        } catch {
            Write-Log "Category sheet: error in Sync-CategorySheet -- $_" -Tag WARN
        }
        if ($catSheet) {
            try {
                Set-CategoryNamedRanges -Workbook $workbook -CatSheet $catSheet -Categories $categories
                Write-Log "Named ranges: $($categories.Count) category/subcategory group(s) configured" -Tag INFO
            } catch {
                Write-Log "Named ranges: error in Set-CategoryNamedRanges -- $_" -Tag WARN
            }
        }
    }
    Write-Log "Accounts: $($validAccounts.Count) account(s) loaded" -Tag INFO

    $syncResults = @()
    foreach ($m in $months) {
        Write-Host ""
        Write-Log "Syncing: $($m.FolderPath)" -Tag STEP
        try {
            $result = Sync-Month -FolderPath $m.FolderPath -SheetName $m.SheetName -Workbook $workbook -ValidAccounts $validAccounts
            $syncResults += [PSCustomObject]@{ SheetName = $m.SheetName; Result = $result }
        } catch {
            Write-Log "Sync error: unhandled exception syncing '$($m.SheetName)' -- $_" -Tag ERROR
            Write-Log "  Skipping this month and continuing" -Tag WARN
            $syncResults += [PSCustomObject]@{ SheetName = $m.SheetName; Result = $null }
        }
    }

    # Remove any default blank sheets (Sheet1, Sheet2, etc.) left by Workbooks.Add().
    # This is done after sync so the workbook always has at least one real sheet,
    # allowing Excel to delete the default ones without error.
    try {
        $defaultSheets = @()
        foreach ($s in $workbook.Sheets) {
            if ($s.Name -match '^Sheet\d+$') { $defaultSheets += $s }
        }
        foreach ($s in $defaultSheets) {
            try { $s.Delete() } catch {}
        }
        if ($defaultSheets.Count -gt 0) {
            Write-Log "Cleanup: removed $($defaultSheets.Count) default sheet(s)" -Tag INFO
        }
    } catch {
        Write-Log "Cleanup: error removing default sheets -- $_" -Tag WARN
    }

    try {
        Set-MonthSheetOrder -Workbook $workbook
    } catch {
        Write-Log "Sheets: error sorting month sheets -- $_" -Tag WARN
    }

    try {
        $workbook.Save()
        Write-Log "Workbook: saved $xlsxName" -Tag INFO
    } catch {
        Write-Log "Workbook: failed to save -- $_" -Tag ERROR
    }
    try {
        $workbook.Close()
        Write-Log "Workbook: closed $xlsxName" -Tag VERB
    } catch {
        Write-Log "Workbook: error closing -- $_" -Tag WARN
    }

    Set-SubcategoryValidationXml -WorkbookPath $wbPath -SyncResults $syncResults
}

try {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Log "Excel: COM instance released" -Tag VERB
} catch {
    Write-Log "Excel: error quitting -- $_" -Tag WARN
}

Write-Host ""
Write-Log "Done!" -Tag STEP
