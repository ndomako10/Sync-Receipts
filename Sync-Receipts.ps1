# Sync-Receipts.ps1  v0.5.0
#
# Syncs receipt filenames from a month folder into a table in a per-year Excel workbook.
# Each year gets its own workbook: 2026.xlsx, 2025.xlsx, etc., stored in ReceiptsRoot.
# The sheet is named in YYMM format (e.g. "2603" for March 2026).
# The table stays in sync with folder contents - rows are added or removed to match.
# Filenames that cannot be parsed are added with blank fields and flagged.
# Account numbers are validated against Accounts.xlsx in ReceiptsRoot.
# Category and Subcategory dropdowns are populated from Categories.json via a generated Category sheet.
#
# Usage:
#   .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts"
#   .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -YearMonth "2603"
#   .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -All
#   .\Sync-Receipts.ps1 -ReceiptsRoot "\\Server\Share\Receipts" -KillExcel
#
# Normally invoked via Run-SyncReceipts.bat or Run-SyncAllReceipts.bat, which read
# RECEIPTS_ROOT from config.bat and pass it as -ReceiptsRoot.
#
# Parameters:
#   -ReceiptsRoot : Path to the folder containing year subfolders and per-year workbooks.
#                   Must be provided explicitly; the script's own folder is not the data root.
#   -YearMonth    : YYMM to sync (e.g. "2603"). Defaults to current month.
#   -WorkbookPath : Full path to the .xlsx workbook to write into. Overrides the default
#                   per-year workbook path ({year}.xlsx in ReceiptsRoot). Useful for testing.
#   -All          : Sync every month folder found under every year folder.
#   -KillExcel    : Kill any running EXCEL.EXE processes before starting. Use when a
#                   previous run crashed and left Excel holding the file lock.

param (
    [string]$ReceiptsRoot = $PSScriptRoot,
    [string]$YearMonth    = (Get-Date -Format "yyMM"),
    [string]$WorkbookPath = "",
    [switch]$All,
    [switch]$KillExcel
)

# ---------------------------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------------------------

function Parse-Receipt {
    param([string]$Stem)
    $pattern = '^(\d{6})\s+(.+?)\s+(-?\$[\d]+\.[\d]{2})\s+(Card|Cash|Checking|Savings)(?:\s+(\d{4}|xxxx))?$'
    if ($Stem -match $pattern) {
        $date = $null
        try {
            $date = [datetime]::ParseExact($Matches[1], "yyMMdd", $null)
        } catch {
            Write-Host "  Warning  : Could not parse date '$($Matches[1])' in '$Stem': $_" -ForegroundColor Yellow
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
    Write-Host "Debug    : Get-ValidAccounts start"
    $accounts = @()
    $xlsxPath = Join-Path $ReceiptsRoot "Accounts.xlsx"
    if (Test-Path $xlsxPath) {
        Write-Host "Debug    : Accounts.xlsx found"
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
            Write-Host "  Warning  : Failed to read Accounts.xlsx: $_" -ForegroundColor Yellow
        } finally {
            if ($accWorkbook) { try { $accWorkbook.Close($false) } catch {} }
        }
        Write-Host "Debug    : Get-ValidAccounts end - $($accounts.Count) account(s)"
        return $accounts
    }
    if ($Workbook) {
        Write-Host "  Warning  : Accounts.xlsx not found in '${ReceiptsRoot}' - falling back to Account sheet (deprecated)" -ForegroundColor Yellow
        $accSheet = $null
        try { $accSheet = $Workbook.Sheets.Item("Account") } catch {}
        if (-not $accSheet) {
            Write-Host "  Warning  : Account sheet not found - account validation skipped" -ForegroundColor Yellow
            return $accounts
        }
        Write-Host "Debug    : Account sheet found"
        $row = 2
        while ($true) {
            $val = $null
            try { $val = $accSheet.Cells.Item($row, 1).Value2 } catch { break }
            if ($null -eq $val -or "$val" -eq "") { break }
            $acct = $val.ToString().Trim().PadLeft(4, "0")
            if ($accounts -notcontains $acct) { $accounts += $acct }
            $row++
        }
        Write-Host "Debug    : Get-ValidAccounts end - $($accounts.Count) account(s)"
    } else {
        Write-Host "  Warning  : Accounts.xlsx not found in '${ReceiptsRoot}' - account validation skipped" -ForegroundColor Yellow
    }
    return $accounts
}

function Get-Categories {
    param([string]$ReceiptsRoot = $PSScriptRoot)
    Write-Host "Debug    : Get-Categories start"
    $jsonPath = Join-Path $ReceiptsRoot "Categories.json"
    if (-not (Test-Path $jsonPath)) {
        Write-Host "  Warning  : Categories.json not found in '$ReceiptsRoot' - dropdowns skipped" -ForegroundColor Yellow
        return $null
    }
    try {
        $json = Get-Content $jsonPath -Raw | ConvertFrom-Json
        $categories = [ordered]@{}
        $json.PSObject.Properties | ForEach-Object {
            $categories[$_.Name] = @($_.Value)
        }
        Write-Host "Debug    : Get-Categories end - $($categories.Count) group(s)"
        return $categories
    } catch {
        Write-Host "  Warning  : Failed to parse Categories.json: $_" -ForegroundColor Yellow
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
    Write-Host "Debug    : Sync-CategorySheet start"
    $catSheet = $null
    try { $catSheet = $Workbook.Sheets.Item("Category") } catch {}
    if (-not $catSheet) {
        try {
            $catSheet = $Workbook.Sheets.Add(
                [System.Reflection.Missing]::Value,
                $Workbook.Sheets.Item($Workbook.Sheets.Count)
            )
            $catSheet.Name = "Category"
            Write-Host "  Sheet    : Created 'Category' sheet"
        } catch {
            Write-Host "  Error    : Could not create Category sheet: $_" -ForegroundColor Red
            return $null
        }
    } else {
        try {
            $catSheet.Cells.Clear()
            Write-Host "  Sheet    : Cleared existing 'Category' sheet"
        } catch {
            Write-Host "  Warning  : Could not clear Category sheet: $_" -ForegroundColor Yellow
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
            Write-Host "  Sheet    : 'Category' sheet hidden"
        } catch {
            Write-Host "  Warning  : Could not hide Category sheet (may be the only visible sheet): $_" -ForegroundColor Yellow
        }
        Write-Host "Debug    : Sync-CategorySheet end - $($Categories.Count) categories written"
    }
    return $catSheet
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
    $catCount  = $Categories.Keys.Count
    $headerRef = "Category!" + $CatSheet.Range(
        $CatSheet.Cells.Item(1, 1),
        $CatSheet.Cells.Item(1, $catCount)
    ).Address($true, $true, 1)
    if ($PSCmdlet.ShouldProcess("named range 'Category' in workbook", "Set")) {
        try {
            $existing = $Workbook.Names.Item("Category")
            if ($existing) { $existing.Delete() }
        } catch {}
        try {
            $Workbook.Names.Add("Category", "=$headerRef") | Out-Null
            Write-Host "  Range    : 'Category' header range -> $headerRef"
        } catch {
            Write-Host "  Range ERR: Could not add 'Category' named range: $_" -ForegroundColor Red
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
            $rangeRef = "Category!" + $CatSheet.Range(
                $CatSheet.Cells.Item(2, $col),
                $CatSheet.Cells.Item($rowCount + 1, $col)
            ).Address($true, $true, 1)
            if ($PSCmdlet.ShouldProcess("named range '$rangeName' in workbook", "Set")) {
                try {
                    $existing = $Workbook.Names.Item($rangeName)
                    if ($existing) { $existing.Delete() }
                } catch {}
                try {
                    $Workbook.Names.Add($rangeName, "=$rangeRef") | Out-Null
                    Write-Host "  Range    : '$rangeName' -> $rangeRef"
                } catch {
                    Write-Host "  Range ERR: '$rangeName' -> $_" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "  Range    : '$rangeName' skipped (no subcategories)" -ForegroundColor Yellow
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
    Write-Host "XmlPatch : Start"
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
        Write-Host "XmlPatch : Read $($entries.Count) entries"

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
                Write-Host "XmlPatch : Skipping '$($sr.SheetName)' (no data)"
                continue
            }
            $lastRow   = $sr.Result.DataEndRow
            $sheetName = $sr.SheetName

            $sheetNode = $wbDoc.SelectSingleNode("//main:sheet[@name='$sheetName']", $nsWb)
            if (-not $sheetNode) {
                Write-Host "XmlPatch : '$sheetName' not found in workbook.xml" -ForegroundColor Yellow
                continue
            }
            $rId = $sheetNode.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

            $relNode = $relsDoc.SelectSingleNode("//rel:Relationship[@Id='$rId']", $nsRels)
            if (-not $relNode) {
                Write-Host "XmlPatch : rId '$rId' not found in workbook.xml.rels" -ForegroundColor Yellow
                continue
            }
            $zipPath = "xl/" + $relNode.GetAttribute("Target")

            if (-not $entries.ContainsKey($zipPath)) {
                Write-Host "XmlPatch : Entry '$zipPath' not in zip" -ForegroundColor Yellow
                continue
            }
            Write-Host "XmlPatch : Patching '$sheetName' -> $zipPath (rows 2..$lastRow)"

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
            Write-Host "XmlPatch : '$sheetName' XML patched"
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
        Write-Host "XmlPatch : New zip written"

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
        Write-Host "XmlPatch : Patched $patchCount zip header(s)"

        if ($bytes.Length -lt 1024) {
            Write-Host "XmlPatch : ERROR - temp file too small ($($bytes.Length) bytes), aborting" -ForegroundColor Red
            Remove-Item $tempPath -Force -ErrorAction SilentlyContinue
            return
        }
        [System.IO.File]::WriteAllBytes($tempPath, $bytes)

        # Step 6: Replace original
        if ($PSCmdlet.ShouldProcess($WorkbookPath, "Overwrite with patched copy")) {
            Copy-Item $tempPath $WorkbookPath -Force
            Write-Host "XmlPatch : Done"
        } else {
            Write-Host "XmlPatch : Skipped (WhatIf)" -ForegroundColor Yellow
        }
        Remove-Item $tempPath -Force

    } catch {
        Write-Host "XmlPatch : ERROR - $_" -ForegroundColor Red
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
        Write-Host "  Error    : Could not read folder '$FolderPath': $_" -ForegroundColor Red
        return $null
    }
    Write-Host "  Files    : $($receipts.Count) receipt(s) found"

    $sheet = $null
    try { $sheet = $Workbook.Sheets.Item($SheetName) } catch {}
    if (-not $sheet) {
        try {
            $sheet      = $Workbook.Sheets.Add()
            $sheet.Name = $SheetName
            Write-Host "  Sheet    : '$SheetName' created"
        } catch {
            Write-Host "  Error    : Could not create sheet '$SheetName': $_" -ForegroundColor Red
            return $null
        }
    } else {
        Write-Host "  Sheet    : '$SheetName' found"
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
                Write-Host "  Sheet    : Preserved $($preserved.Count) Category/Subcategory value(s)"
            }
        }
    } catch {
        Write-Host "  Warning  : Could not read existing Category/Subcategory data: $_" -ForegroundColor Yellow
    }

    if ($sheet.ListObjects.Count -gt 0) {
        try {
            $sheet.ListObjects.Item(1).Unlist()
            Write-Host "  Sheet    : Existing table unlisted"
        } catch {
            Write-Host "  Warning  : Could not unlist existing table: $_" -ForegroundColor Yellow
        }
    }
    try {
        $lastRow = $sheet.UsedRange.Rows.Count
        if ($lastRow -gt 1) {
            $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($lastRow, $NUM_COLS)).Clear()
            Write-Host "  Sheet    : Cleared $lastRow existing row(s)"
        }
    } catch {
        Write-Host "  Warning  : Could not clear existing sheet content: $_" -ForegroundColor Yellow
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
            Write-Host "  Warning  : Could not add hyperlink for '$($r.FileName)': $_" -ForegroundColor Yellow
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
            Write-Host "  Warning  : Error writing data cells for row $row ('$($r.FileName)'): $_" -ForegroundColor Yellow
        }

        $flag = ""
        if (-not $r.ParseOK) {
            $flag = "Could not parse filename"
        } elseif ($r.Account -eq "0000") {
            $flag = "Unknown account number"
        } elseif ($r.Account -ne "" -and $r.Account -ne "xxxx" -and $ValidAccounts.Count -gt 0 -and $ValidAccounts -notcontains $r.Account) {
            $flag = "Account not in Accounts.xlsx"
        }
        if ($flag -ne "") {
            try {
                $sheet.Cells.Item($row, $COL_FLAG).Value2 = $flag
            } catch {
                Write-Host "  Warning  : Could not write flag for row ${row}: $_" -ForegroundColor Yellow
            }
        }

        if ($preserved.ContainsKey($r.FileName)) {
            $saved = $preserved[$r.FileName]
            try {
                if ($saved.Category    -ne "") { $sheet.Cells.Item($row, $COL_CATEGORY).Value2    = $saved.Category }
                if ($saved.Subcategory -ne "") { $sheet.Cells.Item($row, $COL_SUBCATEGORY).Value2 = $saved.Subcategory }
            } catch {
                Write-Host "  Warning  : Could not restore Category/Subcategory for '$($r.FileName)': $_" -ForegroundColor Yellow
            }
        }

        $row++
    }

    $dataEnd    = $row - 1
    Write-Host "  Rows     : header=1, data=2..$dataEnd ($(($dataEnd - 1)) row(s))"
    $tableRange = $sheet.Range($sheet.Cells.Item(1,1), $sheet.Cells.Item($dataEnd, $NUM_COLS))
    $table = $null
    try {
        $table      = $sheet.ListObjects.Add(1, $tableRange, [System.Reflection.Missing]::Value, 1)
        $table.Name = "Receipts_$SheetName"
        Write-Host "  Table    : 'Receipts_$SheetName' created over $($tableRange.Address())"
    } catch {
        Write-Host "  Warning  : Could not create table 'Receipts_$SheetName': $_" -ForegroundColor Yellow
    }

    if ($table) {
        try {
            $table.ListColumns.Item($COL_AMOUNT).DataBodyRange.NumberFormat = '$#,##0.00;[Red]($#,##0.00)'
        } catch {
            Write-Host "  Warning  : Could not set Amount number format: $_" -ForegroundColor Yellow
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
        $_.Account -eq "0000" -or
        ($_.Account -ne "" -and $_.Account -ne "xxxx" -and $ValidAccounts.Count -gt 0 -and $ValidAccounts -notcontains $_.Account)
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
        Write-Host "  Total    : Written to row $sumRow"
    } catch {
        Write-Host "  Warning  : Could not write total row: $_" -ForegroundColor Yellow
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
        Write-Host "  Warning  : Could not set freeze panes: $_" -ForegroundColor Yellow
    }

    $parsed   = ($receipts | Where-Object { $_.ParseOK }).Count
    $unparsed = ($receipts | Where-Object { -not $_.ParseOK }).Count
    Write-Host "  Written  : $parsed receipt(s)"
    if ($unparsed -gt 0) {
        Write-Host "  Unparsed : $unparsed (flagged in sheet)" -ForegroundColor Yellow
    }

    return [PSCustomObject]@{
        DataEndRow = $dataEnd
    }
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

if (-not (Test-Path $ReceiptsRoot)) {
    Write-Error "Receipts folder not found: '$ReceiptsRoot'"
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
    Write-Host "Found    : $($yearGroups.Count) year(s), $totalMonths month folder(s) to sync"
} else {
    $year        = "20" + $YearMonth.Substring(0, 2)
    $monthFolder = Get-ChildItem -Path (Join-Path $ReceiptsRoot $year) -Directory |
                   Where-Object { $_.Name -match "^$YearMonth" } |
                   Select-Object -First 1
    if (-not $monthFolder) {
        Write-Error "Could not find a folder matching '$YearMonth*' under '$ReceiptsRoot\$year'"
        exit 1
    }
    $yearGroups[$year] = @([PSCustomObject]@{ SheetName = $YearMonth; FolderPath = $monthFolder.FullName })
}

# Kill any lingering EXCEL.EXE processes before starting, if requested.
if ($KillExcel) {
    $procs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
    if ($procs) {
        Write-Host "KillExcel: Stopping $($procs.Count) EXCEL.EXE process(es)..."
        $procs | Stop-Process -Force
        Start-Sleep -Seconds 2
        Write-Host "KillExcel: Done"
    } else {
        Write-Host "KillExcel: No EXCEL.EXE processes found"
    }
}

$excel               = New-Object -ComObject Excel.Application
$excel.Visible       = $false
$excel.DisplayAlerts = $false

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
    Write-Host "Workbook : $wbPath"

    # Check for file lock before attempting to open an existing workbook.
    if (Test-Path $wbPath) {
        try {
            $stream = [System.IO.File]::Open($wbPath, 'Open', 'Read', 'None')
            $stream.Close()
        } catch {
            $procs = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
            if ($procs) {
                Write-Host "Error    : '$xlsxName' is locked by $($procs.Count) EXCEL.EXE process(es):" -ForegroundColor Red
                $procs | ForEach-Object { Write-Host "           PID $($_.Id)  Started $($_.StartTime)" -ForegroundColor Red }
                Write-Host "           Re-run with -KillExcel to terminate them automatically." -ForegroundColor Yellow
            } else {
                Write-Host "Error    : '$xlsxName' is locked but no EXCEL.EXE found." -ForegroundColor Red
                Write-Host "           Another process or the OS may be holding the file." -ForegroundColor Yellow
            }
            $excel.Quit()
            exit 1
        }
    }

    # Open the existing workbook, or create a new one if it does not exist yet.
    $workbook = $null
    if (Test-Path $wbPath) {
        try {
            $workbook = $excel.Workbooks.Open($wbPath)
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "Error    : Excel COM threw an exception opening the workbook." -ForegroundColor Red
            Write-Host "           Path   : $wbPath" -ForegroundColor Red
            Write-Host "           Detail : $errMsg" -ForegroundColor Red
            Write-Host ""
            Write-Host "Likely causes:" -ForegroundColor Yellow
            Write-Host "  1. The file is corrupted. Restore $xlsxName from a backup, then re-run." -ForegroundColor Yellow
            Write-Host "  2. Excel blocked the file because it is on a network share." -ForegroundColor Yellow
            Write-Host "     Try opening it manually in Excel first and click Enable Editing." -ForegroundColor Yellow
            Write-Host "  3. A previous Excel process is still running invisibly." -ForegroundColor Yellow
            Write-Host "     Re-run with -KillExcel to terminate it automatically, or end EXCEL.EXE in Task Manager." -ForegroundColor Yellow
            $excel.Quit()
            exit 1
        }
    } else {
        Write-Host "Creating : $xlsxName (new workbook)"
        try {
            $workbook = $excel.Workbooks.Add()
            $workbook.SaveAs($wbPath)
        } catch {
            $errMsg = $_.Exception.Message
            Write-Host "Error    : Failed to create new workbook at '$wbPath': $errMsg" -ForegroundColor Red
            $excel.Quit()
            exit 1
        }
    }
    if (-not $workbook) {
        Write-Error "Failed to open workbook (returned null): $wbPath"
        $excel.Quit()
        exit 1
    }

    Write-Host "Debug    : Workbook opened successfully"
    Write-Host "Debug    : Calling Get-ValidAccounts"
    $validAccounts = @()
    try {
        $validAccounts = Get-ValidAccounts -ReceiptsRoot $ReceiptsRoot -Excel $excel -Workbook $workbook
    } catch {
        Write-Host "  Warning  : Error in Get-ValidAccounts: $_" -ForegroundColor Yellow
    }

    Write-Host "Debug    : Calling Get-Categories"
    $categories = $null
    try {
        $categories = Get-Categories
    } catch {
        Write-Host "  Warning  : Error in Get-Categories: $_" -ForegroundColor Yellow
    }

    if ($categories) {
        $catSheet = $null
        try {
            $catSheet = Sync-CategorySheet -Workbook $workbook -Categories $categories
        } catch {
            Write-Host "  Warning  : Error in Sync-CategorySheet: $_" -ForegroundColor Yellow
        }
        if ($catSheet) {
            try {
                Set-CategoryNamedRanges -Workbook $workbook -CatSheet $catSheet -Categories $categories
                Write-Host "Categories: $($categories.Count) category/subcategory group(s) loaded"
            } catch {
                Write-Host "  Warning  : Error in Set-CategoryNamedRanges: $_" -ForegroundColor Yellow
            }
        }
    }
    Write-Host "Accounts : $($validAccounts.Count) account(s) loaded"

    $syncResults = @()
    foreach ($m in $months) {
        Write-Host ""
        Write-Host "Syncing  : $($m.FolderPath)"
        try {
            $result = Sync-Month -FolderPath $m.FolderPath -SheetName $m.SheetName -Workbook $workbook -ValidAccounts $validAccounts
            $syncResults += [PSCustomObject]@{ SheetName = $m.SheetName; Result = $result }
        } catch {
            Write-Host "  Error    : Unhandled exception syncing '$($m.SheetName)': $_" -ForegroundColor Red
            Write-Host "             Skipping this month and continuing." -ForegroundColor Yellow
            $syncResults += [PSCustomObject]@{ SheetName = $m.SheetName; Result = $null }
        }
    }

    try {
        $workbook.Save()
        Write-Host "Workbook : Saved $xlsxName"
    } catch {
        Write-Host "Error    : Failed to save workbook: $_" -ForegroundColor Red
    }
    try {
        $workbook.Close()
    } catch {
        Write-Host "Warning  : Exception closing workbook: $_" -ForegroundColor Yellow
    }

    Set-SubcategoryValidationXml -WorkbookPath $wbPath -SyncResults $syncResults
}

try {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
} catch {
    Write-Host "Warning  : Exception quitting Excel: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Done!"
