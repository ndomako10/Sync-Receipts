#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Sync-Receipts.ps1
# Run with: Invoke-Pester ./tests

BeforeAll {
    # Load only the function definitions -- skip the main execution block
    $scriptPath = Resolve-Path (Join-Path $PSScriptRoot '..\Sync-Receipts.ps1')
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath, [ref]$null, [ref]$null
    )
    $ast.FindAll(
        { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false
    ) | ForEach-Object { Invoke-Expression $_.Extent.Text }

    # Helper: create a minimal well-formed xlsx using .NET ZipArchive
    function New-MinimalXlsx {
        param(
            [string]$Path,
            [string]$SheetName = '2603',
            [string]$SheetXml  = $null
        )
        if (-not $SheetXml) {
            $SheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
                ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
                '<sheetData>' +
                '<row r="1"><c r="A1" t="inlineStr"><is><t>File Name</t></is></c></row>' +
                '<row r="2"><c r="A2" t="inlineStr"><is><t>receipt.pdf</t></is></c></row>' +
                '</sheetData></worksheet>'
        }
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $files = [ordered]@{
            '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                '<Default Extension="xml" ContentType="application/xml"/>' +
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '</Types>'
            '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>'
            'xl/workbook.xml'            = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                "<sheets><sheet name=""$SheetName"" sheetId=""1"" r:id=""rId1""/></sheets>" +
                '</workbook>'
            'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                '</Relationships>'
            'xl/worksheets/sheet1.xml'   = $SheetXml
        }
        $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($kvp in $files.GetEnumerator()) {
            $e = $zip.CreateEntry($kvp.Key, [System.IO.Compression.CompressionLevel]::Optimal)
            $s = $e.Open()
            $b = [System.Text.Encoding]::UTF8.GetBytes($kvp.Value)
            $s.Write($b, 0, $b.Length)
            $s.Close()
        }
        $zip.Dispose()
    }

    # Helper: read a named entry out of an xlsx (zip)
    function Get-XlsxEntry {
        param([string]$Path, [string]$Entry)
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        $e   = $zip.Entries | Where-Object { $_.FullName -eq $Entry }
        $sr  = New-Object System.IO.StreamReader($e.Open())
        $c   = $sr.ReadToEnd()
        $sr.Close(); $zip.Dispose(); $c
    }
}

# ---------------------------------------------------------------------------
Describe 'Parse-Receipt' {

    Context 'valid filenames' {

        It 'parses all fields: date, vendor, amount, method, account' {
            $r = Parse-Receipt -Stem '260301 Amazon -$10.00 Card 1234'
            $r.OK      | Should -Be $true
            $r.Date    | Should -Be ([datetime]'2026-03-01')
            $r.Vendor  | Should -Be 'Amazon'
            $r.Amount  | Should -Be '-10.00'
            $r.Method  | Should -Be 'Card'
            $r.Account | Should -Be '1234'
        }

        It 'parses a vendor with multiple spaces' {
            $r = Parse-Receipt -Stem '260301 Amazon Prime Video -$15.99 Card 1234'
            $r.OK     | Should -Be $true
            $r.Vendor | Should -Be 'Amazon Prime Video'
        }

        It 'parses amount with no minus sign' {
            $r = Parse-Receipt -Stem '260301 Walmart $5.50 Cash'
            $r.OK     | Should -Be $true
            $r.Amount | Should -Be '5.50'
        }

        It 'accepts Cash method with no account' {
            $r = Parse-Receipt -Stem '260301 Walmart -$5.50 Cash'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Cash'
            $r.Account | Should -Be ''
        }

        It 'accepts Checking method with account' {
            $r = Parse-Receipt -Stem '260301 Bank -$100.00 Checking 5678'
            $r.OK     | Should -Be $true
            $r.Method | Should -Be 'Checking'
            $r.Account| Should -Be '5678'
        }

        It 'accepts Savings method with xxxx account' {
            $r = Parse-Receipt -Stem '260301 ATM -$200.00 Savings xxxx'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Savings'
            $r.Account | Should -Be 'xxxx'
        }

        It 'accepts Card with no account' {
            $r = Parse-Receipt -Stem '260301 Amazon -$10.00 Card'
            $r.OK      | Should -Be $true
            $r.Account | Should -Be ''
        }

        It 'strips the dollar sign from amount' {
            $r = Parse-Receipt -Stem '260301 Shop -$1234.56 Card 9999'
            $r.Amount | Should -Be '-1234.56'
        }

        It 'preserves the negative sign and strips the dollar sign' {
            $r = Parse-Receipt -Stem '260301 Costco -$52.37 Card 1234'
            $r.OK     | Should -Be $true
            $r.Amount | Should -Be '-52.37'
            $r.Amount | Should -Match '^-'
            $r.Amount | Should -Not -Match '\$'
        }
    }

    Context 'invalid filenames' {

        It 'returns ParseOK=false for a non-receipt filename' {
            $r = Parse-Receipt -Stem 'not a receipt'
            $r.OK | Should -Be $false
        }

        It 'returns ParseOK=false for an invalid date' {
            $r = Parse-Receipt -Stem '999999 Amazon -$10.00 Card 1234'
            $r.OK | Should -Be $false
        }

        It 'returns ParseOK=false when method is missing' {
            $r = Parse-Receipt -Stem '260301 Amazon -$10.00'
            $r.OK | Should -Be $false
        }

        It 'returns ParseOK=false when amount has no decimal' {
            $r = Parse-Receipt -Stem '260301 Amazon -$10 Card 1234'
            $r.OK | Should -Be $false
        }

        It 'returns empty strings for all fields when invalid' {
            $r = Parse-Receipt -Stem 'garbage'
            $r.Vendor  | Should -Be ''
            $r.Amount  | Should -Be ''
            $r.Method  | Should -Be ''
            $r.Account | Should -Be ''
        }
    }
}

# ---------------------------------------------------------------------------
Describe 'Set-SubcategoryValidationXml' {

    Context 'happy path -- single sheet' {
        BeforeAll {
            $script:xlsxPath = Join-Path $TestDrive 'happy.xlsx'
            New-MinimalXlsx -Path $script:xlsxPath -SheetName '2603'

            $results = @([PSCustomObject]@{
                SheetName = '2603'
                Result    = [PSCustomObject]@{ DataEndRow = 10 }
            })
            Set-SubcategoryValidationXml -WorkbookPath $script:xlsxPath -SyncResults $results

            $script:sheetXml = Get-XlsxEntry -Path $script:xlsxPath -Entry 'xl/worksheets/sheet1.xml'
            $script:bytes    = [System.IO.File]::ReadAllBytes($script:xlsxPath)
        }

        It 'injects dataValidations with count="2"' {
            $script:sheetXml | Should -Match 'count="2"'
        }

        It 'includes Category formula on the correct sqref' {
            $script:sheetXml | Should -Match 'sqref="G2:G10"'
            $script:sheetXml | Should -Match '<formula1>Category</formula1>'
        }

        It 'includes INDIRECT formula on the correct sqref' {
            $script:sheetXml | Should -Match 'sqref="H2:H10"'
            $script:sheetXml | Should -Match '<formula1>INDIRECT\(G2\)</formula1>'
        }

        It 'adds xr:uid to both dataValidation elements' {
            ([regex]::Matches($script:sheetXml, 'xr:uid="\{[0-9A-F-]+\}"')).Count | Should -Be 2
        }

        It 'sets flag_bits=0x0006 on all local file headers' {
            for ($i = 0; $i -lt $script:bytes.Length - 4; $i++) {
                if ($script:bytes[$i] -eq 0x50 -and $script:bytes[$i+1] -eq 0x4B -and
                    $script:bytes[$i+2] -eq 0x03 -and $script:bytes[$i+3] -eq 0x04) {
                    $script:bytes[$i+6] | Should -Be 0x06
                    $script:bytes[$i+7] | Should -Be 0x00
                }
            }
        }

        It 'sets version_made_by=45 and flag_bits=0x0006 on all central directory headers' {
            for ($i = 0; $i -lt $script:bytes.Length - 4; $i++) {
                if ($script:bytes[$i] -eq 0x50 -and $script:bytes[$i+1] -eq 0x4B -and
                    $script:bytes[$i+2] -eq 0x01 -and $script:bytes[$i+3] -eq 0x02) {
                    $script:bytes[$i+4] | Should -Be 0x2D
                    $script:bytes[$i+5] | Should -Be 0x00
                    $script:bytes[$i+8] | Should -Be 0x06
                    $script:bytes[$i+9] | Should -Be 0x00
                }
            }
        }
    }

    Context 'edge cases' {

        It 'skips a sheet whose DataEndRow is less than 2' {
            $p = Join-Path $TestDrive 'skip-nodata.xlsx'
            New-MinimalXlsx -Path $p -SheetName '2603'
            $results = @([PSCustomObject]@{
                SheetName = '2603'
                Result    = [PSCustomObject]@{ DataEndRow = 1 }
            })
            Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results
            Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet1.xml' | Should -Not -Match 'dataValidations'
        }

        It 'skips a sheet name not present in the workbook' {
            $p = Join-Path $TestDrive 'skip-unknown.xlsx'
            New-MinimalXlsx -Path $p -SheetName '2603'
            $results = @([PSCustomObject]@{
                SheetName = '9999'
                Result    = [PSCustomObject]@{ DataEndRow = 10 }
            })
            Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results
            Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet1.xml' | Should -Not -Match 'dataValidations'
        }

        It 'replaces a pre-existing dataValidations block' {
            $existingSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
                '<sheetData/>' +
                '<dataValidations count="1"><dataValidation type="list" sqref="H2:H5">' +
                '<formula1>INDIRECT(G2)</formula1></dataValidation></dataValidations>' +
                '</worksheet>'
            $p = Join-Path $TestDrive 'replace-dv.xlsx'
            New-MinimalXlsx -Path $p -SheetName '2603' -SheetXml $existingSheetXml
            $results = @([PSCustomObject]@{
                SheetName = '2603'
                Result    = [PSCustomObject]@{ DataEndRow = 15 }
            })
            Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results
            $xml = Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet1.xml'
            $xml | Should -Match 'count="2"'
            ([regex]::Matches($xml, '<dataValidation ')).Count | Should -Be 2
        }

        It 'skips a sheet entry with a null Result' {
            $p = Join-Path $TestDrive 'skip-null.xlsx'
            New-MinimalXlsx -Path $p -SheetName '2603'
            $results = @([PSCustomObject]@{ SheetName = '2603'; Result = $null })
            { Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results } | Should -Not -Throw
        }
    }
}
