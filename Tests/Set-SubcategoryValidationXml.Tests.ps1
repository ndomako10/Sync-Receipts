#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Set-SubcategoryValidationXml
# Run with: Invoke-Pester ./Tests

BeforeAll {
    # Load only the function definitions -- skip the main execution block
    $scriptPath = Resolve-Path (Join-Path $PSScriptRoot '..\Scripts\Sync-Receipts.ps1')
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath, [ref]$null, [ref]$null
    )
    $ast.FindAll(
        { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false
    ) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }

    # Helper: create a minimal well-formed xlsx using .NET ZipArchive
    function New-MinimalXlsx {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [string]$Path,
            [string]$SheetName = '2603',
            [string]$SheetXml  = $null
        )
        if (-not $PSCmdlet.ShouldProcess($Path, 'Create')) { return }
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

    # Helper: create a minimal two-sheet xlsx using .NET ZipArchive
    function New-TwoSheetXlsx {
        [CmdletBinding(SupportsShouldProcess)]
        param(
            [string]$Path,
            [string]$SheetName1 = '2603',
            [string]$SheetName2 = '2602'
        )
        if (-not $PSCmdlet.ShouldProcess($Path, 'Create')) { return }
        Add-Type -AssemblyName System.IO.Compression
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $emptySheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
            ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
            '<sheetData/></worksheet>'
        $files = [ordered]@{
            '[Content_Types].xml'        = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                '<Default Extension="xml" ContentType="application/xml"/>' +
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                '</Types>'
            '_rels/.rels'                = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                '</Relationships>'
            'xl/workbook.xml'            = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                "<sheets><sheet name=""$SheetName1"" sheetId=""1"" r:id=""rId1""/>" +
                "<sheet name=""$SheetName2"" sheetId=""2"" r:id=""rId2""/></sheets>" +
                '</workbook>'
            'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>' +
                '</Relationships>'
            'xl/worksheets/sheet1.xml'   = $emptySheet
            'xl/worksheets/sheet2.xml'   = $emptySheet
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

Describe 'Set-SubcategoryValidationXml' {

    Context 'given a workbook with a single sheet and valid sync results' {
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
            $script:sheetXml | Should -Match '<formula1>INDIRECT\(SUBSTITUTE\(SUBSTITUTE\(SUBSTITUTE\(G2," ","_"\),"&amp;","_"\),"/","_"\)\)</formula1>'
        }

        It 'XML-escapes the ampersand in the INDIRECT formula as &amp;' {
            # A bare & in XML is invalid and causes Excel to report a parse error on open.
            # Verify the injected formula uses &amp; not & so the workbook opens cleanly.
            $script:sheetXml | Should -Not -Match '<formula1>INDIRECT[^<]*"&"[^<]*</formula1>'
            $script:sheetXml | Should -Match '"&amp;"'
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

    Context 'given sync results that do not match workbook content' {

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

        It 'inserts dataValidations before <hyperlinks> when present' {
            $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
                '<sheetData/>' +
                '<hyperlinks><hyperlink ref="A1"/></hyperlinks>' +
                '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' +
                '</worksheet>'
            $p = Join-Path $TestDrive 'hyperlinks-order.xlsx'
            New-MinimalXlsx -Path $p -SheetName '2603' -SheetXml $sheetXml
            $results = @([PSCustomObject]@{ SheetName = '2603'; Result = [PSCustomObject]@{ DataEndRow = 5 } })
            Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results
            $out = Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet1.xml'
            $out.IndexOf('<dataValidations') | Should -BeLessThan ($out.IndexOf('<hyperlinks'))
        }

        It 'patches both sheets in a two-sheet workbook with correct DataEndRow per sheet' {
            $p = Join-Path $TestDrive 'multi-sheet.xlsx'
            New-TwoSheetXlsx -Path $p -SheetName1 '2603' -SheetName2 '2602'
            $results = @(
                [PSCustomObject]@{ SheetName = '2603'; Result = [PSCustomObject]@{ DataEndRow = 10 } },
                [PSCustomObject]@{ SheetName = '2602'; Result = [PSCustomObject]@{ DataEndRow = 5  } }
            )
            Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results
            $xml1 = Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet1.xml'
            $xml2 = Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet2.xml'
            $xml1 | Should -Match 'sqref="G2:G10"'
            $xml2 | Should -Match 'sqref="G2:G5"'
        }

        It 'inserts dataValidations before <tableParts> when no <hyperlinks> present' {
            $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">' +
                '<sheetData/>' +
                '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' +
                '</worksheet>'
            $p = Join-Path $TestDrive 'tableparts-order.xlsx'
            New-MinimalXlsx -Path $p -SheetName '2603' -SheetXml $sheetXml
            $results = @([PSCustomObject]@{ SheetName = '2603'; Result = [PSCustomObject]@{ DataEndRow = 5 } })
            Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results
            $out = Get-XlsxEntry -Path $p -Entry 'xl/worksheets/sheet1.xml'
            $out.IndexOf('<dataValidations') | Should -BeLessThan ($out.IndexOf('<tableParts'))
        }

        It 'aborts without overwriting the workbook when the reconstructed zip is under 1 KB' {
            # Build a two-entry zip: only xl/workbook.xml and xl/_rels/workbook.xml.rels.
            # With DataEndRow=1 no sheet is patched; the reconstructed temp file contains
            # only these two tiny entries and stays well under 1024 bytes, triggering the
            # size-guard early-return path.
            $p = Join-Path $TestDrive 'tiny.xlsx'
            Add-Type -AssemblyName System.IO.Compression
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            $tinyFiles = [ordered]@{
                'xl/workbook.xml'            = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
                    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                    '<sheets><sheet name="2603" sheetId="1" r:id="rId1"/></sheets></workbook>'
                'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                    '</Relationships>'
            }
            $zip = [System.IO.Compression.ZipFile]::Open($p, [System.IO.Compression.ZipArchiveMode]::Create)
            foreach ($kvp in $tinyFiles.GetEnumerator()) {
                $e = $zip.CreateEntry($kvp.Key, [System.IO.Compression.CompressionLevel]::Optimal)
                $s = $e.Open()
                $b = [System.Text.Encoding]::UTF8.GetBytes($kvp.Value)
                $s.Write($b, 0, $b.Length)
                $s.Close()
            }
            $zip.Dispose()

            $originalBytes = [System.IO.File]::ReadAllBytes($p)
            $results = @([PSCustomObject]@{ SheetName = '2603'; Result = [PSCustomObject]@{ DataEndRow = 1 } })
            { Set-SubcategoryValidationXml -WorkbookPath $p -SyncResults $results } | Should -Not -Throw
            # Workbook must not be overwritten -- the size guard fired and discarded the temp file
            [System.IO.File]::ReadAllBytes($p) | Should -Be $originalBytes
        }
    }
}
