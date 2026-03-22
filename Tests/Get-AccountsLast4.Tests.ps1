#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Get-AccountsLast4
# Run with: Invoke-Pester ./Tests

BeforeAll {
    . (Join-Path $PSScriptRoot '..\Scripts\hooks\Invoke-SensitiveDataCheck.ps1')
}

Describe 'Get-AccountsLast4' {

    BeforeAll {
        # Minimal shared strings XML with header and two Last4 values
        $script:sharedStringsWithData = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3" uniqueCount="3">
  <si><t>Last 4</t></si>
  <si><t>1234</t></si>
  <si><t>5678</t></si>
</sst>
'@

        # Worksheet where column A uses shared string references
        $script:worksheetWithSharedStrings = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2" t="s"><v>1</v></c></row>
    <row r="3"><c r="A3" t="s"><v>2</v></c></row>
  </sheetData>
</worksheet>
'@

        # Worksheet where column A uses inline numeric values (no shared string)
        $script:worksheetWithInlineNumbers = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
    <row r="2"><c r="A2"><v>1234</v></c></row>
  </sheetData>
</worksheet>
'@

        $script:emptySharedStrings = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Last 4</t></si>
</sst>
'@

        $script:worksheetHeaderOnly = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>
'@

        $script:worksheetEmpty = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>
'@
    }

    Context 'given worksheet XML with two data rows in column A using shared strings' {
        It 'returns both Last4 values as strings' {
            $result = Get-AccountsLast4 -SharedStringsXml $script:sharedStringsWithData -WorksheetXml $script:worksheetWithSharedStrings
            $result | Should -Contain '1234'
            $result | Should -Contain '5678'
        }

        It 'skips the header row' {
            $result = Get-AccountsLast4 -SharedStringsXml $script:sharedStringsWithData -WorksheetXml $script:worksheetWithSharedStrings
            $result | Should -Not -Contain 'Last 4'
            $result.Count | Should -Be 2
        }
    }

    Context 'given a Last4 stored as an inline number' {
        It 'returns it as a string' {
            $result = Get-AccountsLast4 -SharedStringsXml $script:emptySharedStrings -WorksheetXml $script:worksheetWithInlineNumbers
            $result | Should -Contain '1234'
        }
    }

    Context 'given a worksheet with only the header row' {
        It 'returns an empty array' {
            $result = Get-AccountsLast4 -SharedStringsXml $script:emptySharedStrings -WorksheetXml $script:worksheetHeaderOnly
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'given an empty worksheet' {
        It 'returns an empty array' {
            $result = Get-AccountsLast4 -SharedStringsXml $script:emptySharedStrings -WorksheetXml $script:worksheetEmpty
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'given empty XML strings' {
        It 'returns an empty array without throwing' {
            $result = Get-AccountsLast4 -SharedStringsXml '' -WorksheetXml ''
            $result | Should -BeNullOrEmpty
        }
    }
}
