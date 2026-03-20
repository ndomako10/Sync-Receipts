#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Get-ExcelColumnLetter
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
}

Describe 'Get-ExcelColumnLetter' {

    Context 'single-letter columns (A-Z)' {

        It 'returns A for column 1' {
            Get-ExcelColumnLetter -Col 1 | Should -Be 'A'
        }

        It 'returns B for column 2' {
            Get-ExcelColumnLetter -Col 2 | Should -Be 'B'
        }

        It 'returns Q for column 17 (the category-range use case)' {
            Get-ExcelColumnLetter -Col 17 | Should -Be 'Q'
        }

        It 'returns Z for column 26 (last single-letter column)' {
            Get-ExcelColumnLetter -Col 26 | Should -Be 'Z'
        }
    }

    Context 'two-letter columns (AA-ZZ)' {

        It 'returns AA for column 27 (first two-letter column)' {
            Get-ExcelColumnLetter -Col 27 | Should -Be 'AA'
        }

        It 'returns AB for column 28' {
            Get-ExcelColumnLetter -Col 28 | Should -Be 'AB'
        }

        It 'returns AZ for column 52' {
            Get-ExcelColumnLetter -Col 52 | Should -Be 'AZ'
        }

        It 'returns BA for column 53 (second group of two-letter columns)' {
            Get-ExcelColumnLetter -Col 53 | Should -Be 'BA'
        }

        It 'returns ZZ for column 702 (last two-letter column)' {
            Get-ExcelColumnLetter -Col 702 | Should -Be 'ZZ'
        }
    }

    Context 'three-letter columns' {

        It 'returns AAA for column 703 (first three-letter column)' {
            Get-ExcelColumnLetter -Col 703 | Should -Be 'AAA'
        }

        It 'returns AAB for column 704' {
            Get-ExcelColumnLetter -Col 704 | Should -Be 'AAB'
        }
    }

    Context 'edge case' {

        It 'returns empty string for column 0' {
            Get-ExcelColumnLetter -Col 0 | Should -Be ''
        }
    }
}
