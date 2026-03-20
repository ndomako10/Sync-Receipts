#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for ConvertTo-ExcelRangeName
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

Describe 'ConvertTo-ExcelRangeName' {

    Context 'given a name with no special characters' {
        It 'returns the name unchanged' {
            ConvertTo-ExcelRangeName -Name 'Housing' | Should -Be 'Housing'
        }
        It 'preserves underscores' {
            ConvertTo-ExcelRangeName -Name 'My_Category' | Should -Be 'My_Category'
        }
    }

    Context 'given a name with spaces' {
        It 'replaces each space with an underscore' {
            ConvertTo-ExcelRangeName -Name 'Personal Care' | Should -Be 'Personal_Care'
        }
    }

    Context 'given a name with ampersands' {
        It 'replaces ampersand and surrounding spaces with underscores' {
            ConvertTo-ExcelRangeName -Name 'Food & Dining' | Should -Be 'Food___Dining'
        }
        It 'handles multiple ampersand categories' {
            ConvertTo-ExcelRangeName -Name 'Bills & Utilities' | Should -Be 'Bills___Utilities'
        }
    }

    Context 'given a name with forward slashes' {
        It 'replaces each slash with an underscore' {
            ConvertTo-ExcelRangeName -Name 'Food/Drink' | Should -Be 'Food_Drink'
        }
    }

    Context 'given a name with mixed special characters' {
        It 'replaces all non-identifier characters with underscores' {
            ConvertTo-ExcelRangeName -Name 'A & B/C' | Should -Be 'A___B_C'
        }
    }

    Context 'given a name starting with a digit' {
        It 'prefixes with an underscore so the name is a valid Excel identifier' {
            ConvertTo-ExcelRangeName -Name '401k' | Should -Be '_401k'
        }
    }
}
