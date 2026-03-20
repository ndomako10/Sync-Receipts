#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Test-SubcategoryValid
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

Describe 'Test-SubcategoryValid' {
    BeforeAll {
        $script:cats = [ordered]@{
            'Food & Dining' = @('Groceries', 'Restaurants')
            'Housing'       = @('Rent', 'Utilities')
            'Shopping'      = @()
        }
    }

    Context 'given a valid category and matching subcategory' {
        It 'returns true' {
            Test-SubcategoryValid -Category 'Food & Dining' -Subcategory 'Groceries' -Categories $script:cats |
                Should -BeTrue
        }
        It 'returns true for another valid pair' {
            Test-SubcategoryValid -Category 'Housing' -Subcategory 'Utilities' -Categories $script:cats |
                Should -BeTrue
        }
    }

    Context 'given a subcategory that does not belong to the category' {
        It 'returns false' {
            Test-SubcategoryValid -Category 'Food & Dining' -Subcategory 'Rent' -Categories $script:cats |
                Should -BeFalse
        }
    }

    Context 'given a category not present in the list' {
        It 'returns false' {
            Test-SubcategoryValid -Category 'Unknown' -Subcategory 'Groceries' -Categories $script:cats |
                Should -BeFalse
        }
    }

    Context 'given a category with no subcategories' {
        It 'returns false for any subcategory' {
            Test-SubcategoryValid -Category 'Shopping' -Subcategory 'Anything' -Categories $script:cats |
                Should -BeFalse
        }
    }

    Context 'given a null or empty Categories hashtable' {
        It 'returns true when Categories is null so existing values are preserved' {
            Test-SubcategoryValid -Category 'Food & Dining' -Subcategory 'Groceries' -Categories $null |
                Should -BeTrue
        }
        It 'returns true when Categories is empty so existing values are preserved' {
            Test-SubcategoryValid -Category 'Food & Dining' -Subcategory 'Groceries' -Categories @{} |
                Should -BeTrue
        }
    }
}
