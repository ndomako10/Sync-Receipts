#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Read-PreservedCategoryValues
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

Describe 'Read-PreservedCategoryValues' {

    Context 'standard layout' {

        It 'returns category and subcategory keyed by filename' {
            $data = @(
                @('File Name','Date','Vendor','Amount','Method','Account','Category','Subcategory','Flag'),
                @('260301 Landlord -$1200.00 Checking 4455.pdf','','','','','','Housing','Rent / Mortgage',''),
                @('260302 Walmart -$58.22 Card 3232.pdf','','','','','','Food','Groceries','')
            )
            $result = Read-PreservedCategoryValues -SheetData $data
            $result.Count                                              | Should -Be 2
            $result['260301 Landlord -$1200.00 Checking 4455.pdf'].Category    | Should -Be 'Housing'
            $result['260301 Landlord -$1200.00 Checking 4455.pdf'].Subcategory | Should -Be 'Rent / Mortgage'
        }

        It 'skips rows where both Category and Subcategory are blank' {
            $data = @(
                @('File Name','Category','Subcategory'),
                @('260301 receipt.pdf','Food','Groceries'),
                @('260302 receipt.pdf','','')
            )
            $result = Read-PreservedCategoryValues -SheetData $data
            $result.Count | Should -Be 1
            $result.ContainsKey('260302 receipt.pdf') | Should -BeFalse
        }

        It 'preserves a row with only Category filled in' {
            $data = @(
                @('File Name','Category','Subcategory'),
                @('260301 receipt.pdf','Housing','')
            )
            $result = Read-PreservedCategoryValues -SheetData $data
            $result['260301 receipt.pdf'].Category    | Should -Be 'Housing'
            $result['260301 receipt.pdf'].Subcategory | Should -Be ''
        }
    }

    Context 'mismatched layout' {

        It 'finds columns regardless of order' {
            $data = @(
                @('Category','File Name','Flag','Subcategory'),
                @('Food','260301 receipt.pdf','','Groceries')
            )
            $result = Read-PreservedCategoryValues -SheetData $data
            $result['260301 receipt.pdf'].Category    | Should -Be 'Food'
            $result['260301 receipt.pdf'].Subcategory | Should -Be 'Groceries'
        }

        It 'returns empty hashtable when File Name column is absent' {
            $data = @(
                @('Date','Category','Subcategory'),
                @('1-Mar','Food','Groceries')
            )
            $result = Read-PreservedCategoryValues -SheetData $data
            $result.Count | Should -Be 0
        }

        It 'returns empty hashtable when neither Category nor Subcategory column exists' {
            $data = @(
                @('File Name','Date','Vendor'),
                @('260301 receipt.pdf','1-Mar','Walmart')
            )
            $result = Read-PreservedCategoryValues -SheetData $data
            $result.Count | Should -Be 0
        }

        It 'returns empty hashtable when SheetData has fewer than 2 rows' {
            $result = Read-PreservedCategoryValues -SheetData @(@('File Name','Category','Subcategory'))
            $result.Count | Should -Be 0
        }
    }
}
