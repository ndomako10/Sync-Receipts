#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Get-Categories
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

# ---------------------------------------------------------------------------
# Get-ValidAccounts is not unit-testable: it requires a live Excel COM
# object ($Excel.Workbooks.Open) to read Accounts.xlsx. Integration testing
# is local-only (see Tests/run-sync-test.bat).
# ---------------------------------------------------------------------------

Describe 'Get-Categories' {

    Context 'reading from Categories.json' {

        It 'returns an ordered hashtable with the correct keys' {
            $json = '{"Food":["Groceries","Restaurants"],"Housing":["Rent / Mortgage","HOA Fees"]}'
            $path = Join-Path $TestDrive 'Categories.json'
            Set-Content $path $json
            $result = Get-Categories -ReceiptsRoot $TestDrive
            $result              | Should -BeOfType [System.Collections.Specialized.OrderedDictionary]
            $result.Keys         | Should -Contain 'Food'
            $result.Keys         | Should -Contain 'Housing'
        }

        It 'preserves category order from the JSON file' {
            $json = '{"Zebra":["Z1"],"Apple":["A1"],"Mango":["M1"]}'
            $path = Join-Path $TestDrive 'Categories.json'
            Set-Content $path $json
            $result = Get-Categories -ReceiptsRoot $TestDrive
            @($result.Keys)[0] | Should -Be 'Zebra'
            @($result.Keys)[1] | Should -Be 'Apple'
            @($result.Keys)[2] | Should -Be 'Mango'
        }

        It 'returns subcategories as a string array' {
            $json = '{"Food":["Groceries","Restaurants","Fast Food"]}'
            $path = Join-Path $TestDrive 'Categories.json'
            Set-Content $path $json
            $result = Get-Categories -ReceiptsRoot $TestDrive
            $result['Food'].Count | Should -Be 3
            $result['Food'][0]    | Should -Be 'Groceries'
            $result['Food'][2]    | Should -Be 'Fast Food'
        }

        It 'returns null when Categories.json is absent' {
            $emptyDir = Join-Path $TestDrive 'empty'
            New-Item $emptyDir -ItemType Directory -Force | Out-Null
            Get-Categories -ReceiptsRoot $emptyDir | Should -BeNullOrEmpty
        }

        It 'returns null and does not throw when Categories.json is malformed' {
            $path = Join-Path $TestDrive 'Categories.json'
            Set-Content $path 'not valid json {'
            { Get-Categories -ReceiptsRoot $TestDrive } | Should -Not -Throw
            Get-Categories -ReceiptsRoot $TestDrive | Should -BeNullOrEmpty
        }
    }
}
