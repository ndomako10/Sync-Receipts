#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Get-Methods
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

Describe 'Get-Methods' {

    Context 'reading from Methods.json' {

        It 'returns the token list from a valid JSON file' {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            '["Card","Checking","Wire"]' | Set-Content (Join-Path $dir 'Methods.json')
            $result = Get-Methods -ConfigRoot $dir
            $result | Should -Contain 'Card'
            $result | Should -Contain 'Checking'
            $result | Should -Contain 'Wire'
            $result.Count | Should -Be 3
        }

        It 'returns defaults when Methods.json is absent' {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $result = Get-Methods -ConfigRoot $dir
            $result | Should -Contain 'Card'
            $result | Should -Contain 'Checking'
            $result.Count | Should -Be 6
        }

        It 'returns defaults when Methods.json is malformed' {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            'not valid json' | Set-Content (Join-Path $dir 'Methods.json')
            $result = Get-Methods -ConfigRoot $dir
            $result | Should -Contain 'Card'
        }

        It 'returns defaults when Methods.json contains an empty array' {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            '[]' | Set-Content (Join-Path $dir 'Methods.json')
            $result = Get-Methods -ConfigRoot $dir
            $result | Should -Contain 'Card'
            $result.Count | Should -Be 6
        }

        It 'skips tokens containing spaces or special characters' {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            '["Card","Bad Token","Wire"]' | Set-Content (Join-Path $dir 'Methods.json')
            $result = Get-Methods -ConfigRoot $dir
            $result | Should -Contain 'Card'
            $result | Should -Contain 'Wire'
            $result | Should -Not -Contain 'Bad Token'
        }

        It 'returns defaults when all tokens are invalid' {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            '["bad token","another bad"]' | Set-Content (Join-Path $dir 'Methods.json')
            $result = Get-Methods -ConfigRoot $dir
            $result | Should -Contain 'Card'
        }
    }
}
