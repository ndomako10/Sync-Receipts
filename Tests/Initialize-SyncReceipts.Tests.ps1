#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Initialize-SyncReceipts.ps1 helpers
# Run with: Invoke-Pester ./Tests

BeforeAll {
    $scriptPath = Resolve-Path (Join-Path $PSScriptRoot '..\Scripts\Initialize-SyncReceipts.ps1')
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath, [ref]$null, [ref]$null
    )
    $ast.FindAll(
        { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false
    ) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
}

Describe 'Copy-ConfigTemplate' {

    Context 'given the destination file does not exist' {
        BeforeEach {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $script:src  = Join-Path $dir 'source.json'
            $script:dest = Join-Path $dir 'dest.json'
            'template-content' | Set-Content $script:src
            Mock Write-OK   {}
            Mock Write-Skip {}
            Mock Write-Step {}
            Mock Write-Fail {}
        }

        It 'copies the source file to the destination' {
            Copy-ConfigTemplate -Source $script:src -Destination $script:dest -Label 'Test File'
            Test-Path $script:dest | Should -BeTrue
        }

        It 'returns $true' {
            $result = Copy-ConfigTemplate -Source $script:src -Destination $script:dest -Label 'Test File'
            $result | Should -Be $true
        }
    }

    Context 'given the destination file already exists' {
        BeforeEach {
            $dir = Join-Path $TestDrive ([System.Guid]::NewGuid())
            New-Item -ItemType Directory -Path $dir | Out-Null
            $script:src  = Join-Path $dir 'source.json'
            $script:dest = Join-Path $dir 'dest.json'
            'template-content'  | Set-Content $script:src
            'original-content'  | Set-Content $script:dest
            Mock Write-OK   {}
            Mock Write-Skip {}
            Mock Write-Step {}
            Mock Write-Fail {}
        }

        It 'does not overwrite the destination file' {
            Copy-ConfigTemplate -Source $script:src -Destination $script:dest -Label 'Test File'
            Get-Content $script:dest | Should -Be 'original-content'
        }

        It 'returns $false' {
            $result = Copy-ConfigTemplate -Source $script:src -Destination $script:dest -Label 'Test File'
            $result | Should -Be $false
        }
    }
}
