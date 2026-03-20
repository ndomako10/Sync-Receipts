#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Write-SyncLog
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

Describe 'Write-SyncLog' {

    Context 'output format' {

        It 'includes a [HH:mm:ss] timestamp' {
            Mock Write-Host {}
            Write-SyncLog 'msg'
            Should -Invoke Write-Host -ParameterFilter { $Object -match '^\[\d{2}:\d{2}:\d{2}\]' }
        }

        It 'includes the message text' {
            Mock Write-Host {}
            Write-SyncLog 'hello world'
            Should -Invoke Write-Host -ParameterFilter { $Object -match 'hello world' }
        }

        It 'pads INFO tag to 5 characters' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag INFO
            Should -Invoke Write-Host -ParameterFilter { $Object -match '\[INFO \]' }
        }

        It 'pads STEP tag to 5 characters' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag STEP
            Should -Invoke Write-Host -ParameterFilter { $Object -match '\[STEP \]' }
        }

        It 'pads WARN tag to 5 characters' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag WARN
            Should -Invoke Write-Host -ParameterFilter { $Object -match '\[WARN \]' }
        }

        It 'does not pad ERROR (already 5 characters)' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag ERROR
            Should -Invoke Write-Host -ParameterFilter { $Object -match '\[ERROR\]' }
        }

        It 'pads VERB tag to 5 characters in the Write-Verbose message' {
            Mock Write-Verbose {}
            Write-SyncLog 'msg' -Tag VERB
            Should -Invoke Write-Verbose -ParameterFilter { $Message -match '\[VERB \]' }
        }
    }

    Context 'tag routing and color' {

        It 'INFO writes to Write-Host and not Write-Verbose' {
            Mock Write-Host {}
            Mock Write-Verbose {}
            Write-SyncLog 'msg' -Tag INFO
            Should -Invoke Write-Host    -Times 1
            Should -Invoke Write-Verbose -Times 0
        }

        It 'STEP writes to Write-Host with Cyan foreground' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag STEP
            Should -Invoke Write-Host -Times 1 -ParameterFilter { $ForegroundColor -eq 'Cyan' }
        }

        It 'WARN writes to Write-Host with Yellow foreground' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag WARN
            Should -Invoke Write-Host -Times 1 -ParameterFilter { $ForegroundColor -eq 'Yellow' }
        }

        It 'ERROR writes to Write-Host with Red foreground' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag ERROR
            Should -Invoke Write-Host -Times 1 -ParameterFilter { $ForegroundColor -eq 'Red' }
        }

        It 'VERB writes to Write-Verbose and not Write-Host' {
            Mock Write-Host {}
            Mock Write-Verbose {}
            Write-SyncLog 'msg' -Tag VERB
            Should -Invoke Write-Verbose -Times 1
            Should -Invoke Write-Host    -Times 0
        }

        It 'INFO does not set a foreground color' {
            Mock Write-Host {}
            Write-SyncLog 'msg' -Tag INFO
            Should -Invoke Write-Host -Times 1 -ParameterFilter { $null -eq $ForegroundColor }
        }
    }

    Context 'default tag' {

        It 'defaults to INFO when no tag is specified' {
            Mock Write-Host {}
            Write-SyncLog 'default tag test'
            Should -Invoke Write-Host -Times 1 -ParameterFilter { $Object -match '\[INFO \]' }
        }
    }

    Context 'parameter validation' {

        It 'throws for an unrecognised tag' {
            { Write-SyncLog 'msg' -Tag 'UNKNOWN' } | Should -Throw
        }
    }
}
