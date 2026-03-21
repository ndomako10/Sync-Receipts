#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Get-ReceiptFlag
# Run with: Invoke-Pester ./Tests

BeforeAll {
    $scriptPath = Resolve-Path (Join-Path $PSScriptRoot '..\Scripts\Sync-Receipts.ps1')
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $scriptPath, [ref]$null, [ref]$null
    )
    $ast.FindAll(
        { $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] },
        $false
    ) | ForEach-Object { . ([scriptblock]::Create($_.Extent.Text)) }
}

Describe 'Get-ReceiptFlag' {

    Context 'given a parse error with a non-empty ParseError string' {
        It 'returns the parse error string' {
            $r = [PSCustomObject]@{ ParseOK=$false; ParseError='Unrecognised method'; Method=''; Account='' }
            Get-ReceiptFlag -ParseResult $r | Should -Be 'Unrecognised method'
        }
    }

    Context 'given a parse error with an empty ParseError string' {
        It 'returns "Could not parse filename"' {
            $r = [PSCustomObject]@{ ParseOK=$false; ParseError=''; Method=''; Account='' }
            Get-ReceiptFlag -ParseResult $r | Should -Be 'Could not parse filename'
        }
    }

    Context 'given a successful parse with no method' {
        It 'returns "Method missing"' {
            $r = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method=''; Account='' }
            Get-ReceiptFlag -ParseResult $r | Should -Be 'Method missing'
        }
    }

    Context 'given a successful parse with account "xxxx"' {
        It 'returns "Account obfuscated"' {
            $r = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Card'; Account='xxxx' }
            Get-ReceiptFlag -ParseResult $r | Should -Be 'Account obfuscated'
        }
    }

    Context 'given a successful parse with account "----"' {
        It 'returns "Account unknown"' {
            $r = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Card'; Account='----' }
            Get-ReceiptFlag -ParseResult $r | Should -Be 'Account unknown'
        }
    }

    Context 'given an account not present in ValidAccounts' {
        It 'returns "Account not in Accounts.xlsx"' {
            $r        = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Card'; Account='8888' }
            $accounts = @(@{ Last4='1234'; Method='Card'; Account='My Card'; Status='Active' })
            Get-ReceiptFlag -ParseResult $r -ValidAccounts $accounts | Should -Be 'Account not in Accounts.xlsx'
        }
    }

    Context 'given an account in ValidAccounts with Status Inactive' {
        It 'returns "Account inactive"' {
            $r        = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Card'; Account='9999' }
            $accounts = @(@{ Last4='9999'; Method='Card'; Account='My Card'; Status='Inactive' })
            Get-ReceiptFlag -ParseResult $r -ValidAccounts $accounts | Should -Be 'Account inactive'
        }
    }

    Context 'given a valid Card receipt with an Active account in ValidAccounts' {
        It 'returns an empty string' {
            $r        = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Card'; Account='1234' }
            $accounts = @(@{ Last4='1234'; Method='Card'; Account='My Card'; Status='Active' })
            Get-ReceiptFlag -ParseResult $r -ValidAccounts $accounts | Should -Be ''
        }
    }

    Context 'given a Cash receipt with no account' {
        It 'returns an empty string' {
            $r = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Cash'; Account='' }
            Get-ReceiptFlag -ParseResult $r | Should -Be ''
        }
    }

    Context 'given a non-Cash receipt with an account and an empty ValidAccounts list' {
        It 'returns an empty string' {
            $r = [PSCustomObject]@{ ParseOK=$true; ParseError=''; Method='Card'; Account='1234' }
            Get-ReceiptFlag -ParseResult $r -ValidAccounts @() | Should -Be ''
        }
    }
}
