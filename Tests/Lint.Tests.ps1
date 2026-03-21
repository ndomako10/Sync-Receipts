#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }, PSScriptAnalyzer
#
# PSScriptAnalyzer lint tests for all Scripts/*.ps1 files
# Run with: Invoke-Pester ./Tests

BeforeAll {
    $script:scriptPath    = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\Sync-Receipts.ps1'
    $script:settingsPath  = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) '.config') 'PSScriptAnalyzerSettings.psd1'

    if (-not (Test-Path $script:settingsPath)) {
        throw "PSScriptAnalyzerSettings.psd1 not found at '$script:settingsPath' -- lint results would be unreliable"
    }

    $script:results = Invoke-ScriptAnalyzer -Path $script:scriptPath -Settings $script:settingsPath
}

Describe 'PSScriptAnalyzer -- Sync-Receipts.ps1' {

    It 'has no parse errors' {
        $errors = $script:results | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:results | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:results | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}

Describe 'PSScriptAnalyzer -- Initialize-SyncReceipts.ps1' {

    BeforeAll {
        $setupPath    = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\Initialize-SyncReceipts.ps1'
        $script:setupResults = Invoke-ScriptAnalyzer -Path $setupPath -Settings $script:settingsPath
    }

    It 'has no parse errors' {
        $errors = $script:setupResults | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:setupResults | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:setupResults | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}

Describe 'PSScriptAnalyzer -- New-AccountsTemplate.ps1' {

    BeforeAll {
        $path = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\New-AccountsTemplate.ps1'
        $script:newAcctResults = Invoke-ScriptAnalyzer -Path $path -Settings $script:settingsPath
    }

    It 'has no parse errors' {
        $errors = $script:newAcctResults | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:newAcctResults | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:newAcctResults | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}

Describe 'PSScriptAnalyzer -- Install-GitHooks.ps1' {

    BeforeAll {
        $path = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\Install-GitHooks.ps1'
        $script:installHooksResults = Invoke-ScriptAnalyzer -Path $path -Settings $script:settingsPath
    }

    It 'has no parse errors' {
        $errors = $script:installHooksResults | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:installHooksResults | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:installHooksResults | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}

Describe 'PSScriptAnalyzer -- hooks/Invoke-CommitMsgCheck.ps1' {

    BeforeAll {
        $path = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\hooks\Invoke-CommitMsgCheck.ps1'
        $script:commitMsgResults = Invoke-ScriptAnalyzer -Path $path -Settings $script:settingsPath
    }

    It 'has no parse errors' {
        $errors = $script:commitMsgResults | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:commitMsgResults | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:commitMsgResults | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}

Describe 'PSScriptAnalyzer -- hooks/Invoke-PreCommitCheck.ps1' {

    BeforeAll {
        $path = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\hooks\Invoke-PreCommitCheck.ps1'
        $script:preCommitResults = Invoke-ScriptAnalyzer -Path $path -Settings $script:settingsPath
    }

    It 'has no parse errors' {
        $errors = $script:preCommitResults | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:preCommitResults | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:preCommitResults | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}

Describe 'PSScriptAnalyzer -- hooks/Invoke-PrePushCheck.ps1' {

    BeforeAll {
        $path = Join-Path (Split-Path $PSScriptRoot -Parent) 'Scripts\hooks\Invoke-PrePushCheck.ps1'
        $script:prePushResults = Invoke-ScriptAnalyzer -Path $path -Settings $script:settingsPath
    }

    It 'has no parse errors' {
        $errors = $script:prePushResults | Where-Object { $_.Severity -eq 'ParseError' }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Error severity' {
        $errors = $script:prePushResults | Where-Object { $_.Severity -eq 'Error' }
        $errors | ForEach-Object { Write-Host "  ERROR: $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Red }
        $errors | Should -BeNullOrEmpty
    }

    It 'has no rule violations at Warning severity' {
        $warnings = $script:prePushResults | Where-Object { $_.Severity -eq 'Warning' }
        $warnings | ForEach-Object { Write-Host "  WARN:  $($_.RuleName) line $($_.Line) -- $($_.Message)" -ForegroundColor Yellow }
        $warnings | Should -BeNullOrEmpty
    }
}
