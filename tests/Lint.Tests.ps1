#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }, PSScriptAnalyzer
#
# PSScriptAnalyzer lint tests for Sync-Receipts.ps1
# Run with: Invoke-Pester ./tests

BeforeAll {
    $script:scriptPath    = Join-Path (Split-Path $PSScriptRoot -Parent) 'Sync-Receipts.ps1'
    $script:settingsPath  = Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) '.config') 'PSScriptAnalyzerSettings.psd1'

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
