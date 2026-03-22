#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for Invoke-SensitiveDataCheck
# Run with: Invoke-Pester ./Tests

BeforeAll {
    . (Join-Path $PSScriptRoot '..\Scripts\hooks\Invoke-SensitiveDataCheck.ps1')
}

Describe 'Invoke-SensitiveDataCheck' {

    BeforeAll {
        $script:accountPattern = [PSCustomObject]@{
            name      = 'account-number-in-context'
            regex     = '\b(Card|Checking|Savings|Check|Wire|Transfer)\s+\d{4}\b'
            fileTypes = @('.md', '.ps1', '.json')
            message   = 'Possible account number adjacent to payment method token'
            allowlist = @()
        }
        $script:credentialPattern = [PSCustomObject]@{
            name      = 'credential-assignment'
            regex     = '(?i)(password|secret|api_key|token)\s*[=:]\s*(?!(<[^>]+>|\$|\{|placeholder|example|test|fake|xxx|\*+))'
            fileTypes = @('.ps1', '.md', '.json')
            message   = 'Possible credential value detected'
            allowlist = @()
        }
        $script:emailPattern = [PSCustomObject]@{
            name      = 'email-address'
            regex     = '[a-zA-Z0-9._%+\-]+@(?!example\.com|placeholder\.com)[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}'
            fileTypes = @('.ps1', '.md', '.json')
            message   = 'Possible real email address'
            allowlist = @()
        }
    }

    Context 'given a file with no pattern matches' {
        It 'returns an empty array' {
            $result = Invoke-SensitiveDataCheck -Content 'No sensitive data here.' -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'given a line with an account number adjacent to a method token' {
        It 'returns a finding with the correct pattern name' {
            $result = Invoke-SensitiveDataCheck -Content 'Card 1234' -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result.PatternName | Should -Be 'account-number-in-context'
        }

        It 'includes the correct line number in the finding' {
            $content = "Line one`nCard 1234`nLine three"
            $result = Invoke-SensitiveDataCheck -Content $content -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result.LineNumber | Should -Be 2
        }

        It 'returns one finding per matching line' {
            $content = "Card 1234`nNo match`nChecking 5678"
            $result = Invoke-SensitiveDataCheck -Content $content -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result.Count | Should -Be 2
        }

        It 'includes the matched line text in the finding' {
            $result = Invoke-SensitiveDataCheck -Content 'Card 1234' -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result.Line | Should -Be 'Card 1234'
        }
    }

    Context 'given a line with a credential assignment' {
        It 'returns a finding for the credential-assignment pattern' {
            $result = Invoke-SensitiveDataCheck -Content 'password=hunter2' -FileName 'config.ps1' -Patterns @($script:credentialPattern)
            $result.PatternName | Should -Be 'credential-assignment'
        }

        It 'does not flag a placeholder value' {
            $result = Invoke-SensitiveDataCheck -Content 'password=placeholder' -FileName 'config.ps1' -Patterns @($script:credentialPattern)
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'given a line with a real email address' {
        It 'returns a finding for the email-address pattern' {
            $result = Invoke-SensitiveDataCheck -Content 'Contact: user@gmail.com' -FileName 'doc.md' -Patterns @($script:emailPattern)
            $result.PatternName | Should -Be 'email-address'
        }

        It 'does not flag an example.com address' {
            $result = Invoke-SensitiveDataCheck -Content 'Email: user@example.com' -FileName 'doc.md' -Patterns @($script:emailPattern)
            $result | Should -BeNullOrEmpty
        }

        It 'does not flag a placeholder.com address' {
            $result = Invoke-SensitiveDataCheck -Content 'Email: user@placeholder.com' -FileName 'doc.md' -Patterns @($script:emailPattern)
            $result | Should -BeNullOrEmpty
        }
    }

    Context 'given a line ending with # nocheck' {
        It 'returns no findings for that line' {
            $result = Invoke-SensitiveDataCheck -Content 'Card 1234 # nocheck' -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result | Should -BeNullOrEmpty
        }

        It 'still checks other lines in the same file' {
            $content = "Card 1234 # nocheck`nCard 5678"
            $result = Invoke-SensitiveDataCheck -Content $content -FileName 'doc.md' -Patterns @($script:accountPattern)
            $result.Count | Should -Be 1
            $result.LineNumber | Should -Be 2
        }
    }

    Context 'given a matching line that contains an allowlist entry' {
        It 'returns no findings for that line' {
            $pattern = [PSCustomObject]@{
                name      = 'account-number-in-context'
                regex     = '\b(Card|Checking)\s+\d{4}\b'
                fileTypes = @('.md')
                message   = 'Possible account number'
                allowlist = @('example-vendor')
            }
            $result = Invoke-SensitiveDataCheck -Content 'example-vendor Card 1234' -FileName 'doc.md' -Patterns @($pattern)
            $result | Should -BeNullOrEmpty
        }

        It 'still flags lines that do not contain an allowlist entry' {
            $pattern = [PSCustomObject]@{
                name      = 'account-number-in-context'
                regex     = '\b(Card|Checking)\s+\d{4}\b'
                fileTypes = @('.md')
                message   = 'Possible account number'
                allowlist = @('example-vendor')
            }
            $result = Invoke-SensitiveDataCheck -Content 'real-vendor Card 1234' -FileName 'doc.md' -Patterns @($pattern)
            $result | Should -Not -BeNullOrEmpty
        }
    }

    Context 'given a pattern whose fileTypes does not match the filename' {
        It 'does not return findings for that pattern' {
            $pattern = [PSCustomObject]@{
                name      = 'account-number-in-context'
                regex     = '\b(Card)\s+\d{4}\b'
                fileTypes = @('.ps1')
                message   = 'Possible account number'
                allowlist = @()
            }
            $result = Invoke-SensitiveDataCheck -Content 'Card 1234' -FileName 'doc.md' -Patterns @($pattern)
            $result | Should -BeNullOrEmpty
        }

        It 'applies the pattern when the filename extension matches' {
            $pattern = [PSCustomObject]@{
                name      = 'account-number-in-context'
                regex     = '\b(Card)\s+\d{4}\b'
                fileTypes = @('.ps1')
                message   = 'Possible account number'
                allowlist = @()
            }
            $result = Invoke-SensitiveDataCheck -Content 'Card 1234' -FileName 'script.ps1' -Patterns @($pattern)
            $result | Should -Not -BeNullOrEmpty
        }
    }

    Context 'when the patterns list is empty' {
        It 'returns an empty array' {
            $result = Invoke-SensitiveDataCheck -Content 'Card 1234' -FileName 'doc.md' -Patterns @()
            $result | Should -BeNullOrEmpty
        }
    }
}
