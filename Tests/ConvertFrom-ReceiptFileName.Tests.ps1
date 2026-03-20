#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.0' }
#
# Pester unit tests for ConvertFrom-ReceiptFileName
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

Describe 'ConvertFrom-ReceiptFileName' {

    Context 'valid filenames' {

        # Baseline
        It 'parses all fields: date, vendor, amount, method, account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$10.00 Card 1234'
            $r.OK      | Should -Be $true
            $r.Date    | Should -Be ([datetime]'2026-03-01')
            $r.Vendor  | Should -Be 'Amazon'
            $r.Amount  | Should -Be '-10.00'
            $r.Method  | Should -Be 'Card'
            $r.Account | Should -Be '1234'
        }

        # Vendor
        It 'parses a vendor with multiple spaces' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon Prime Video -$15.99 Card 1234'
            $r.OK     | Should -Be $true
            $r.Vendor | Should -Be 'Amazon Prime Video'
        }

        # Amount
        It 'parses amount with no minus sign' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Walmart $5.50 Cash'
            $r.OK     | Should -Be $true
            $r.Amount | Should -Be '5.50'
        }

        It 'strips the dollar sign from amount' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Shop -$1234.56 Card 9999'
            $r.Amount | Should -Be '-1234.56'
        }

        It 'preserves the negative sign and strips the dollar sign' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Costco -$52.37 Card 1234'
            $r.OK     | Should -Be $true
            $r.Amount | Should -Be '-52.37'
            $r.Amount | Should -Match '^-'
            $r.Amount | Should -Not -Match '\$'
        }

        # Method and Account -- Cash
        It 'accepts Cash method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Walmart -$5.50 Cash'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Cash'
            $r.Account | Should -Be ''
        }

        # Method and Account -- Card
        It 'accepts Card with xxxx redacted account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$34.99 Card xxxx'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Card'
            $r.Account | Should -Be 'xxxx'
        }

        It 'accepts Card with ---- unknown account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Costco -$67.50 Card ----'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Card'
            $r.Account | Should -Be '----'
        }

        It 'accepts 0000 as a regular account number' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Shop -$5.00 Card 0000'
            $r.OK      | Should -Be $true
            $r.Account | Should -Be '0000'
        }

        # Method and Account -- Checking / Savings
        It 'accepts Checking method with account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Bank -$100.00 Checking 5678'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Checking'
            $r.Account | Should -Be '5678'
        }

        It 'accepts Savings method with xxxx account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 ATM -$200.00 Savings xxxx'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Savings'
            $r.Account | Should -Be 'xxxx'
        }

        # Method and Account -- Check / Wire / Transfer
        It 'accepts Check method with account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Landlord -$1200.00 Check 1234'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Check'
            $r.Account | Should -Be '1234'
        }

        It 'accepts Wire method with account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Vendor -$500.00 Wire 5678'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Wire'
            $r.Account | Should -Be '5678'
        }

        It 'accepts Transfer method with account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Friend -$50.00 Transfer 9012'
            $r.OK      | Should -Be $true
            $r.Method  | Should -Be 'Transfer'
            $r.Account | Should -Be '9012'
        }
    }

    Context 'invalid filenames' {

        # Structural failures
        It 'returns ParseOK=false for a non-receipt filename' {
            $r = ConvertFrom-ReceiptFileName -Stem 'not a receipt'
            $r.OK | Should -Be $false
        }

        It 'returns ParseOK=false and ParseError for a non-receipt filename' {
            $r = ConvertFrom-ReceiptFileName -Stem 'garbage'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Could not parse filename'
        }

        It 'returns empty strings for all fields when invalid' {
            $r = ConvertFrom-ReceiptFileName -Stem 'garbage'
            $r.Vendor  | Should -Be ''
            $r.Amount  | Should -Be ''
            $r.Method  | Should -Be ''
            $r.Account | Should -Be ''
        }

        It 'returns ParseOK=false when amount has no decimal' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$10 Card 1234'
            $r.OK | Should -Be $false
        }

        # Date validation
        It 'returns ParseOK=false and ParseError for an out-of-range month' {
            $r = ConvertFrom-ReceiptFileName -Stem '261316 Amazon -$10.00 Card 1234'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Month out of range'
        }

        It 'returns ParseOK=false and ParseError for an out-of-range day' {
            $r = ConvertFrom-ReceiptFileName -Stem '260332 Amazon -$10.00 Card 1234'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Day out of range'
        }

        It 'returns ParseOK=false and ParseError for an invalid date (e.g. Feb 31)' {
            $r = ConvertFrom-ReceiptFileName -Stem '260231 Amazon -$10.00 Card 1234'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Invalid date'
        }

        # Method validation
        It 'returns ParseError="Unrecognised method" for an unknown method token' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Shop -$10.00 Debit 1234'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Unrecognised method'
        }

        # Missing account
        It 'rejects Card method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$10.00 Card'
            $r.OK | Should -Be $false
        }

        It 'rejects Checking method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Bank -$100.00 Checking'
            $r.OK | Should -Be $false
        }

        It 'rejects Savings method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 ATM -$200.00 Savings'
            $r.OK | Should -Be $false
        }

        It 'rejects Check method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Landlord -$1200.00 Check'
            $r.OK | Should -Be $false
        }

        It 'rejects Wire method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Vendor -$500.00 Wire'
            $r.OK | Should -Be $false
        }

        It 'rejects Transfer method with no account' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Friend -$50.00 Transfer'
            $r.OK | Should -Be $false
        }
    }

    Context 'no Method or Account' {

        It 'returns OK=true when Method and Account are omitted' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$10.00'
            $r.OK | Should -Be $true
        }

        It 'returns empty Method and Account when both are omitted' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$10.00'
            $r.Method  | Should -Be ''
            $r.Account | Should -Be ''
        }

        It 'returns ParseError="" when Method is omitted' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon -$10.00'
            $r.ParseError | Should -Be ''
        }

        It 'still parses date, vendor, and amount correctly when Method is omitted' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Amazon Prime -$34.99'
            $r.Date   | Should -Be ([datetime]'2026-03-01')
            $r.Vendor | Should -Be 'Amazon Prime'
            $r.Amount | Should -Be '-34.99'
        }

        It 'still rejects an out-of-range month when Method is omitted' {
            $r = ConvertFrom-ReceiptFileName -Stem '261316 Amazon -$10.00'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Month out of range'
        }
    }

    Context '-DateFormat variations' {

        It 'parses yyyyMMdd format' {
            $r = ConvertFrom-ReceiptFileName -Stem '20260316 Sunoco $5.27 Card 9080' -DateFormat 'yyyyMMdd'
            $r.OK     | Should -Be $true
            $r.Date   | Should -Be ([datetime]'2026-03-16')
            $r.Vendor | Should -Be 'Sunoco'
        }

        It 'parses yy-MM-dd format' {
            $r = ConvertFrom-ReceiptFileName -Stem '26-03-16 CVS -$12.00 Cash' -DateFormat 'yy-MM-dd'
            $r.OK     | Should -Be $true
            $r.Date   | Should -Be ([datetime]'2026-03-16')
            $r.Method | Should -Be 'Cash'
        }

        It 'returns ParseOK=false and ParseError for out-of-range month in yyyyMMdd' {
            $r = ConvertFrom-ReceiptFileName -Stem '20261316 Amazon -$10.00 Card 1234' -DateFormat 'yyyyMMdd'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Month out of range'
        }

        It 'returns ParseOK=false and ParseError for out-of-range day in yy-MM-dd' {
            $r = ConvertFrom-ReceiptFileName -Stem '26-03-32 Amazon -$10.00 Card 1234' -DateFormat 'yy-MM-dd'
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Day out of range'
        }

        It 'returns ParseError="" on a successful parse' {
            $r = ConvertFrom-ReceiptFileName -Stem '260316 Sunoco $5.27 Card 9080'
            $r.OK         | Should -Be $true
            $r.ParseError | Should -Be ''
        }

        It 'parses M-d-yy format with single-digit month and day' {
            $r = ConvertFrom-ReceiptFileName -Stem '3-1-26 CVS $12.00 Cash' -DateFormat 'M-d-yy'
            $r.OK     | Should -Be $true
            $r.Date   | Should -Be ([datetime]'2026-03-01')
            $r.Vendor | Should -Be 'CVS'
        }

        It 'parses M-d-yy format with double-digit month and day' {
            $r = ConvertFrom-ReceiptFileName -Stem '12-31-26 CVS $12.00 Cash' -DateFormat 'M-d-yy'
            $r.OK   | Should -Be $true
            $r.Date | Should -Be ([datetime]'2026-12-31')
        }
    }

    Context '-Methods parameter' {

        It 'accepts a custom token when it is included in -Methods' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Shop -$10.00 Crypto 1234' -Methods @('Card', 'Crypto')
            $r.OK     | Should -Be $true
            $r.Method | Should -Be 'Crypto'
        }

        It 'rejects a token not in -Methods even if it is a valid word' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Shop -$10.00 Wire 1234' -Methods @('Card', 'Checking')
            $r.OK         | Should -Be $false
            $r.ParseError | Should -Be 'Unrecognised method'
        }

        It 'always accepts Cash regardless of -Methods' {
            $r = ConvertFrom-ReceiptFileName -Stem '260301 Shop -$5.00 Cash' -Methods @('Card')
            $r.OK     | Should -Be $true
            $r.Method | Should -Be 'Cash'
        }
    }
}
