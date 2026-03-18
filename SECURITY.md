# Security Policy

## Supported Versions

Only the latest release is actively maintained.

| Version | Supported |
|---------|-----------|
| 0.5.x   | Yes       |
| < 0.5   | No        |

## Reporting a Vulnerability

Please do not report security vulnerabilities via public GitHub issues.

Use [GitHub's private vulnerability reporting](https://github.com/ndomako10/Sync-Receipts/security/advisories/new)
to submit a report. You will receive a response within 7 days.

## Security Considerations

This script operates on personal financial data:

- `Accounts.xlsx` and per-year workbooks (`2026.xlsx`, etc.) contain account numbers and transaction history
- The `-ReceiptsRoot` parameter should always point to a trusted, access-controlled directory
- Receipt filenames are parsed via regex -- do not process filenames from untrusted sources
- Excel COM runs as the current user -- ensure the workbook path is not writable by other users
