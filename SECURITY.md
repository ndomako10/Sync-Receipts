# Security Policy

## Supported Versions

Only the latest release is actively maintained.

| Version | Supported |
|---------|-----------|
| 4.x     | Yes       |
| < 4     | No        |

## Reporting a Vulnerability

Please do not report security vulnerabilities via public GitHub issues.

Use [GitHub's private vulnerability reporting](https://github.com/ndomako10/Sync-Receipts/security/advisories/new)
to submit a report. You will receive a response within 7 days.

## Security Considerations

This script operates on personal financial data:

- `Config\Accounts.xlsx` and per-year workbooks (`2026.xlsx`, etc.) contain account numbers and transaction history; both are gitignored
- `Config\Categories.json` may reflect personal spending habits; it is gitignored
- `Config\Config.ini` contains the `RECEIPTS_ROOT` path, which may reveal network share structure or usernames; it is gitignored and must not be committed
- The `-ReceiptsRoot` parameter should always point to a trusted, access-controlled directory
- Receipt filenames are parsed via regex -- do not process filenames from untrusted sources
- Excel COM runs as the current user -- ensure the workbook path is not writable by other users
