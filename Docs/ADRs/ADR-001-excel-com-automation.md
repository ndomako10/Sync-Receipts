# ADR-001: Use Excel COM Automation for Workbook Generation

## Status
Accepted

## Context
The tool writes formatted, multi-sheet Excel workbooks with hyperlinks, structured
tables, named ranges, and dependent dropdown validation. Several approaches exist
for generating .xlsx files from PowerShell on Windows.

## Decision
Use Excel COM automation (`New-Object -ComObject Excel.Application`) as the primary
mechanism for creating and updating workbooks.

## Alternatives considered
- **OpenXML SDK / EPPlus** -- library-based .xlsx generation without requiring Excel.
  Rejected because the formatting requirements (structured tables, hyperlinks, named
  ranges, conditional dropdown validation) are easier to express through the Excel
  object model than raw XML. Also adds a dependency management burden on a machine
  with no package manager configured.
- **CSV export** -- simple and portable. Rejected because it cannot represent
  hyperlinks, formatted tables, or dropdown validation, all of which are core to
  the workbook's usability as a manual review tool.
- **Google Sheets API** -- cloud-based, no local Excel requirement. Rejected because
  receipts may contain sensitive financial data and the tool runs on a Windows machine
  with local/network file access; no Google account or internet dependency is acceptable.

## Consequences
- Requires Excel (2016+) installed on the machine running the script.
- Windows-only; not portable to macOS or Linux.
- Excel must be closed or the workbook unlocked before syncing; a hung Excel process
  requires Kill-Excel.bat to recover.
- COM interaction with hidden sheets can deadlock; workarounds are needed (e.g.
  `AskToUpdateLinks = $false`, `UpdateLinks = 0` on open, and the post-save XML
  patch for validation that COM cannot inject on hidden sheets).
