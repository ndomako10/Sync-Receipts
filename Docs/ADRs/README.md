# Architecture Decision Records

This directory contains Architecture Decision Records (ADRs) for the Sync-Receipts project.

ADRs capture *why* a significant decision was made -- the context, alternatives considered, and tradeoffs accepted. They are permanent records and are never deleted; superseded records are updated to reference their replacement.

## Index

| ADR | Title | Status |
|-----|-------|--------|
| [ADR-001](ADR-001-excel-com-automation.md) | Use Excel COM Automation for Workbook Generation | Accepted |
| [ADR-002](ADR-002-filename-as-metadata.md) | Encode Receipt Metadata in the Filename | Accepted |
| [ADR-003](ADR-003-post-save-xml-patch.md) | Inject Dropdown Validation via Post-Save XML Patch | Accepted |
| [ADR-004](ADR-004-per-year-workbooks.md) | One Workbook per Year Instead of a Single Workbook | Accepted |
| [ADR-005](ADR-005-config-folder-for-personal-data.md) | Store Personal Data Files in Config/ and Gitignore Them | Accepted |
| [ADR-006](ADR-006-separate-workbooks-root.md) | Add WorkbooksRoot Parameter to Separate Workbook Output from Receipts | Accepted |
| [ADR-007](ADR-007-xlsx-not-xlsm.md) | Keep Workbooks as .xlsx (Non-Macro-Enabled) | Accepted |

## Template

See the [global CLAUDE.md](~/.claude/CLAUDE.md) for the ADR template and lifecycle rules.
