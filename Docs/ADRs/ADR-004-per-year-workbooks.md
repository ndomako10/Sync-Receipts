# ADR-004: One Workbook per Year Instead of a Single Workbook

## Status
Accepted

## Context
Initially the script wrote all months into a single `Receipts.xlsx` workbook. As the
number of months grew, re-syncing the full file became slower and the workbook harder
to navigate. Excel also imposes practical limits on sheet count and file size for COM
performance. (Issue #7.)

## Decision
Create one workbook per calendar year (`2025.xlsx`, `2026.xlsx`, etc.) in
`WorkbooksRoot`. Each workbook contains one sheet per month (named by YYMM, e.g.
`2603`). A new workbook is created automatically on first sync for that year.

## Alternatives considered
- **Single workbook, all years** -- simple path resolution, one file to open for
  review. Rejected because it grows unboundedly, makes year-end archiving awkward,
  and COM performance degrades with large workbook sheet counts.
- **One workbook per month** -- finest granularity, smallest files. Rejected because
  it makes cross-month review require opening multiple files and named ranges
  (Category/Subcategory dropdowns) would need to be duplicated across every workbook.

## Consequences
- Switching years requires opening a different file; no single view across years.
- The `-Year` parameter (issue #18) and `-All` parameter support syncing multiple
  workbooks in one invocation.
- `WorkbooksRoot` (issue #31) can be set independently of `ReceiptsRoot`, allowing
  workbooks on a fast local drive while receipts remain on a network share.
