# ADR-002: Encode Receipt Metadata in the Filename

## Status
Accepted

## Context
Each receipt file needs to carry date, vendor, amount, payment method, and account
information so the sync script can populate workbook rows without a separate database
or manual data entry step. The metadata must survive file copies and moves across
network shares.

## Decision
Encode all metadata directly in the receipt filename using the format:

    yyMMdd Vendor $Amount [Method [Account]].ext

`yyMMdd` is a .NET ParseExact format string; `Method` and `Account` are optional.
The script parses this with a regex in `ConvertFrom-ReceiptFileName`.

## Alternatives considered
- **Sidecar JSON files** -- a `260316 Vendor $5.00 Card 1234.json` alongside each
  receipt. Rejected as the initial design because it doubles the file count, complicates
  moves/copies (two files must travel together), and requires a separate write step
  after each receipt is filed. Sidecar files are being reconsidered for issue #36
  (OCR-driven category pre-population and split receipts).
- **Database / SQLite** -- a central store keyed by filename or hash. Rejected because
  it adds a dependency not present on the target machine, requires a sync/migration
  story, and breaks the "receipts are self-contained files on a share" model.
- **Extended file attributes / NTFS ADS** -- metadata in alternate data streams.
  Rejected because ADS are stripped when files are copied across network shares or
  to non-NTFS volumes.
- **Excel workbook as the sole record** -- enter data directly into the workbook without
  parsing filenames. Rejected because it requires manual entry for every receipt and
  loses the benefit of the filename as a durable, portable record.

## Consequences
- Receipt files must be named correctly before syncing; malformed filenames produce
  parse errors flagged in the workbook's Flag column.
- The date format is configurable via `-DateFormat` (default `yyMMdd`) to accommodate
  different naming preferences (issue #25).
- Method and Account are optional since v1.0.0 (issue #28); missing Method rows are
  flagged rather than rejected.
- Split receipts (multiple categories or payment methods) cannot be expressed in a
  single filename; this remains an open limitation tracked in issues #17 and #23.
