# ADR-009: Accounts.xlsx Schema Restructure

## Status
Accepted

> **Note:** This is a placeholder drafted from issue #57. The original ADR was authored on a
> separate machine and not yet pushed. Before replacing this file with the original:
>
> 1. **Compare the two files side by side** -- diff this placeholder against the original and
>    note any sections that exist in one but not the other (Context points, Alternatives,
>    Consequences, or schema column differences).
> 2. **Summarize the differences** -- identify whether the original contains additional
>    rationale, a different schema, or decisions that contradict what is here.
> 3. **Choose a merge strategy** -- suggested options:
>    - *Replace entirely* -- if the original is complete and consistent with this ADR, discard
>      this file and commit the original as-is.
>    - *Merge* -- if this placeholder contains content not present in the original (e.g.
>      finalized Method token set, filename grammar rules), incorporate it into the original
>      before committing.
>    - *Keep placeholder* -- if the original is incomplete or significantly out of date, retain
>      this file and close the matter.
> 4. **Remove this note block** once the final version is committed.

## Context

The original `Accounts.xlsx` schema stored Last 4, Network (Visa/Mastercard), and Type
(Debit/Credit) columns. As the tool has grown to support per-account spend summaries (#58)
and statement reconciliation (#60), several gaps became apparent:

1. No way to mark an account as inactive once a card is cancelled or replaced. The script
   cannot distinguish between an unknown account (data-entry error) and a retired account
   (expected but no longer in use).
2. No way to identify the balance-holding account a payment instrument draws from, which is
   needed to route statement reconciliation correctly.
3. The Network and Type columns add noise without aiding any script logic -- descriptive
   Account names make them redundant.
4. Some Last 4 values appear under multiple payment methods (e.g. the same Last 4 for both
   Checking and Wire). Without a Method column the script cannot disambiguate.
5. The filename Method token and Last 4 were both optional, allowing ambiguous and
   incomplete receipt filenames through without flagging.

## Decision

### Accounts table schema

| Column | Description | Notes |
|--------|-------------|-------|
| Last 4 | Account number as it appears in receipt filenames | Lookup key |
| Method | Payment method token | Only populated when the same Last 4 appears under multiple methods; blank when Last 4 is unambiguous. Uses the same token values as the filename. |
| Holder | Person the account belongs to | |
| Institution | Bank, card issuer, or wallet provider | e.g. Chase, PayPal, Cash App |
| Account | Name of the balance-holding account | Blank for standalone credit cards. Shared across instruments that draw from the same account (e.g. a debit card and its underlying checking account both carry the same Account name). |
| Status | `Active` or `Inactive` | User-managed dropdown validation |

Drop the Network and Type columns entirely.

### Filename Method tokens

The Method token in the filename is the payment method. Last 4 in the Accounts table uses
the same token values for disambiguation when needed.

| Token | Description | Last 4 required |
|-------|-------------|-----------------|
| `Cash` | Cash payment | No -- Cash never has an account number |
| `Card` | Debit or credit card | Yes |
| `Check` | Physical paper check | Yes |
| `Checking` | ACH, eCheck, direct debit, online bill pay from a checking account | Yes |
| `Savings` | Electronic transfer from a savings account | Yes |
| `Wire` | Bank wire transfer | Yes |
| `Transfer` | P2P or digital wallet payment (PayPal, Venmo, Zelle, Cash App, etc.) -- Institution column provides platform specificity | Yes |

### Filename grammar rules

- `Cash` -- Method only; Last 4 never expected.
- All other methods -- Method + Last 4 required. Last 4 may be:
  - A 4-digit account number (e.g. `1234`)
  - `xxxx` -- account exists but number is intentionally redacted
  - `----` -- account genuinely unknown at time of filing
- Method without Last 4 (non-Cash) -- invalid; flagged as missing account.
- Last 4 without Method -- invalid; flagged as missing method.
- Neither Method nor Last 4 -- valid; flagged as no account recorded.

### Lookup behaviour

| Filename | Last 4 value | Lookup |
|----------|-------------|--------|
| `Card 1234` | Digits | Exact match on Method + Last 4 (or Last 4 alone if Method column is blank for that row) |
| `Card xxxx` | Redacted | No exact match; flagged as redacted account |
| `Card ----` | Unknown | No lookup; flagged as unknown account |
| `Cash` | Absent | No lookup |
| *(neither)* | Absent | No lookup; flagged as no account recorded |

## Alternatives considered

- **Keep Network and Type, add Status and Account** -- rejected because Network and Type
  are never read by the script and add maintenance burden.
- **Separate inactive-accounts list** -- rejected because it duplicates the account registry
  and requires keeping two files in sync.
- **Account Type column instead of Method** -- considered adding an `Account Type` column
  (Credit Card / Debit Card / Checking / etc.) distinct from the filename Method token.
  Rejected as redundant -- the filename token already conveys account type at the granularity
  needed for parsing and lookup. Institution provides additional specificity where needed.
- **Draws From column for debit cards** -- considered a foreign-key-style column pointing
  debit card rows to their underlying checking account. Rejected in favour of sharing the
  same Account name across both rows, which is simpler and achieves the same
  reconciliation routing.

## Consequences

- `Get-ValidAccounts` must return structured objects (Last 4, Method, Status, Account)
  rather than a flat list of strings, so `Write-MonthSheet` can flag inactive accounts
  distinctly from unrecognised ones.
- `ConvertFrom-ReceiptFileName` must enforce that non-Cash methods require a Last 4 value.
  This was already enforced for `Card`, `Checking`, and `Savings`. The new tokens `Check`,
  `Wire`, and `Transfer` follow the same rule -- no behaviour change for existing receipts.
- `Accounts.template.xlsx` must be updated to the new schema with Status dropdown validation
  and the expanded Method token set.
- Existing `Accounts.xlsx` files must be migrated manually: rename/remove columns and
  populate Account and Status values.
- Truly unrecognised accounts (not present in the sheet at all) continue to be flagged as
  before; inactive accounts get a distinct flag value.
