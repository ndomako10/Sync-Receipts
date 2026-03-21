# ADR-009: Accounts.xlsx Schema Restructure

## Status
Accepted

## Context

The original `Accounts.xlsx` schema stored Last 4, Network (Visa/Mastercard), and Type
(Debit/Credit) columns. As the tool has grown to support per-account spend summaries (#58)
and statement reconciliation (#60), several gaps became apparent:

1. **No inactive flag** -- no way to mark an account as inactive once a card is cancelled or
   replaced. The script cannot distinguish between an unknown account (data-entry error) and a
   retired account (expected but no longer in use).

2. **No balance-holding account** -- no way to identify which balance-holding account a payment
   instrument draws from, which is needed to route statement reconciliation correctly and build
   per-account spend summaries.

3. **Network and Type add noise** -- neither column is read by the script. Descriptive Account
   names make Type redundant, and Type was also being used inconsistently (Debit/Credit for
   cards, but Checking/Savings for shared Last 4 rows to serve as a method disambiguator).

4. **Shared Last 4 values** -- the same 4-digit number can appear under multiple payment
   methods (e.g. `Checking` and `Savings`), each drawing from a different balance-holding
   account. Account number alone is insufficient to identify the bank account.

5. **Apple Pay tokenization** -- Apple Pay generates a unique Device Account Number (DAN) per
   device per card. A physical debit card added to a phone and a watch produces three distinct
   4-digit numbers in receipt filenames (`1111`, `2222`, `3333`), all drawing from the same
   bank account, but all appearing as the original card number on bank statements. Receipt-side
   account numbers and statement-side account numbers therefore differ for the same transaction.

## Decision

### Accounts table schema

| Column | Description | Notes |
|--------|-------------|-------|
| Last 4 | Account number as it appears in receipt filenames | Lookup key |
| Method | Payment method token | Always populated. Uses the same token values as the filename. Entries that share a Last 4 value must each have a distinct Method to disambiguate (e.g. `Checking` and `Savings`). |
| Holder | Person the account belongs to | |
| Institution | Bank, card issuer, or wallet provider | e.g. First National Bank, Apple Pay Phone, OnePay |
| Account | Name of the balance-holding account this instrument draws from | e.g. `MyBank Checking`. Every row maps to exactly one balance-holding account. |
| Status | `Active` or `Inactive` | User-managed dropdown validation |

Drop the Network and Type columns entirely.

### Active/Inactive workflow

The Status column is user-managed. A card should remain **Active** until its final bank
statement has been reconciled (see #60), even if the physical card has been closed. This covers
the edge case where a card closed mid-month still produces a final statement the following
month. Once the final statement is reconciled, mark the card **Inactive**. No script logic
enforces this -- the workflow is the guard.

### Apple Pay alias handling deferred to #60

Rather than adding a Source Account column to map DANs back to their physical card numbers,
this mapping will be handled via separate opt-in config at reconciliation time (e.g.
`Config/Aliases.json`). This keeps the core schema clean for the majority of users who do not
use tokenized payment methods. Users with Apple Pay will have unmatched reconciliation entries
until #60 is built and the alias config is populated.

### Filename Method tokens

The Method token in the filename is the payment method. The Method column in Accounts.xlsx
uses the same token values.

| Token | Description | Last 4 required |
|-------|-------------|-----------------|
| `Cash` | Cash payment | No -- Cash never has an account number |
| `Card` | Debit or credit card | Yes |
| `Check` | Physical paper check | Yes |
| `Checking` | ACH, eCheck, direct debit, online bill pay from a checking account | Yes |
| `Savings` | Electronic transfer from a savings account | Yes |
| `Wire` | Bank wire transfer | Yes |
| `Transfer` | P2P or digital wallet payment (PayPal, Venmo, Zelle, Cash App, etc.) | Yes |

### Filename grammar rules

- `Cash` -- Method only; Last 4 never expected.
- All other methods -- Method + Last 4 required. Last 4 may be:
  - A 4-digit account number (e.g. `1234`)
  - `xxxx` -- account exists but number is intentionally redacted
  - `----` -- account genuinely unknown at time of filing
- Method without Last 4 (non-Cash) -- invalid; flagged as missing account.
- Last 4 without Method -- invalid; flagged as missing method.
- Neither Method nor Last 4 -- valid; not flagged; receipt is recorded with no account.

### Lookup behaviour

| Filename | Last 4 value | Lookup |
|----------|-------------|--------|
| `Card 1234` | Digits | Exact match on Method + Last 4 |
| `Card xxxx` | Redacted | No exact match; flagged as redacted account |
| `Card ----` | Unknown | No lookup; flagged as unknown account |
| `Cash` | Absent | No lookup |
| *(neither)* | Absent | No lookup; no flag |

## Alternatives considered

- **Keep Network and Type, add Status and Account** -- rejected because Network and Type
  are never read by the script and add maintenance burden.

- **Separate inactive-accounts list** -- rejected because it duplicates the account registry
  and requires keeping two files in sync.

- **Source Account column** -- a column mapping alias account numbers (e.g. Apple Pay DANs)
  back to their physical card number, allowing the script to follow the reference chain to
  resolve Account. Rejected because it solves an edge case (tokenized payments) that most
  users will never encounter, adding schema complexity and script logic for a minority use
  case. Deferred to #60 where it can be implemented as opt-in config.

- **Manual Account on every row (no alias handling)** -- populate Account on every row
  including Apple Pay aliases, accepting the redundancy. Rejected because it doesn't solve
  the reconciliation problem -- the reconciliation feature would still have no way to match a
  DAN in a receipt to the physical card number on a bank statement.

- **Account Type column instead of Method** -- considered adding an Account Type column
  (Credit Card / Debit Card / Checking / etc.) distinct from the filename Method token.
  Rejected as redundant -- the filename token already conveys account type at the granularity
  needed for parsing and lookup.

- **Draws From column for debit cards** -- considered a foreign-key-style column pointing
  debit card rows to their underlying checking account. Rejected in favour of sharing the
  same Account name across both rows, which is simpler and achieves the same reconciliation
  routing.

- **Method optional in Accounts.xlsx** -- considered making Method blank for unambiguous Last
  4 values and only populating it for shared Last 4 rows. Rejected because always populating
  Method makes the table self-describing and avoids a silent partial-match lookup that could
  incorrectly resolve an account when two rows share a Last 4 but only one has a Method.

## Consequences

- `Get-ValidAccounts` must return structured objects (Last 4, Method, Status, Account) rather
  than a flat list of strings, so `Write-MonthSheet` can flag inactive accounts distinctly
  from unrecognised ones.
- The Method column lookup matches on Method + Last 4.
- `ConvertFrom-ReceiptFileName` must enforce that non-Cash methods require a Last 4 value.
  Omitting both Method and Last 4 is valid and produces no flag.
- `Config/Templates/Accounts.template.xlsx` must be updated to the new schema with Status dropdown validation.
- Existing `Accounts.xlsx` files must be migrated manually: rename/remove columns and
  populate Account, Method, and Status values.
- When #60 is implemented, Apple Pay alias mapping should be handled via a separate opt-in
  config rather than by extending this schema.
- Truly unrecognised accounts (not present in the sheet at all) continue to be flagged as
  before; inactive accounts get a distinct flag value.
