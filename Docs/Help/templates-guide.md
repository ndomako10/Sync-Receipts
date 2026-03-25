# Template Initialization & Management

How templates work, when to regenerate them, and the one-time setup flow.

## What Are Templates?

Templates are "default" configuration files committed to git in `Config/Templates/`. They serve two purposes:

1. **Setup**: During `Initialize-SyncReceipts.ps1`, templates are copied to `Config/` for first-time setup
2. **Reference**: Templates remain in git as schema documentation for future schema changes

**Key distinction:**
- `Config/Templates/` -- committed to git; never edited for daily use
- `Config/` -- gitignored; contains user's actual configuration

---

## Templates Available

| Template File | Destination | Purpose |
|---------------|-------------|---------|
| `Config.template.ini` | `Config/Config.ini` | Machine-local settings (RECEIPTS_ROOT, etc.) |
| `Accounts.template.xlsx` | `Config/Accounts.xlsx` | Default accounts table; schema reference |
| `Categories.template.json` | `Config/Categories.json` | Default categories; hierarchical structure |
| `Methods.template.json` | `Config/Methods.json` | Default method tokens (Cash, CC, etc.) |
| `SensitivePatterns.template.json` | `Config/SensitivePatterns.json` | Default sensitive data detection patterns |

---

## One-Time Setup Flow

User runs `Setup.bat`:

```
Setup.bat
  |
  [`----> Scripts\Initialize-SyncReceipts.ps1
      |
      [|----> Verify PowerShell version & prerequisites
      |
      [|----> Copy Config.template.ini -> Config/Config.ini
      |   [`----> Prompts user for RECEIPTS_ROOT
      |
      [|----> Copy all templates from Config/Templates/ -> Config/
      |   [|----> Accounts.template.xlsx -> Accounts.xlsx
      |   [|----> Categories.template.json -> Categories.json
      |   [|----> Methods.template.json -> Methods.json
      |   [`----> SensitivePatterns.template.json -> SensitivePatterns.json
      |
      [|----> Scripts\Install-GitHooks.ps1
      |   [`----> Copies pre-commit, commit-msg, pre-push hooks to .git/hooks/
      |
      [`----> Create desktop shortcuts in RECEIPTS_ROOT for batch launchers
```

**Result:** User has working config files in `Config/` and can start syncing.

---

## Config.ini

### Template: `Config.template.ini`

```ini
RECEIPTS_ROOT=
WORKBOOKS_ROOT=
```

### First-Time Setup

1. `Initialize-SyncReceipts.ps1` prompts user:
   ```
   Enter the path to your Receipts folder:
   ```

2. Copies template and fills in RECEIPTS_ROOT:
   ```ini
   RECEIPTS_ROOT=C:\Users\YourName\Documents\Receipts
   WORKBOOKS_ROOT=
   ```

3. User can optionally edit to set WORKBOOKS_ROOT

### Updating

Users edit `Config/Config.ini` directly (not the template).

**Do NOT edit** `Config.template.ini` unless changing default setup instructions.

---

## Accounts.xlsx

### Template: `Accounts.template.xlsx`

Committed Excel workbook with empty table (headers only).

**Columns:**
- Last4
- Method
- Holder
- Institution
- Account
- Status

### First-Time Setup

1. `Initialize-SyncReceipts.ps1` copies template -> `Config/Accounts.xlsx`
2. User opens and adds their account records

### Updating Schema

When schema changes (e.g., adding "CreditLimit" column):

1. **Update template:**
   ```powershell
   Scripts\New-AccountsTemplate.ps1
   ```

   This:
   - Reads current `Config/Accounts.xlsx` (preserves user data)
   - Updates schema in `Config/Templates/Accounts.template.xlsx`
   - Does NOT overwrite user's `Config/Accounts.xlsx`

2. **Commit new template:**
   ```bash
   git add Config/Templates/Accounts.template.xlsx
   git commit -m "chore(config): add CreditLimit column to accounts schema"
   ```

3. **Users update manually** (on next setup or by editing existing workbook)

---

## Categories.json

### Template: `Categories.template.json`

```json
{
  "Food & Dining": ["Groceries", "Restaurants", "Delivery"],
  "Transportation": ["Gas", "Parking", "Public Transit"],
  "Utilities": ["Electric", "Gas", "Water", "Internet"]
}
```

### First-Time Setup

1. `Initialize-SyncReceipts.ps1` copies template -> `Config/Categories.json`
2. User opens and customizes categories

### Updating

Users edit `Config/Categories.json` directly.

**To ship new default categories to all users:**

1. Update `Config/Templates/Categories.template.json`
2. Commit: `git commit -m "chore(config): update default categories"`
3. **Note:** Existing users must manually copy from template (automatic copy only happens during setup)

---

## Methods.json

### Template: `Methods.template.json`

```json
["Cash", "CC", "DB", "AP", "VND"]
```

### First-Time Setup

1. `Initialize-SyncReceipts.ps1` copies template -> `Config/Methods.json`
2. User can customize if additional method tokens needed

### Updating

Users edit `Config/Methods.json` directly.

**To ship new default methods:**

1. Update `Config/Templates/Methods.template.json`
2. Commit: `git commit -m "chore(config): add CHK method token"`
3. **Note:** Existing users must manually copy from template

---

## SensitivePatterns.json

### Template: `SensitivePatterns.template.json`

```json
{
  "SSN": "\\b\\d{3}-\\d{2}-\\d{4}\\b",
  "Account Number": "\\b[0-9]{10,17}\\b",
  "Routing Number": "\\b\\d{9}\\b",
  "Credit Card": "\\b[0-9]{13,19}\\b",
  "Phone": "\\b\\d{3}[-.]?\\d{3}[-.]?\\d{4}\\b"
}
```

### First-Time Setup

1. `Initialize-SyncReceipts.ps1` copies template -> `Config/SensitivePatterns.json`
2. Pre-commit hook uses patterns to block commits with sensitive data

### Updating

Users edit `Config/SensitivePatterns.json` directly to:
- Add custom patterns (e.g., medical record numbers)
- Remove overly strict patterns
- Adjust regex to match specific formats

**To ship new default patterns:**

1. Update `Config/Templates/SensitivePatterns.template.json`
2. Commit: `git commit -m "chore(config): add medical record detection pattern"`
3. **Note:** Existing users must manually copy from template

---

## When to Regenerate Templates

### New-AccountsTemplate.ps1

**When:** Schema changes to Accounts.xlsx (adding/removing/renaming columns)

**How:**
```powershell
Scripts\New-AccountsTemplate.ps1
```

**What it does:**
- Reads user's current `Config/Accounts.xlsx`
- Extracts column schema
- Writes new `Config/Templates/Accounts.template.xlsx` with same schema
- Preserves user data in `Config/Accounts.xlsx` (does not modify)

**Why separately:** Excel files can't be easily version-controlled like JSON. This script ensures the template always reflects the current schema used by the application.

### Other Templates

For `Config.template.ini`, `Categories.template.json`, `Methods.template.json`, `SensitivePatterns.template.json`:
- Edit the template directly in `Config/Templates/`
- Commit the change
- Users are responsible for updating their own `Config/` copy if desired

**Example workflow:**
```bash
# Update default categories
nano Config/Templates/Categories.template.json
git add Config/Templates/Categories.template.json
git commit -m "chore(config): add new categories"
# Existing users still have old categories; they can manually copy from template if desired
```

---

## Troubleshooting

### "Setup created Config/Config.ini but it's blank"

**Causes:**
- User didn't respond to prompt
- Setup script crashed mid-way

**Fix:**
```bash
rm Config/Config.ini
Setup.bat  # Run again
```

### "Templates in Config/Templates/ are outdated"

**Note:** Templates are committed to git. They should always match the current schema.

**If templates are behind:**
```bash
# Check schema in your code
# Update template to match
nano Config/Templates/Categories.template.json

# Commit
git add Config/Templates/Categories.template.json
git commit -m "chore(config): sync template with code"

# Users should update their Config/ copies manually
```

### "I updated Config/Templates/X but my Config/X didn't change"

**This is expected.** Templates are only copied during initial setup (`Initialize-SyncReceipts.ps1`).

**To update your config:**
```bash
# Manual copy
cp Config/Templates/Categories.template.json Config/Categories.json

# Or delete and re-run setup
rm Config/Categories.json
Setup.bat
```

---

## Related

- [config-schema.md](config-schema.md) -- Contents of each config file
- [workflow.md](workflow.md) -- Initialize-SyncReceipts.ps1 in setup workflow
- [architecture-quick-ref.md](architecture-quick-ref.md) -- Config file locations
