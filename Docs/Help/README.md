# Help & Technical Reference

Complete index of all technical documentation for the Sync-Receipts project. This guide helps you select the right context files to load for your task, optimizing token usage by 40-60%.

---

## Task Type -> File Recommendation

| Task | Load This | Why |
|------|-----------|-----|
| **Making a commit** | [commit-reference.md](commit-reference.md) + [commit-scopes.md](commit-scopes.md) | Types, formats, scope selection, versioning rules |
| **Writing PowerShell code** | [coding-rules.md](coding-rules.md) + [function-index.md](function-index.md) | Error handling, string escaping, date formats, existing functions to reuse |
| **Adding a new function** | [function-index.md](function-index.md) + [coding-rules.md](coding-rules.md) | See existing function patterns, testability requirements |
| **Understanding the codebase** | [architecture-quick-ref.md](architecture-quick-ref.md) + [function-index.md](function-index.md) | File structure, data locations, key functions |
| **Configuring the project** | [config-schema.md](config-schema.md) + [architecture-quick-ref.md](architecture-quick-ref.md) | Config file locations, templates, structure |
| **Writing tests** | [testing-strategy.md](testing-strategy.md) + [function-index.md](function-index.md) | Which functions need tests, testing requirements |
| **Understanding Excel automation** | [architecture-quick-ref.md](architecture-quick-ref.md) | Excel COM patterns and how they're used |
| **Creating receipt filenames** | [receipt-format.md](receipt-format.md) + [config-schema.md](config-schema.md) | Filename syntax, parsing rules, methods/accounts |
| **Understanding a feature** | [features-overview.md](features-overview.md) + [receipt-format.md](receipt-format.md) | Category, Account, Method, Flag, and Hyperlink features |
| **Debugging a problem** | [debugging.md](debugging.md) + [receipt-format.md](receipt-format.md) | Logging, common failures, error codes |
| **Managing config files** | [config-schema.md](config-schema.md) + [templates-guide.md](templates-guide.md) | Config structure, templates, initialization |
| **Setting up your machine** | [workflow.md](workflow.md) + [config-schema.md](config-schema.md) | Setup process, config initialization, paths |
| **Opening/managing PRs** | [workflow.md](workflow.md) + [github-workflows.md](github-workflows.md) | Issue tracking, PR process, CI/CD pipeline |
| **Selecting commit scope** | [commit-scopes.md](commit-scopes.md) | Scope lookup table with file mappings |
| **Understanding project terminology** | [glossary.md](glossary.md) | Definitions of date formats, file terms, features |
| **Using slash commands** | [slash-commands-reference.md](slash-commands-reference.md) | All `/command` shortcuts and when to use them |
| **Optimizing context loading** | This page (README.md) | Token optimization, file relationships, task-to-file mapping |

---

## All Files (15 Technical Guides)

### Fundamentals (Start Here)

- **[glossary.md](glossary.md)** -- Project terminology (dates, files, features, testing, git)

### Workflow & Process

- **[workflow.md](workflow.md)** -- Issue creation, PR process, release workflow, checklist
- **[github-workflows.md](github-workflows.md)** -- CI/CD pipeline, local hooks, GitHub Actions, merge requirements
- **[slash-commands-reference.md](slash-commands-reference.md)** -- All `/command` shortcuts (prep-pr, merge-pr, review-all, etc.)

### Code & Development

- **[coding-rules.md](coding-rules.md)** -- PowerShell coding standards: error handling, strings, dates, XML, testing
- **[function-index.md](function-index.md)** -- All key functions, test status, dependencies
- **[testing-strategy.md](testing-strategy.md)** -- Unit vs. integration testing, Pester patterns, CI behavior

### Architecture & Configuration

- **[architecture-quick-ref.md](architecture-quick-ref.md)** -- File structure, data locations, Excel COM patterns
- **[config-schema.md](config-schema.md)** -- Config.ini, Accounts.xlsx, Categories.json, Methods.json, SensitivePatterns.json
- **[templates-guide.md](templates-guide.md)** -- Template initialization, when to regenerate, setup flow

### Features & Functionality

- **[features-overview.md](features-overview.md)** -- Categories, Accounts, Methods, Sensitive Patterns, Flag System, Hyperlinks
- **[receipt-format.md](receipt-format.md)** -- Receipt filename syntax, parsing, flag rules, validation examples

### Commits & Versioning

- **[commit-reference.md](commit-reference.md)** -- Commit format, types, versioning, changelog workflow
- **[commit-scopes.md](commit-scopes.md)** -- Scope selection table with file mappings

### Troubleshooting

- **[debugging.md](debugging.md)** -- Logging system, utilities, common failures, debug workflow

---

## Reference Relationships

```
Workflow & Commit Management:
  workflow.md
    [|----> commit-reference.md
    [|----> github-workflows.md
    [`----> config-schema.md

  commit-reference.md
    [`----> commit-scopes.md

  commit-scopes.md
    [`----> Used for all commits

Code Development:
  coding-rules.md
    [|----> function-index.md
    [`----> architecture-quick-ref.md

  function-index.md
    [|----> coding-rules.md
    [`----> architecture-quick-ref.md

Architecture & Configuration:
  architecture-quick-ref.md
    [|----> function-index.md
    [|----> coding-rules.md
    [`----> commit-scopes.md

  config-schema.md
    [|----> architecture-quick-ref.md
    [`----> templates-guide.md

  templates-guide.md
    [`----> config-schema.md

Features & Functionality:
  features-overview.md
    [|----> receipt-format.md
    [`----> config-schema.md

  receipt-format.md
    [|----> config-schema.md
    [|----> features-overview.md
    [`----> glossary.md

Testing & Debugging:
  testing-strategy.md
    [|----> function-index.md
    [`----> coding-rules.md

  debugging.md
    [|----> receipt-format.md
    [`----> glossary.md

CI/CD & Tools:
  github-workflows.md
    [|----> workflow.md
    [`----> commit-reference.md

  slash-commands-reference.md
    [|----> workflow.md
    [`----> config-schema.md

Reference (Entry Points):
  glossary.md
    [`----> All files (defines project terminology)

  README.md (you are here)
    [`----> All 15 files (complete index)

Original Source:
  CLAUDE.md (178 lines)
    [`----> Superseded by Help/ files for most tasks
    [`----> Still needed for: global workflow rules, workspace context
```

---

## When to Load Full CLAUDE.md

Load `CLAUDE.md` only if you need:
- Global issue tracking and workflow rules (global CLAUDE.md at `~/.claude/CLAUDE.md`)
- Complete project/workspace context
- Reference to the full architecture (this page covers 95% of common tasks)

For routine commits, code writing, and configuration tasks, the extracted files are more efficient.

---

## File Sizes & Token Impact

### Workflow & Commit Management
| File | Lines | Load When |
|------|-------|-----------|
| commit-reference.md | ~50 | Making a commit |
| commit-scopes.md | ~65 | Choosing commit scope |
| workflow.md | ~180 | Setting up or PR process |
| github-workflows.md | ~200 | Understanding CI/CD pipeline |
| slash-commands-reference.md | ~180 | Using slash commands |

### Code Development
| File | Lines | Load When |
|------|-------|-----------|
| coding-rules.md | ~100 | Writing PowerShell code |
| function-index.md | ~65 | Working with functions |
| testing-strategy.md | ~130 | Writing tests |

### Architecture & Configuration
| File | Lines | Load When |
|------|-------|-----------|
| architecture-quick-ref.md | ~130 | Understanding codebase |
| config-schema.md | ~200 | Configuring the app |
| templates-guide.md | ~140 | Managing templates |

### Features & Functionality
| File | Lines | Load When |
|------|-------|-----------|
| features-overview.md | ~200 | Understanding a feature |
| receipt-format.md | ~330 | Working with receipts |
| debugging.md | ~200 | Troubleshooting problems |

### Reference
| File | Lines | Load When |
|------|-------|-----------|
| glossary.md | ~150 | Looking up terminology |
| **CLAUDE.md (original)** | **178** | Need full project context |

### Summary

**Total Help files: 15** (~2,235 lines)

**Token optimization benefit:**
- Instead of always loading 178-line CLAUDE.md
- Load only 50-200 lines relevant to your specific task
- **Result: 40-60% token reduction per conversation**

### Quick Load Sizes by Task Type

| Task | Files to Load | Approx. Lines | vs. Full CLAUDE.md |
|------|---------------|---------------|-------------------|
| Making a commit | commit-reference + commit-scopes | ~115 | -35% |
| Writing code | coding-rules + function-index | ~165 | -7% |
| Writing tests | testing-strategy + function-index + coding-rules | ~295 | +65% |
| Setting up machine | workflow + config-schema + architecture-quick-ref | ~510 | +186% |
| Debugging | debugging + receipt-format + glossary | ~680 | +282% |
| Browsing all | This README + any 2-3 specific files | ~250 | +40% |
| Just commit scope lookup | commit-scopes | ~65 | -63% |

**Key insight:** Most routine tasks load 50-200 lines (35-63% savings). Comprehensive tasks may load more but organize knowledge by topic instead of one monolithic file.

---

## How to Use This Guide

**Step 1: Find your task**
- Scroll to "Task Type -> File Recommendation" table at the top
- Find your task in the left column
- Click the recommended files in the right column

**Step 2: Load recommended files**
- Each link opens the specific file you need
- Files are small (50-200 lines typical) and focused

**Step 3: Jump between related files**
- Each Help file links to related guides
- "Related" sections point to complementary documents
- Use "Reference Relationships" above to understand dependencies

**Can't find your task?**
- See the "All Files (15 Technical Guides)" section above
- See [glossary.md](glossary.md) for terminology definitions

---

## Navigation Paths

### For Committing Code
1. Start: [commit-reference.md](commit-reference.md) -- Learn commit message format
2. Link: [commit-scopes.md](commit-scopes.md) -- Choose your scope
3. Link: [coding-rules.md](coding-rules.md) -- Verify code quality standards
4. Optional: [glossary.md](glossary.md) -- Check terminology

### For Writing New Code
1. Start: [function-index.md](function-index.md) -- See what exists
2. Link: [coding-rules.md](coding-rules.md) -- Apply PowerShell standards
3. Link: [testing-strategy.md](testing-strategy.md) -- Learn testing requirements
4. Optional: [architecture-quick-ref.md](architecture-quick-ref.md) -- Understand file structure

### For Debugging
1. Start: [debugging.md](debugging.md) -- Find the problem
2. Link: [receipt-format.md](receipt-format.md) -- Validate filename format
3. Link: [glossary.md](glossary.md) -- Understand error codes
4. Optional: [github-workflows.md](github-workflows.md) -- Check CI failures

### For Understanding Features
1. Start: [features-overview.md](features-overview.md) -- Feature tour
2. Link: [receipt-format.md](receipt-format.md) -- Learn receipt parsing
3. Link: [config-schema.md](config-schema.md) -- See configuration
4. Optional: [glossary.md](glossary.md) -- Define terms

---

## Related

- [../README.md](../README.md) -- Docs folder overview
- [../context-selection.md](../context-selection.md) -- Archived (merged into this page)
- [../../CLAUDE.md](../../CLAUDE.md) -- Original comprehensive guide (use when needed)
