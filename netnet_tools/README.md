# Net-Net Analysis Workbook Tools

Tools for validating and updating net-net stock analysis Excel workbooks with data from Investing.com Pro exports.

## Quick Start: Automated Workflow

The easiest way to analyze a new stock is with the automated workflow:

```bash
# One-time setup: configure Investing.com Pro credentials
python netnet_main.py setup-credentials

# Analyze a stock (downloads financials, creates workbook, populates data)
python netnet_main.py analyze 7859 --exchange TYO

# With price for autoscoring
python netnet_main.py analyze 7859 --exchange TYO --price 879
```

This will:
1. Log into your Investing.com Pro account
2. Search for the company and download Income Statement & Balance Sheet
3. Create a company subfolder (e.g., `/toso/`)
4. Copy an existing workbook as a template
5. Update it with the new company's financial data
6. Populate the Overview tab (Ticker, Country, Currency)

## Requirements

- Python 3.10+
- openpyxl
- playwright (for automated downloads)
- keyring (for secure credential storage)

```bash
pip install openpyxl playwright keyring
playwright install chromium
```

## Tools Overview

### netnet_validator.py

Maps workbook structure and validates new data against existing workbooks.

**Commands:**

1. **map** - Extract structure from an existing workbook to JSON
2. **validate** - Check if new export data is compatible with existing structure
3. **diff** - Show what would change if new data were applied

### netnet_updater.py

Safely updates workbooks with new Investing.com export data.

**Features:**
- Creates timestamped backups before any changes
- Never overwrites formulas - only updates raw data cells
- Validates structure compatibility before updating
- Can extend formulas to new period columns

## Workbook Structure

The tools expect workbooks with this structure:

### Raw Data Sheets
- `[company]_IS` - Income Statement data
- `[company]_bs` - Balance Sheet data

Each sheet has:
- Company name in C2
- Sheet type label in C6 ("Income Statement" or "Balance Sheet")
- Period codes in row 8 (FH-19, FH-18, ..., FH)
- Period end dates in row 10 (2016-03-31, 2016-09-30, ...)
- Financial line items starting row 12, labels in column C, data in D onwards

### Calculation Sheets
- `ncav` - Net current asset value
- `profitability` - Margins and growth
- `ro` - Return calculations
- `piotrosky` - F-score
- `C7` - Checklist

These sheets use `_xlfn.NUMBERVALUE()` formulas to pull from raw data sheets.

## Usage Examples

### 1. Map an existing workbook structure

```bash
python netnet_validator.py map --workbook almedio.xlsx --output almedio_structure.json
```

This creates a JSON file documenting:
- Sheet names and company name
- Row-to-label mappings for IS and BS
- Column-to-date mappings
- Period codes

### 2. Validate new data before updating

```bash
python netnet_validator.py validate \
    --structure almedio_structure.json \
    --is-file "Almedio Inc - Income Statement.xlsx" \
    --bs-file "Almedio Inc - Balance Sheet.xlsx"
```

This checks:
- Row labels match expected positions
- Reports any missing or new rows
- Shows new periods available in export

### 3. Generate a diff report

```bash
python netnet_validator.py diff \
    --workbook almedio.xlsx \
    --is-file "Almedio Inc - Income Statement.xlsx" \
    --bs-file "Almedio Inc - Balance Sheet.xlsx" \
    --verbose
```

Shows exactly what values would change without making any modifications.

### 4. Preview update (dry run)

```bash
python netnet_updater.py \
    --workbook almedio.xlsx \
    --is-file "Almedio Inc - Income Statement.xlsx" \
    --bs-file "Almedio Inc - Balance Sheet.xlsx" \
    --dry-run --verbose
```

### 5. Actually update the workbook

```bash
python netnet_updater.py \
    --workbook almedio.xlsx \
    --is-file "Almedio Inc - Income Statement.xlsx" \
    --bs-file "Almedio Inc - Balance Sheet.xlsx"
```

A timestamped backup is automatically created (e.g., `almedio_backup_20250118_143022.xlsx`).

### 6. Update and add new periods

```bash
python netnet_updater.py \
    --workbook almedio.xlsx \
    --is-file "Almedio Inc - Income Statement.xlsx" \
    --bs-file "Almedio Inc - Balance Sheet.xlsx" \
    --extend-periods
```

This will:
- Add new date columns to IS and BS sheets
- Extend formulas in calculation sheets

### 7. Force update (bypass validation)

```bash
python netnet_updater.py \
    --workbook almedio.xlsx \
    --is-file "Almedio Inc - Income Statement.xlsx" \
    --force
```

Use with caution - only when you know the structure differences are acceptable.

## Expected Row Positions

### Income Statement
| Row | Label |
|-----|-------|
| 12 | Revenue |
| 15 | Cost of Revenues |
| 16 | Gross Profit |
| 24 | Operating Income |
| 33 | Net Income to Stockholders |
| 40 | Weighted Average Basic Shares Out. |
| 43 | EBITDA |
| 44 | EBIT |

### Balance Sheet
| Row | Label |
|-----|-------|
| 12 | Cash And Equivalents |
| 13 | Short Term Investments |
| 14 | Accounts Receivable, Net |
| 15 | Inventory |
| 18 | Total Current Assets |
| 24 | Goodwill |
| 25 | Other Intangibles |
| 27 | Total Assets |
| 32 | Current Portion of LT Debt |
| 35 | Total Current Liabilities |
| 37 | Long-term Debt |
| 40 | Total Liabilities |
| 51 | Total Equity |

## Safety Features

1. **Backups** - Timestamped backup created before any modification
2. **Formula protection** - Never overwrites cells containing formulas
3. **Validation** - Structure checked before updates
4. **Dry run** - Preview changes without modification
5. **Detailed logging** - All operations logged for audit
6. **Fail-fast** - Stops on unexpected conditions

## Output Files

- `[company]_structure.json` - Structure mapping
- `[company]_backup_[timestamp].xlsx` - Backup before update
- Optional JSON reports for validation and diff results

## Troubleshooting

### "Row X mismatch" error
The row label in the export doesn't match the expected label. This usually means:
- Investing.com changed their export format
- You're using an export from a different data source
- The workbook template was modified

Use `--force` to bypass, or update the structure JSON.

### "Could not find Income Statement sheet"
The workbook must have sheets ending in `_IS` and `_bs`. Rename sheets if needed.

### Formula references broken after adding columns
The formula extension logic handles most patterns, but complex formulas may need manual adjustment. Always verify calculation sheets after adding new periods.

---

## Automated Workflow Tools

### netnet_main.py

Main CLI entry point for the automated workflow.

**Commands:**

```bash
# Set up Investing.com Pro credentials (one-time)
python netnet_main.py setup-credentials

# Verify credentials are stored correctly
python netnet_main.py verify-credentials

# Analyze a new stock
python netnet_main.py analyze TICKER [options]

# List existing company analyses
python netnet_main.py list
```

**Analyze Options:**

| Option | Description |
|--------|-------------|
| `--exchange`, `-e` | Exchange code (TYO, NYSE, etc.) - helps disambiguate tickers |
| `--price`, `-p` | Stock price for autoscoring |
| `--skip-download` | Use existing files instead of downloading |
| `--folder`, `-f` | Path to folder with existing export files |
| `--no-headless` | Show browser window (useful for debugging) |
| `--dry-run` | Show what would be done without making changes |

**Examples:**

```bash
# Full automation
python netnet_main.py analyze 7859 --exchange TYO

# Use already-downloaded files
python netnet_main.py analyze 7859 --skip-download --folder ./Toso

# Debug mode (see browser)
python netnet_main.py analyze 7859 --exchange TYO --no-headless
```

### credentials.py

Secure credential storage using OS keyring (Windows Credential Manager, macOS Keychain, Linux Secret Service).

```bash
# Interactive setup
python credentials.py setup

# Verify credentials exist
python credentials.py verify

# Delete stored credentials
python credentials.py delete
```

Alternatively, set environment variables:
```bash
export INVESTING_COM_EMAIL=your@email.com
export INVESTING_COM_PASSWORD=yourpassword
```

### investing_scraper.py

Playwright-based web automation for Investing.com Pro.

```bash
# Test scraper directly
python investing_scraper.py 7859 --exchange TYO
```

### file_manager.py

File and folder management utilities.

```bash
# Test file manager functions
python file_manager.py
```

### overview_populator.py

Populate Overview tab fields.

```bash
# Read current values
python overview_populator.py workbook.xlsx --read

# Populate fields
python overview_populator.py workbook.xlsx --ticker 7859 --exchange TYO

# Clear manual fields (for template prep)
python overview_populator.py workbook.xlsx --clear
```

### netnet_workflow.py

Workflow orchestrator (usually called via netnet_main.py).

```bash
# Run workflow directly
python netnet_workflow.py 7859 --exchange TYO
```

---

## Workflow Architecture

```
netnet_main.py (CLI Entry Point)
       │
       ▼
netnet_workflow.py (Orchestrator)
       │
       ├──► credentials.py (Secure credential storage)
       │
       ├──► investing_scraper.py (Web automation)
       │         • Login to Investing.com Pro
       │         • Search company by ticker
       │         • Download IS & BS Excel exports
       │
       ├──► file_manager.py (File operations)
       │         • Create company subfolder
       │         • Copy template workbook
       │         • Organize downloaded files
       │
       ├──► netnet_updater.py (Data update)
       │         • Rename sheets to new company prefix
       │         • Update formula references
       │         • Populate with financial data
       │
       ├──► overview_populator.py (Overview tab)
       │         • Set Ticker, Country, Currency
       │
       └──► netnet_autoscore.py (Scoring)
              • Score Piotroski & C7 criteria
```
