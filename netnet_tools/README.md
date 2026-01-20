# Net-Net Analysis Workbook Tools

Tools for validating and updating net-net stock analysis Excel workbooks with data from Investing.com Pro exports.

## Requirements

- Python 3.10+
- openpyxl

```bash
pip install openpyxl
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
