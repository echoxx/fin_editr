# Claude Code Prompt: Net-Net Auto-Scoring (Phase 3)

I need you to build a Python script that auto-populates quantitative Piotroski F-Score and C7 criteria in my net-net valuation Excel workbooks, based on computed financial data.

## Context

I have Excel workbooks for analyzing net-net stocks with this structure:

**Raw data sheets (populated by separate automation):**
- `[company]_IS` - Income Statement from Investing.com (data starts row 12, columns D onwards)
- `[company]_bs` - Balance Sheet from Investing.com (data starts row 12, columns D onwards)

**Calculation sheets (formulas auto-compute from raw data):**
- `ncav` - NCAV, debt ratios, per-share metrics
- `ro` - ROA, ROE calculations
- `profitability` - Margins, growth rates

**Scoring sheets (currently manual - this is what we're automating):**
- `piotrosky` - Piotroski F-Score (9 criteria, scores in column K, rows 1-9)
- `C7` - Evan Bleker's checklist (scores in column H; Core criteria rows 2-10, Ranking criteria rows 14-22)

## Data locations in my template

```python
DATA_LOCATIONS = {
    'ncav': {
        'shares_outstanding': 20,      # Row 20, values in columns D onwards
        'net_current_assets': 23,      # Row 23
        'current_ratio': 26,           # Row 26
        'ncav_per_share': 34,          # Row 34
        'ncav_qoq_change': 35,         # Row 35 (percentage)
        'ncav_yoy_change': 36,         # Row 36 (percentage)
        'debt_to_assets': 52,          # Row 52
        'debt_to_equity': 53,          # Row 53
        'liabilities_to_equity': 54,   # Row 54
    },
    'profitability': {
        'sales': 2,                    # Row 2
        'gross_margin': 7,             # Row 7 (percentage)
        'net_income': 10,              # Row 10
        'ni_margin': 11,               # Row 11 (percentage)
    },
    'ro': {
        'roa_semi': 14,                # Row 14 (semi-annual ROA)
        'roa_annual': 21,              # Row 21 (annualized ROA)
        'asset_turnover': 33,          # Row 33 (Sales/Assets)
    },
    'piotrosky': {
        'score_column': 'K',           # Column K for scores
        'total_row': 10,               # Row 10 for total
    },
    'c7': {
        'score_column': 'H',           # Column H for scores
        'core_rows': (2, 10),          # Rows 2-10 for core criteria
        'ranking_rows': (14, 22),      # Rows 14-22 for ranking criteria
        'core_total_row': 11,
        'ranking_total_row': 23,
    }
}
```

## Core Python patterns to use

### Pattern 1: Dual workbook approach (read values, preserve formulas)

```python
from openpyxl import load_workbook

def load_workbook_dual(filepath):
    """
    Load workbook twice:
    - data_only=True to read calculated values
    - data_only=False to write scores without destroying formulas
    """
    wb_values = load_workbook(filepath, data_only=True)
    wb_formulas = load_workbook(filepath, data_only=False)
    return wb_values, wb_formulas
```

### Pattern 2: Extract time series from a row with missing data handling

```python
def get_row_series(sheet, row_num, start_col=4, max_cols=20):
    """
    Extract values from a row, handling None/empty cells.
    Returns list of (column_index, value) tuples for non-None values.
    Column 4 = D, Column 5 = E, etc.
    """
    series = []
    for col in range(start_col, start_col + max_cols):
        val = sheet.cell(row=row_num, column=col).value
        if val is not None and val != '' and val != 'NA':
            # Handle string numbers and percentages
            if isinstance(val, str):
                val = val.replace('%', '').replace(',', '').strip()
                try:
                    val = float(val)
                except ValueError:
                    continue
            series.append((col, val))
    return series

def get_latest_n_values(series, n=2):
    """
    Get the N most recent values from a series.
    Assumes rightmost columns are most recent.
    Returns list of values in chronological order (oldest first).
    """
    if len(series) < n:
        return None
    # Take last n values, return in chronological order
    return [v for (col, v) in series[-n:]]
```

### Pattern 3: Trend detection with edge case handling

```python
def is_increasing(series, periods=2):
    """
    Check if metric is increasing over the specified periods.
    Returns: True (increasing), False (not increasing), None (insufficient data)
    """
    values = get_latest_n_values(series, periods)
    if values is None:
        return None
    # Compare most recent to prior
    return values[-1] > values[-2]

def is_decreasing(series, periods=2):
    """
    Check if metric is decreasing over the specified periods.
    """
    values = get_latest_n_values(series, periods)
    if values is None:
        return None
    return values[-1] < values[-2]

def is_stable_or_decreasing(series, periods=2):
    """
    Check if metric is stable or decreasing (for share dilution check).
    """
    values = get_latest_n_values(series, periods)
    if values is None:
        return None
    return values[-1] <= values[-2]

def is_positive(series):
    """
    Check if most recent value is positive.
    """
    if not series:
        return None
    return series[-1][1] > 0

def get_latest_value(series):
    """
    Get the most recent value from a series.
    """
    if not series:
        return None
    return series[-1][1]
```

### Pattern 4: Safe percentage change calculation

```python
def pct_change(old_val, new_val):
    """
    Calculate percentage change, handling division by zero.
    Returns None if calculation not possible.
    """
    if old_val is None or new_val is None:
        return None
    if old_val == 0:
        return None if new_val == 0 else float('inf')
    return (new_val - old_val) / abs(old_val)

def yoy_change_from_series(series):
    """
    Calculate YoY change from a semi-annual series.
    Compares most recent to 2 periods prior (1 year for semi-annual).
    """
    if len(series) < 3:
        return None
    # For semi-annual data, YoY means comparing period N to period N-2
    values = [v for (col, v) in series]
    return pct_change(values[-3], values[-1])
```

### Pattern 5: Scoring logic with confidence tracking

```python
from dataclasses import dataclass
from typing import Optional

@dataclass
class ScoreResult:
    criterion: str
    score: Optional[int]  # 0, 1, or None if can't determine
    confidence: str       # 'high', 'medium', 'low', 'manual_required'
    reasoning: str        # Explanation of how score was determined
    values_used: dict     # The actual values used in determination

def score_positive_earnings(profitability_sheet) -> ScoreResult:
    """Piotroski criterion 1: Positive net income"""
    series = get_row_series(profitability_sheet, row_num=10)  # NI row
    latest = get_latest_value(series)
    
    if latest is None:
        return ScoreResult(
            criterion="Positive earnings",
            score=None,
            confidence="manual_required",
            reasoning="Could not find net income data",
            values_used={}
        )
    
    score = 1 if latest > 0 else 0
    return ScoreResult(
        criterion="Positive earnings",
        score=score,
        confidence="high",
        reasoning=f"Net income = {latest:,.0f} ({'positive' if score else 'negative/zero'})",
        values_used={"net_income": latest}
    )

def score_increasing_roa(ro_sheet) -> ScoreResult:
    """Piotroski criterion 3: Increasing ROA"""
    series = get_row_series(ro_sheet, row_num=21)  # Annualized ROA row
    
    if len(series) < 2:
        return ScoreResult(
            criterion="Increasing ROA",
            score=None,
            confidence="manual_required",
            reasoning="Insufficient ROA data for trend",
            values_used={}
        )
    
    values = get_latest_n_values(series, 2)
    increasing = values[-1] > values[-2]
    
    return ScoreResult(
        criterion="Increasing ROA",
        score=1 if increasing else 0,
        confidence="high",
        reasoning=f"ROA: {values[-2]:.2%} -> {values[-1]:.2%} ({'increasing' if increasing else 'decreasing'})",
        values_used={"roa_prior": values[-2], "roa_current": values[-1]}
    )

def score_ncav_burn_rate(ncav_sheet, threshold=-0.10) -> ScoreResult:
    """C7 criterion: Low NCAV burn rate (YoY decline < threshold)"""
    # Try to get YoY change directly from row 36
    yoy_series = get_row_series(ncav_sheet, row_num=36)
    
    if yoy_series:
        latest_yoy = get_latest_value(yoy_series)
        if latest_yoy is not None:
            # Note: positive = growing, negative = declining
            score = 1 if latest_yoy > threshold else 0
            return ScoreResult(
                criterion="Low NCAV burn rate",
                score=score,
                confidence="high",
                reasoning=f"NCAV YoY change: {latest_yoy:.1%} (threshold: {threshold:.0%})",
                values_used={"ncav_yoy_change": latest_yoy}
            )
    
    # Fallback: calculate from NCAV per share series
    ncav_series = get_row_series(ncav_sheet, row_num=34)
    yoy_change = yoy_change_from_series(ncav_series)
    
    if yoy_change is None:
        return ScoreResult(
            criterion="Low NCAV burn rate",
            score=None,
            confidence="manual_required",
            reasoning="Could not calculate NCAV trend",
            values_used={}
        )
    
    score = 1 if yoy_change > threshold else 0
    return ScoreResult(
        criterion="Low NCAV burn rate",
        score=score,
        confidence="medium",
        reasoning=f"Calculated NCAV YoY change: {yoy_change:.1%}",
        values_used={"ncav_yoy_change_calculated": yoy_change}
    )
```

## Scoring criteria to implement

### Piotroski F-Score (9 criteria)

| # | Criterion | Logic | Implementation |
|---|-----------|-------|----------------|
| 1 | Positive earnings | NI > 0 | `profitability` row 10, latest value > 0 |
| 2 | Positive OCF | OCF > 0 | If not available, flag manual. Could proxy with NI + D&A if desperate |
| 3 | Increasing ROA | ROA_t > ROA_t-1 | `ro` row 21, compare latest 2 periods |
| 4 | OCF > NI (accruals) | Quality check | Flag manual if OCF not available |
| 5 | Decreasing LT debt ratio | (Debt/Assets)_t < (Debt/Assets)_t-1 | `ncav` row 52, check if decreasing |
| 6 | Increasing current ratio | CR_t > CR_t-1 | `ncav` row 26, check if increasing |
| 7 | No share dilution | Shares_t <= Shares_t-1 | `ncav` row 20, check stable/decreasing |
| 8 | Increasing gross margin | GM_t > GM_t-1 | `profitability` row 7, check if increasing |
| 9 | Increasing asset turnover | AT_t > AT_t-1 | `ro` row 33, check if increasing |

### C7 Core Criteria (9 criteria)

| # | Criterion | Auto? | Logic |
|---|-----------|-------|-------|
| 1 | Not majority Chinese | Partial | Check Overview sheet country field, flag for review |
| 2 | Low P/NCAV (< 0.67) | With price | `ncav` row 34 latest ÷ price input |
| 3 | Low Debt/Equity | Yes | `ncav` row 53 latest < 0.5 |
| 4 | Adequate past earnings | Yes | `profitability` row 11, any period > 5% margin |
| 5 | Past price above NCAV | No | Requires historical price data |
| 6 | Existing operations | No | Qualitative - flag manual |
| 7 | Not selling shares | Yes | `ncav` row 20, stable/decreasing trend |
| 8 | Market cap < $50mm | With price | price × shares, check < 50M |
| 9 | Low NCAV burn rate | Yes | YoY decline < 10-15% (use -0.10 threshold) |

### C7 Ranking Criteria (9 criteria)

| # | Criterion | Auto? | Logic |
|---|-----------|-------|-------|
| 1 | Current ratio > 1.5x | Yes | `ncav` row 26 latest > 1.5 |
| 2 | Not financial/RE/fund | Partial | Check Overview, flag for review |
| 3 | Buying back stock | Yes | `ncav` row 20, decreasing (not just stable) |
| 4 | Low P/Net Cash | With price | Requires net cash calc and price |
| 5 | Insider ownership > 10% | Partial | Check ownership sheet if populated |
| 6 | Insider buys > sells | No | Requires transaction data |
| 7 | Positive burn rate | Yes | NCAV YoY change > 0 (growing) |
| 8 | Reasonable insider pay | No | Qualitative |
| 9 | Dividend yield | With price | Requires dividend and price data |

## Script requirements

### Command line interface

```python
import argparse

def main():
    parser = argparse.ArgumentParser(description='Auto-score net-net Piotroski and C7 criteria')
    parser.add_argument('workbook', help='Path to the Excel workbook')
    parser.add_argument('--price', type=float, help='Current share price for P/NCAV calculations')
    parser.add_argument('--output', help='Output path (default: modify in place)')
    parser.add_argument('--report-only', action='store_true', help='Print report without modifying workbook')
    parser.add_argument('--burn-threshold', type=float, default=-0.10, help='NCAV burn rate threshold (default: -0.10)')
    
    args = parser.parse_args()
    # ... implementation
```

### Output behavior

1. **Write scores to workbook:**
   - Piotroski scores go in column K, rows 1-9
   - C7 Core scores go in column H, rows 2-10
   - C7 Ranking scores go in column H, rows 14-22

2. **Add notes in adjacent column:**
   - Column L for Piotroski, Column I for C7
   - Format: "AUTO: [reasoning]" or "MANUAL: [what's needed]"

3. **Print summary report:**
```
=== PIOTROSKI F-SCORE ===
1. Positive earnings:     1  (AUTO) NI = 125.3mm
2. Positive OCF:          -  (MANUAL) OCF data not available
3. Increasing ROA:        1  (AUTO) 4.2% -> 5.1%
...
Auto-scored: 7/9 | Manual required: 2/9 | Total: 6/9

=== C7 CORE CRITERIA ===
...
```

4. **Create backup before modifying:**
   - Save original as `[filename]_backup_[timestamp].xlsx`

## Important edge cases to handle

1. **Sheet name variations:** The IS/BS sheets may be named `[company]_IS` or `[company]_is` (case variation). Use case-insensitive matching.

2. **Empty calculation sheets:** If formulas haven't been recalculated, values will be None. Detect this and warn user to run recalc.py first.

3. **Negative denominators:** When calculating ratios with potentially negative equity, handle appropriately.

4. **Semi-annual vs annual data:** The data is semi-annual. For YoY comparisons, compare period N to period N-2.

5. **String values in cells:** Some cells may contain "NA", "-", or other non-numeric strings. Handle gracefully.

6. **Missing sheets:** Not all workbooks may have all sheets. Check for sheet existence before accessing.

## Test file

Use this workbook: carmate.xlsx

## Development approach

1. First, write a diagnostic function that reads all relevant data points and prints them, to verify you're pulling the right values.

2. Then implement one scoring function at a time, testing each.

3. Finally, wire up the full script with CLI and file I/O.

Start by examining the workbook structure to confirm the row numbers match what I've specified, then proceed with implementation.
