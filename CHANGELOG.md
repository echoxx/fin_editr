# Changelog

## 2026-01-29

### Added - Cell Reference Traceability for Autoscoring

Added the ability to trace autoscored values back to their source cells in the workbook.

#### New Features:
- **CellRef dataclass**: Tracks exact source cell (sheet, row, column, value, date) for each value used in scoring
- **Reference columns**: New columns M (Piotroski) and J (C7) showing source cell addresses
- **Date columns**: New columns N (Piotroski) and K (C7) showing the reporting period(s) used
- **Live formulas**: Single-value references are Excel formulas (e.g., `=ncav!AK34`) that update automatically

#### Updated Methods:
- `WorkbookEvaluator.get_date_for_column()` - Gets date from header row
- `WorkbookEvaluator.get_latest_cell_ref()` - Returns CellRef for most recent value
- `WorkbookEvaluator.get_latest_n_cell_refs()` - Returns CellRefs for N most recent values
- All scoring methods now populate `source_refs` field in ScoreResult

#### Example Output:
| Score | Notes | Reference | Period |
|-------|-------|-----------|--------|
| 0 | AUTO: CR: 3.41x -> 3.41x (↓) | ncav!AJ26; ncav!AK26 | 2025-03-31 → 2025-06-30 |
| 1 | AUTO: D/E = 0.04 (threshold: 0.5) | =ncav!AK53 | 2025-06-30 |
