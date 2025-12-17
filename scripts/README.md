# Haleon MEA Budget Sufficiency Tools

Tools for syncing Excel budget data with PowerPoint presentations.

## Quick Start - Sync Excel → PPT

### Option 1: Streamlit App (Recommended)
```bash
./run_app.sh
# Opens http://localhost:8501
# Upload Excel + PPT files, click Sync
```

### Option 2: Command Line
```bash
source .venv/bin/activate
python scripts/update_ppt_from_excel.py \
    --ppt "input/your_presentation.pptx" \
    --excel "output/your_budget.xlsx" \
    --output-dir output
```

## What Gets Synced

| Field | Excel Column | Description |
|-------|--------------|-------------|
| budget_2026 | E | 2026 Budget |
| sufficient_2026 | G | 2026 Sufficient |
| gap_gbp | H | Gap (GBP) |
| gap_pct | I | Gap (%) |
| awa, con, pur | J, K, L | CEJ allocations |
| tv, digital, others | R, S, V | Media % |
| long_campaigns | AD | Long campaign count |
| short_campaigns | AF | Short campaign count |
| long_pct | AH | Long campaign % |

## Markets Covered

| Market | Excel Rows | PPT Slides |
|--------|------------|------------|
| KSA | 5-11 | 22 |
| GNE | 13-18 | 28 |
| South Africa | 20-27 | 62 |
| Turkey | 29-34 | 68 |
| Pakistan | 36-40 | 56 |
| Egypt | 42-44 | 34 |
| Morocco | 46 | 40 |
| FSA | 48-50 | 51 |
| Kenya | 52-54 | 45 |
| Nigeria | 58-59 | 77 |

## Output Files

Each sync creates:
- `*_UPDATED_YYYYMMDD_HHMMSS.pptx` - Updated presentation
- `BACKUP_*_YYYYMMDD_HHMMSS.pptx` - Original backup
- `sync_report_YYYYMMDD_HHMMSS.json` - Change log

---

# Data Validation Framework

Comprehensive validation harness for verifying Excel data matches PPT source of truth.

## Quick Start

```bash
# Run full validation
python scripts/validator.py \
    --excel-data output/data/updated_excel_extracted.json \
    --ppt-data output/data/ppt_extracted_data.json \
    --output-dir output/reports

# Run adversarial tests (prove validator catches errors)
python scripts/adversarial_tests.py

# Run sampling strategy tests
python scripts/sampling_tests.py
```

## Architecture

```
scripts/
├── validator.py          # Core validation engine
├── adversarial_tests.py  # Tests that intentionally try to break validation
├── sampling_tests.py     # Sample-based validation across dimensions
├── extract_ppt_tables.py # PPT data extraction
├── extract_excel_data.py # Original Excel extraction
├── extract_updated_excel.py # Updated Excel extraction
├── apply_ppt_to_excel.py # Apply PPT values to Excel
└── generate_diff_report.py # Generate diff reports
```

## Validation Pipeline

1. **Extract PPT Data** → `ppt_extracted_data.json`
2. **Extract Excel Data** → `updated_excel_extracted.json`
3. **Compare Values** → Field-by-field comparison with tolerances
4. **Generate Reports** → `validation_report.json` + `validation_report.md`

## Field Definitions

| Field | Type | Tolerance | Description |
|-------|------|-----------|-------------|
| budget_2026 | currency | ±£1 | 2026 Budget |
| sufficient_2026 | currency | ±£1 | 2026 Sufficient |
| gap_gbp | currency | ±£1 | Gap (GBP) |
| gap_pct | percentage | ±0.1% | Gap (%) |
| awa | percentage | ±0.1% | AWA allocation |
| con | percentage | ±0.1% | CON allocation |
| pur | percentage | ±0.1% | PUR allocation |
| tv | percentage | ±0.1% | TV media % |
| digital | percentage | ±0.1% | Digital media % |
| others | percentage | ±0.1% | Others media % |
| long_campaigns | integer | exact | Long campaign count |
| short_campaigns | integer | exact | Short campaign count |
| long_pct | percentage | ±0.1% | Long campaign % |

## Adding New Fields

1. Add field definition to `FIELD_DEFINITIONS` in `validator.py`:
```python
'new_field': {'type': 'percentage', 'label': 'New Field Name'},
```

2. Add extraction logic to both PPT and Excel extractors

3. Update column mapping in `apply_ppt_to_excel.py`:
```python
FIELD_TO_COL = {
    ...
    'new_field': 25,  # Excel column index
}
```

## Updating Business Rules

### Tolerance Changes
```bash
python scripts/validator.py \
    --tolerance-pct 0.005 \    # 0.5% tolerance
    --tolerance-currency 5.0   # £5 tolerance
```

### Custom Comparison Logic
Edit `compare_values()` in `validator.py` to add new field types or rules.

## Test Coverage

### Adversarial Tests (15 tests)
- Scale errors (% as whole number)
- Sign flips (negative → positive)
- Missing records
- Currency/integer mismatches
- Rounding boundaries
- Null vs zero handling
- Tolerance boundaries

### Sampling Tests (8 tests)
- All fields present (typical record)
- Missing categories (zero values)
- Single record groups
- Large value formatting
- Negative values
- CEJ/media allocation splits

## Reports Generated

| File | Format | Purpose |
|------|--------|---------|
| validation_report.json | JSON | Machine-readable full report |
| validation_report.md | Markdown | Human-readable summary |
| adversarial_test_results.json | JSON | Adversarial test outcomes |
| sampling_test_results.json | JSON | Sampling test outcomes |

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | All validations passed |
| 1 | One or more validations failed |

## Example Output

```
============================================================
VALIDATION RESULTS
============================================================
Records checked: 44
Fields checked: 572
Exact matches: 572
Mismatches: 0
Missing in Excel: 0
Missing in PPT: 1
Pass rate: 100.0%

STATUS: PASS
```
