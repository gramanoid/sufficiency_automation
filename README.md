# Haleon MEA Budget Sufficiency Sync Tool

Sync Excel budget data with PowerPoint presentations while preserving all formatting.

## Quick Start

### 1. Setup (First Time Only)
```bash
# Create virtual environment and install dependencies
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Run the App
```bash
./run_app.sh
# Or manually:
source .venv/bin/activate
streamlit run streamlit_app.py
```

Open http://localhost:8501 in your browser.

### 3. Use the App
1. Upload your Excel file (with "2026 Sufficiency" sheet)
2. Upload your PowerPoint file
3. Click **Sync Excel → PPT**
4. Download the updated PPT

## What Gets Synced

| Field | Excel Column | Description |
|-------|--------------|-------------|
| 2026 Budget | E | Budget amount |
| 2026 Sufficient | G | Sufficient amount |
| Gap (GBP) | H | Gap in pounds |
| Gap (%) | I | Gap percentage |
| AWA / CON / PUR | J / K / L | CEJ allocations |
| TV / Digital / Others | R / S / V | Media percentages |
| Long/Short Campaigns | AD / AF | Campaign counts |
| Long % | AH | Long campaign percentage |

## Markets Covered

| Market | Excel Rows |
|--------|------------|
| KSA | 5-11 |
| GNE | 13-18 |
| South Africa | 20-27 |
| Turkey | 29-34 |
| Pakistan | 36-40 |
| Egypt | 42-44 |
| Morocco | 46 |
| FSA | 48-50 |
| Kenya | 52-54 |
| Nigeria | 58-59 |

## Output Files

Each sync creates:
- `*_UPDATED_YYYYMMDD_HHMMSS.pptx` - Updated presentation
- `BACKUP_*_YYYYMMDD_HHMMSS.pptx` - Original backup
- `sync_report_*.json` - Detailed change log

## Command Line Usage

```bash
source .venv/bin/activate
python scripts/update_ppt_from_excel.py \
    --ppt "path/to/presentation.pptx" \
    --excel "path/to/budget.xlsx" \
    --output-dir output
```

## Project Structure

```
.
├── streamlit_app.py        # Streamlit web application
├── run_app.sh              # One-click launcher
├── requirements.txt        # Python dependencies
├── scripts/
│   ├── update_ppt_from_excel.py  # Core sync engine
│   ├── validator.py              # Data validation
│   └── ...                       # Other utilities
├── input/                  # Place source files here
└── output/                 # Generated files appear here
```

## How It Works

1. **Reads Excel**: Extracts values from "2026 Sufficiency" sheet
2. **Finds PPT Tables**: Locates data tables by detecting CATEGORY/BRAND columns
3. **Matches Brands**: Links Excel rows to PPT rows by brand name
4. **Updates Text Only**: Changes cell text while preserving all formatting (fonts, colors, borders)

## Troubleshooting

**"Brand not found in Excel"**
- Check brand spelling matches between Excel and PPT
- Verify the brand is in the correct market row range

**Formatting looks different**
- The tool preserves PPT formatting, but normalizes number display (e.g., removes extra spaces)

**Missing market**
- Ensure the market name appears in the slide text for detection
