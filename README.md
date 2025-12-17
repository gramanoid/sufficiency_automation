# Haleon MEA Data Sync Tool

Sync Excel data with PowerPoint presentations while preserving all formatting.

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
3. Click **Sync Data**
4. Download the updated PPT

## What Gets Synced

| Field | Excel Column | Description |
|-------|--------------|-------------|
| Primary Metric | E | Main value |
| Secondary Metric | G | Secondary value |
| Variance | H | Difference |
| Variance % | I | Percentage difference |
| AWA / CON / PUR | J / K / L | Category allocations |
| Channel 1 / 2 / 3 | R / S / V | Distribution percentages |
| Long/Short Items | AD / AF | Item counts |
| Long % | AH | Long item percentage |

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
    --excel "path/to/data.xlsx" \
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

1. **Reads Excel**: Extracts values from data sheet
2. **Finds PPT Tables**: Locates data tables by detecting CATEGORY/BRAND columns
3. **Matches Rows**: Links Excel rows to PPT rows by name
4. **Updates Text Only**: Changes cell text while preserving all formatting (fonts, colors, borders)

## Troubleshooting

**"Item not found in Excel"**
- Check spelling matches between Excel and PPT
- Verify the item is in the correct market row range

**Formatting looks different**
- The tool preserves PPT formatting, but normalizes number display (e.g., removes extra spaces)

**Missing market**
- Ensure the market name appears in the slide text for detection
