#!/usr/bin/env python3
"""
Extract data from UPDATED Excel file for validation
"""

import pandas as pd
import json
import re
from pathlib import Path


def parse_currency(val):
    """Parse currency to float"""
    if pd.isna(val) or val == '-' or val == '' or val == 0:
        return None
    if isinstance(val, (int, float)):
        return float(val) if val != 0 else None
    text = str(val).strip()
    text = re.sub(r'[£$€,\s]', '', text)
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    if text == '-' or text == '':
        return None
    try:
        return float(text)
    except ValueError:
        return val


def parse_percentage(val):
    """Parse percentage to float (0-1 scale)"""
    if pd.isna(val) or val == '-' or val == '':
        return None
    if isinstance(val, (int, float)):
        if val == 0:
            return 0.0
        return float(val) if abs(val) <= 1 else float(val) / 100
    text = str(val).strip().replace('%', '')
    if text == '-' or text == '':
        return None
    try:
        v = float(text)
        return v if abs(v) <= 1 else v / 100
    except ValueError:
        return val


def main():
    base_dir = Path(__file__).parent.parent
    xlsx_path = base_dir / 'output' / 'Haleon - 2026 MEA Budget Sufficiency_UPDATED.xlsx'

    print(f"Extracting from: {xlsx_path}")

    df_raw = pd.read_excel(xlsx_path, sheet_name='2026 Sufficiency', header=None)

    valid_markets = {'KSA', 'GNE', 'TURKEY', 'SOUTH AFRICA', 'EGYPT', 'MOROCCO',
                     'FSA', 'KENYA', 'PAKISTAN', 'NIGERIA', 'ALGERIA', 'OSA'}

    records = []
    current_market = None
    current_category = None

    for idx in range(4, len(df_raw)):
        row = df_raw.iloc[idx]

        market_val = row.iloc[1]
        category_val = row.iloc[2]
        brand_val = row.iloc[3]

        if pd.isna(brand_val) or str(brand_val).strip() == '':
            continue

        brand = str(brand_val).strip()

        if brand.upper() in ['BRAND', 'MARKET', 'CATEGORY', 'TOTAL']:
            continue

        if pd.notna(market_val) and str(market_val).strip():
            market_str = str(market_val).strip().upper()
            if market_str in valid_markets or market_str not in ['MARKET', 'NAN']:
                current_market = market_str

        if pd.notna(category_val) and str(category_val).strip():
            cat_str = str(category_val).strip()
            if cat_str.upper() not in ['CATEGORY', 'NAN']:
                current_category = cat_str

        if not current_market or current_market == 'MARKET':
            continue

        record = {
            'market': current_market,
            'category': current_category,
            'brand': brand,
            'is_total': 'TOTAL' in brand.upper(),
            'source': 'updated_excel',
            'excel_row': idx + 1,
        }

        # Budget columns (E=4, F=5, G=6 in 0-indexed -> 5,6,7 in 1-indexed)
        record['budget_2026'] = parse_currency(row.iloc[4])
        record['nice_to_have'] = parse_currency(row.iloc[5])
        record['sufficient_2026'] = parse_currency(row.iloc[6])

        # Gap columns (H=7, I=8)
        record['gap_gbp'] = parse_currency(row.iloc[7])
        record['gap_pct'] = parse_percentage(row.iloc[8])

        # CEJ percentages (J=9, K=10, L=11)
        record['awa'] = parse_percentage(row.iloc[9])
        record['con'] = parse_percentage(row.iloc[10])
        record['pur'] = parse_percentage(row.iloc[11])

        # Media percentages (R=17, S=18, V=21 in 0-indexed)
        record['tv'] = parse_percentage(row.iloc[17])
        record['digital'] = parse_percentage(row.iloc[18])
        record['others'] = parse_percentage(row.iloc[21])

        # Campaign counts (AD=29, AF=31, AH=33 in 0-indexed)
        try:
            val = row.iloc[29]
            record['long_campaigns'] = int(float(val)) if pd.notna(val) and val != '-' else None
        except (ValueError, TypeError):
            record['long_campaigns'] = None

        try:
            val = row.iloc[31]
            record['short_campaigns'] = int(float(val)) if pd.notna(val) and val != '-' else None
        except (ValueError, TypeError):
            record['short_campaigns'] = None

        record['long_pct'] = parse_percentage(row.iloc[33])

        if record['budget_2026'] is not None or record['brand']:
            records.append(record)

    output = {
        'summary': {
            'total_records': len(records),
            'source_file': str(xlsx_path.name),
            'records_by_market': {}
        },
        'records': records
    }

    for r in records:
        market = r.get('market', 'UNKNOWN')
        if market not in output['summary']['records_by_market']:
            output['summary']['records_by_market'][market] = 0
        output['summary']['records_by_market'][market] += 1

    output_path = base_dir / 'output' / 'data' / 'updated_excel_extracted.json'
    with open(output_path, 'w') as f:
        json.dump(output, f, indent=2, default=str)

    print(f"Extracted {len(records)} records")
    print(f"Saved to: {output_path}")


if __name__ == '__main__':
    main()
