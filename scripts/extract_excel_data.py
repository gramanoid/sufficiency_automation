#!/usr/bin/env python3
"""
Excel Data Extractor - Extracts data from 2026 Sufficiency sheet
Handles complex headers and preserves structure for comparison
"""

import pandas as pd
import json
import re

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
        # If > 1, assume it's percentage form
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
    xlsx_path = 'Haleon - 2026 MEA Budget Sufficiency_271125_Final 1.xlsx'

    # Read raw to understand structure
    df_raw = pd.read_excel(xlsx_path, sheet_name='2026 Sufficiency', header=None)

    # Column mapping (verified from header row 3):
    # Col 1: MARKET, Col 2: CATEGORY, Col 3: BRAND
    # Col 4: 2026 BUDGET, Col 5: NICE TO HAVE, Col 6: 2026 SUFFICIENT
    # Col 7: GBP 000 (gap), Col 8: % (gap)
    # Col 9: AWA, Col 10: CON, Col 11: PUR (CEJ percentages)
    # Col 17: TV, Col 18: Digital, Col 19: eCom, Col 20: Innovation, Col 21: Others (Media %)
    # Col 29: LONG Camps, Col 31: SHORT Camps, Col 33: Long %

    # Valid markets
    valid_markets = {'KSA', 'GNE', 'TURKEY', 'SOUTH AFRICA', 'EGYPT', 'MOROCCO',
                     'FSA', 'KENYA', 'PAKISTAN', 'NIGERIA', 'ALGERIA', 'OSA'}

    records = []
    current_market = None
    current_category = None

    # Data starts at row 4 (index 4 in 0-based)
    for idx in range(4, len(df_raw)):
        row = df_raw.iloc[idx]

        # Get market, category, brand
        market_val = row.iloc[1]
        category_val = row.iloc[2]
        brand_val = row.iloc[3]

        # Skip if no brand or brand is a header
        if pd.isna(brand_val) or str(brand_val).strip() == '':
            continue

        brand = str(brand_val).strip()

        # Skip header rows that got included
        if brand.upper() in ['BRAND', 'MARKET', 'CATEGORY', 'TOTAL']:
            continue

        # Update market tracking
        if pd.notna(market_val) and str(market_val).strip():
            market_str = str(market_val).strip().upper()
            if market_str in valid_markets or market_str not in ['MARKET', 'NAN']:
                current_market = market_str

        # Update category tracking
        if pd.notna(category_val) and str(category_val).strip():
            cat_str = str(category_val).strip()
            if cat_str.upper() not in ['CATEGORY', 'NAN']:
                current_category = cat_str

        # Skip if we don't have a valid market
        if not current_market or current_market == 'MARKET':
            continue

        record = {
            'market': current_market,
            'category': current_category,
            'brand': brand,
            'is_total': 'TOTAL' in brand.upper(),
            'source': 'excel',
            'excel_row': idx + 1,  # 1-based row number for Excel reference
        }

        # Budget columns (4, 5, 6)
        record['budget_2026'] = parse_currency(row.iloc[4])
        record['nice_to_have'] = parse_currency(row.iloc[5])
        record['sufficient_2026'] = parse_currency(row.iloc[6])

        # Gap columns (7, 8)
        record['gap_gbp'] = parse_currency(row.iloc[7])
        record['gap_pct'] = parse_percentage(row.iloc[8])

        # CEJ percentages (9, 10, 11)
        record['awa'] = parse_percentage(row.iloc[9])
        record['con'] = parse_percentage(row.iloc[10])
        record['pur'] = parse_percentage(row.iloc[11])

        # Media percentages (17, 18, 19, 20, 21)
        record['tv'] = parse_percentage(row.iloc[17])
        record['digital'] = parse_percentage(row.iloc[18])
        record['ecom'] = parse_percentage(row.iloc[19])
        record['innovation'] = parse_percentage(row.iloc[20])
        record['others'] = parse_percentage(row.iloc[21])

        # Campaign counts (29, 31) and Long % (33)
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

        # Only add valid records (must have budget or brand)
        if record['budget_2026'] is not None or record['brand']:
            records.append(record)

    # Summary
    markets = set(r['market'] for r in records if r['market'])

    output = {
        'summary': {
            'total_records': len(records),
            'markets': sorted(list(markets)),
            'records_by_market': {}
        },
        'records': records
    }

    for r in records:
        market = r.get('market', 'UNKNOWN')
        if market not in output['summary']['records_by_market']:
            output['summary']['records_by_market'][market] = 0
        output['summary']['records_by_market'][market] += 1

    with open('excel_extracted_data.json', 'w') as f:
        json.dump(output, f, indent=2, default=str)

    print(f"Extraction complete!")
    print(f"Total records: {len(records)}")
    print(f"Markets: {sorted(markets)}")
    print(f"\nRecords by market:")
    for market, count in sorted(output['summary']['records_by_market'].items()):
        print(f"  {market}: {count}")

    # Show sample records
    print("\n=== Sample Records ===")
    for market in ['KSA', 'TURKEY']:
        print(f"\n--- {market} ---")
        for r in records:
            if r['market'] == market and not r.get('is_total'):
                print(f"  {r['brand']}: Budget={r['budget_2026']}, Sufficient={r['sufficient_2026']}")
                if r['brand'] == 'Sensodyne':
                    print(f"    AWA={r['awa']}, CON={r['con']}, PUR={r['pur']}")
                    print(f"    TV={r['tv']}, Digital={r['digital']}, Others={r['others']}")
                    print(f"    Long={r['long_campaigns']}, Short={r['short_campaigns']}, Long%={r['long_pct']}")

if __name__ == '__main__':
    main()
