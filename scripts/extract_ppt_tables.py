#!/usr/bin/env python3
"""
PPT Table Extractor v2 - Properly parses OOXML table structure
Handles column alignment correctly by using actual table cells
"""

import xml.etree.ElementTree as ET
import os
import json
import re
from pathlib import Path

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
}

def parse_currency(text):
    """Parse currency string to float"""
    if not text or text.strip() in ['-', 'N/A', 'NA', '', '£']:
        return None
    text = text.strip()
    # Remove currency symbols, spaces, and formatting
    text = re.sub(r'[£$€\s,]', '', text)
    # Handle parentheses for negative
    if text.startswith('(') and text.endswith(')'):
        text = '-' + text[1:-1]
    # Handle trailing dash
    if text == '-' or text == '':
        return None
    try:
        return float(text)
    except ValueError:
        return None

def parse_percentage(text):
    """Parse percentage string to float (0-1 scale)"""
    if not text or text.strip() in ['-', 'N/A', 'NA', '']:
        return None
    
    original_text = text.strip()
    has_percent_sign = '%' in original_text
    
    text = original_text.replace('%', '').replace(' ', '')
    if text == '-' or text == '':
        return None
    try:
        val = float(text)
        # If original had % sign, always divide by 100
        if has_percent_sign:
            return val / 100
        # Otherwise use heuristics: if > 1 or < -1, assume percentage form
        if abs(val) > 1:
            return val / 100
        return val
    except ValueError:
        return None

def parse_integer(text):
    """Parse integer from text"""
    if not text or text.strip() in ['-', 'N/A', 'NA', '']:
        return None
    text = text.strip()
    try:
        return int(float(text))
    except ValueError:
        return None

def extract_table_cells(tbl_element):
    """Extract all cells from a table element as a 2D array"""
    rows = []
    for tr in tbl_element.findall('.//a:tr', NS):
        row_cells = []
        for tc in tr.findall('.//a:tc', NS):
            # Extract all text from cell
            texts = []
            for t in tc.findall('.//a:t', NS):
                if t.text:
                    texts.append(t.text)
            cell_text = ''.join(texts).strip()

            # Get gridSpan for merged cells
            grid_span = int(tc.get('gridSpan', '1'))

            # Add the cell text
            row_cells.append(cell_text)
            # Add empty cells for spans (to maintain column alignment)
            for _ in range(grid_span - 1):
                row_cells.append('')

        rows.append(row_cells)
    return rows

def is_data_table(rows):
    """Check if this is a budget/sufficiency data table"""
    if not rows or len(rows) < 2:
        return False

    # Check header row for signature columns
    header = ' '.join(str(c).upper() for c in rows[0])
    signatures = ['BRAND', 'BUDGET', 'CATEGORY', 'SUFFICIENT', 'AWA', 'CON']
    matches = sum(1 for s in signatures if s in header)
    return matches >= 4

def find_column_indices(header_row):
    """Map column names to indices"""
    col_map = {}
    for idx, cell in enumerate(header_row):
        cell_upper = str(cell).upper().strip()

        if 'CATEGORY' in cell_upper:
            col_map['category'] = idx
        elif 'BRAND' in cell_upper and 'LONG' not in cell_upper:
            col_map['brand'] = idx
        elif '2026' in cell_upper and 'BUDGET' in cell_upper and 'SUFFICIENT' not in cell_upper:
            col_map['budget_2026'] = idx
        elif 'SUFFICIENT' in cell_upper or 'MUST HAVE' in cell_upper:
            col_map['sufficient_2026'] = idx
        elif cell_upper == 'GBP 000' or cell_upper == 'GBP000':
            col_map['gap_gbp'] = idx
        elif cell_upper == '%' and 'gap_pct' not in col_map:
            col_map['gap_pct'] = idx
        elif cell_upper == 'AWA':
            col_map['awa'] = idx
        elif cell_upper == 'CON':
            col_map['con'] = idx
        elif cell_upper == 'PUR':
            col_map['pur'] = idx
        elif cell_upper == 'TV':
            col_map['tv'] = idx
        elif cell_upper == 'DIGITAL':
            col_map['digital'] = idx
        elif cell_upper == 'OTHERS':
            col_map['others'] = idx
        elif 'LONG' in cell_upper and 'CAMP' in cell_upper:
            col_map['long_campaigns'] = idx
        elif 'SHORT' in cell_upper and 'CAMP' in cell_upper:
            col_map['short_campaigns'] = idx
        elif cell_upper == 'LONG %' or (cell_upper == '%' and 'long_pct' not in col_map and idx > 10):
            col_map['long_pct'] = idx

    return col_map

def extract_market_from_slide(slide_path):
    """Extract market name from slide content"""
    tree = ET.parse(slide_path)
    root = tree.getroot()

    markets = ['KSA', 'GNE', 'TURKEY', 'SOUTH AFRICA', 'EGYPT', 'MOROCCO',
               'FSA', 'KENYA', 'PAKISTAN', 'NIGERIA', 'ALGERIA']

    # Get all text from slide
    all_text = []
    for t in root.findall('.//a:t', NS):
        if t.text:
            all_text.append(t.text.upper())

    full_text = ' '.join(all_text)

    for market in markets:
        if market in full_text:
            return market

    return 'UNKNOWN'

def parse_data_rows(rows, col_map, market):
    """Parse data rows into structured records"""
    records = []
    current_category = None

    # Skip header row
    for row in rows[1:]:
        # Get cell values safely
        def get_val(key):
            idx = col_map.get(key)
            if idx is not None and idx < len(row):
                return row[idx]
            return None

        category_val = get_val('category')
        brand_val = get_val('brand')

        # Update category tracking
        if category_val and category_val.strip():
            current_category = category_val.strip()

        # Skip if no brand
        if not brand_val or not brand_val.strip():
            continue

        brand = brand_val.strip()

        # Skip TOTAL rows for now (we'll handle separately)
        is_total = 'TOTAL' in brand.upper()

        record = {
            'market': market,
            'category': current_category,
            'brand': brand,
            'is_total': is_total,
        }

        # Parse numeric fields
        record['budget_2026'] = parse_currency(get_val('budget_2026'))
        record['sufficient_2026'] = parse_currency(get_val('sufficient_2026'))
        record['gap_gbp'] = parse_currency(get_val('gap_gbp'))
        record['gap_pct'] = parse_percentage(get_val('gap_pct'))

        # CEJ percentages
        record['awa'] = parse_percentage(get_val('awa'))
        record['con'] = parse_percentage(get_val('con'))
        record['pur'] = parse_percentage(get_val('pur'))

        # Media percentages
        record['tv'] = parse_percentage(get_val('tv'))
        record['digital'] = parse_percentage(get_val('digital'))
        record['others'] = parse_percentage(get_val('others'))

        # Campaign counts
        record['long_campaigns'] = parse_integer(get_val('long_campaigns'))
        record['short_campaigns'] = parse_integer(get_val('short_campaigns'))
        record['long_pct'] = parse_percentage(get_val('long_pct'))

        # Validation: check that media percentages roughly sum to 1
        tv = record['tv'] or 0
        digital = record['digital'] or 0
        others = record['others'] or 0
        media_sum = tv + digital + others

        if media_sum > 0:
            record['_media_sum_check'] = round(media_sum, 2)

        records.append(record)

    return records

def main():
    slides_dir = Path('/tmp/pptx_extract/ppt/slides')

    all_records = []
    table_info = []

    # Process each slide
    for slide_file in sorted(slides_dir.glob('slide*.xml'),
                            key=lambda x: int(re.search(r'slide(\d+)', x.name).group(1))):
        slide_num = int(re.search(r'slide(\d+)', slide_file.name).group(1))

        tree = ET.parse(slide_file)
        root = tree.getroot()

        # Find tables
        for tbl in root.findall('.//a:tbl', NS):
            rows = extract_table_cells(tbl)

            if not is_data_table(rows):
                continue

            # Get market from slide context
            market = extract_market_from_slide(slide_file)

            # Find column mappings from header
            col_map = find_column_indices(rows[0])

            # Parse data rows
            records = parse_data_rows(rows, col_map, market)

            table_info.append({
                'slide_num': slide_num,
                'market': market,
                'columns_found': list(col_map.keys()),
                'column_indices': col_map,
                'row_count': len(rows),
                'records_extracted': len(records),
                'header_row': rows[0] if rows else [],
            })

            all_records.extend(records)

    # Build output
    output = {
        'summary': {
            'total_tables_processed': len(table_info),
            'total_records': len(all_records),
            'non_total_records': len([r for r in all_records if not r.get('is_total')]),
            'records_by_market': {},
        },
        'tables': table_info,
        'records': all_records,
    }

    # Count by market
    for r in all_records:
        if r.get('is_total'):
            continue
        market = r.get('market', 'UNKNOWN')
        if market not in output['summary']['records_by_market']:
            output['summary']['records_by_market'][market] = 0
        output['summary']['records_by_market'][market] += 1

    # Save output
    with open('ppt_extracted_data.json', 'w') as f:
        json.dump(output, f, indent=2, default=str)

    # Print summary
    print("="*60)
    print("PPT EXTRACTION COMPLETE (v2)")
    print("="*60)
    print(f"Tables processed: {len(table_info)}")
    print(f"Total records: {len(all_records)}")
    print(f"Non-total records: {output['summary']['non_total_records']}")
    print(f"\nRecords by market:")
    for market, count in sorted(output['summary']['records_by_market'].items()):
        print(f"  {market}: {count}")

    # Show sample data with validation
    print("\n" + "="*60)
    print("SAMPLE RECORDS WITH VALIDATION")
    print("="*60)

    for market in ['KSA', 'TURKEY', 'SOUTH AFRICA']:
        print(f"\n--- {market} ---")
        market_records = [r for r in all_records if r['market'] == market and not r.get('is_total')]
        for r in market_records[:3]:
            print(f"  {r['brand']}:")
            print(f"    Budget: £{r['budget_2026']:,.0f}" if r['budget_2026'] else "    Budget: None")
            print(f"    Sufficient: £{r['sufficient_2026']:,.0f}" if r['sufficient_2026'] else "    Sufficient: None")
            print(f"    Gap: {r['gap_pct']*100:.1f}%" if r['gap_pct'] else "    Gap: 0%")
            print(f"    CEJ: AWA={r['awa']:.0%}, CON={r['con']:.0%}, PUR={r['pur']:.0%}" if r['awa'] else "    CEJ: N/A")
            print(f"    Media: TV={r['tv']:.0%}, Digital={r['digital']:.0%}, Others={r['others']:.0%}" if r['tv'] is not None else "    Media: N/A")
            print(f"    Campaigns: Long={r['long_campaigns']}, Short={r['short_campaigns']}, Long%={r['long_pct']:.0%}" if r['long_campaigns'] is not None else "    Campaigns: N/A")
            if '_media_sum_check' in r:
                check = r['_media_sum_check']
                status = "✓" if 0.95 <= check <= 1.05 else "⚠️"
                print(f"    Media Sum Check: {check:.2f} {status}")

if __name__ == '__main__':
    main()
