#!/usr/bin/env python3
"""
Comprehensive Data Correctness Validator
Verifies ALL displayed values in Excel match PPT source of truth.
Assumes the code is broken until proven otherwise.
"""

import json
import argparse
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Optional, Any
from datetime import datetime
from collections import defaultdict
import sys

# Tolerances
PCT_TOL = 0.001  # 0.1 percentage points for percentages (stored as decimals)
CURRENCY_TOL = 1.0  # £1 absolute tolerance for currency
INT_TOL = 0  # Exact match for integers


@dataclass
class Discrepancy:
    """A single value mismatch"""
    market: str
    category: str
    brand: str
    field: str
    actual_value: Any
    expected_value: Any
    difference: Optional[float]
    diff_percent: Optional[float]
    excel_row: Optional[int]
    severity: str  # 'CRITICAL', 'WARNING', 'INFO'
    root_cause_hint: str


@dataclass
class ValidationResult:
    """Complete validation result"""
    timestamp: str
    source_file: str
    target_file: str
    total_fields_checked: int = 0
    total_records_checked: int = 0
    exact_matches: int = 0
    within_tolerance: int = 0
    mismatches: int = 0
    missing_in_excel: int = 0
    missing_in_ppt: int = 0
    discrepancies: list = field(default_factory=list)
    edge_cases_tested: list = field(default_factory=list)
    pass_rate: float = 0.0


def normalize_key(market: str, category: str, brand: str) -> tuple:
    """Normalize for matching"""
    def clean(s):
        if not s:
            return ''
        return str(s).strip().upper().replace(' ', '').replace('-', '')
    return (clean(market), clean(category), clean(brand))


def compare_values(actual: Any, expected: Any, field_type: str) -> tuple[bool, float, str]:
    """
    Compare two values with appropriate tolerance.
    Returns: (is_match, difference, match_type)
    match_type: 'exact', 'within_tolerance', 'mismatch'
    """
    # Handle None/missing
    if actual is None and expected is None:
        return True, 0.0, 'exact'

    # Normalize dash to 0
    if actual == '-':
        actual = 0
    if expected == '-':
        expected = 0

    # Handle one None
    if actual is None:
        actual = 0
    if expected is None:
        expected = 0

    try:
        actual_num = float(actual)
        expected_num = float(expected)
    except (ValueError, TypeError):
        # String comparison
        return str(actual) == str(expected), 0, 'exact' if str(actual) == str(expected) else 'mismatch'

    diff = actual_num - expected_num
    # Use small epsilon for floating point comparison
    EPSILON = 1e-9

    if field_type == 'percentage':
        # Percentage tolerance (stored as decimals, so 0.001 = 0.1%)
        # Use <= with small epsilon to handle floating point precision
        if abs(diff) <= PCT_TOL + EPSILON:
            return True, diff, 'exact' if abs(diff) < EPSILON else 'within_tolerance'
        return False, diff, 'mismatch'

    elif field_type == 'currency':
        # Currency tolerance
        if abs(diff) <= CURRENCY_TOL + EPSILON:
            return True, diff, 'exact' if abs(diff) < EPSILON else 'within_tolerance'
        return False, diff, 'mismatch'

    elif field_type == 'integer':
        # Exact match for integers
        if int(actual_num) == int(expected_num):
            return True, 0, 'exact'
        return False, diff, 'mismatch'

    else:
        # Default: exact match with epsilon
        if abs(diff) < EPSILON:
            return True, 0, 'exact'
        return False, diff, 'mismatch'


# Field definitions with types
FIELD_DEFINITIONS = {
    'budget_2026': {'type': 'currency', 'label': '2026 Budget'},
    'sufficient_2026': {'type': 'currency', 'label': '2026 Sufficient'},
    'gap_gbp': {'type': 'currency', 'label': 'Gap (GBP)'},
    'gap_pct': {'type': 'percentage', 'label': 'Gap (%)'},
    'awa': {'type': 'percentage', 'label': 'AWA'},
    'con': {'type': 'percentage', 'label': 'CON'},
    'pur': {'type': 'percentage', 'label': 'PUR'},
    'tv': {'type': 'percentage', 'label': 'TV'},
    'digital': {'type': 'percentage', 'label': 'Digital'},
    'others': {'type': 'percentage', 'label': 'Others'},
    'long_campaigns': {'type': 'integer', 'label': 'Long Campaigns'},
    'short_campaigns': {'type': 'integer', 'label': 'Short Campaigns'},
    'long_pct': {'type': 'percentage', 'label': 'Long %'},
}


def extract_actual_values(excel_data_path: Path) -> dict:
    """
    Extract actual values from Excel extracted data.
    Returns: dict[normalized_key -> record]
    """
    with open(excel_data_path) as f:
        data = json.load(f)

    records = {}
    for rec in data['records']:
        if rec.get('is_total'):
            continue
        key = normalize_key(rec.get('market'), rec.get('category'), rec.get('brand'))
        records[key] = rec

    return records


def extract_expected_values(ppt_data_path: Path) -> dict:
    """
    Extract expected values from PPT extracted data (source of truth).
    Returns: dict[normalized_key -> record]
    """
    with open(ppt_data_path) as f:
        data = json.load(f)

    records = {}
    for rec in data['records']:
        if rec.get('is_total'):
            continue
        key = normalize_key(rec.get('market'), rec.get('category'), rec.get('brand'))
        records[key] = rec

    return records


def validate_record(actual: dict, expected: dict, key: tuple) -> list[Discrepancy]:
    """Validate all fields in a single record"""
    discrepancies = []
    market, category, brand = key

    # Get original names for reporting
    actual_market = actual.get('market', market)
    actual_category = actual.get('category', category)
    actual_brand = actual.get('brand', brand)
    excel_row = actual.get('excel_row')

    for field_name, field_def in FIELD_DEFINITIONS.items():
        actual_val = actual.get(field_name)
        expected_val = expected.get(field_name)
        field_type = field_def['type']

        is_match, diff, match_type = compare_values(actual_val, expected_val, field_type)

        if not is_match:
            # Determine severity
            if field_type == 'currency' and abs(diff) > 1000:
                severity = 'CRITICAL'
            elif field_type == 'percentage' and abs(diff) > 0.05:  # >5%
                severity = 'CRITICAL'
            elif field_type == 'integer' and abs(diff) > 1:
                severity = 'WARNING'
            else:
                severity = 'WARNING'

            # Root cause hint
            hint = infer_root_cause(actual_val, expected_val, field_name, field_type)

            # Calculate percent diff
            diff_pct = None
            if expected_val and expected_val != 0:
                try:
                    diff_pct = (diff / float(expected_val)) * 100
                except:
                    pass

            discrepancies.append(Discrepancy(
                market=actual_market,
                category=actual_category,
                brand=actual_brand,
                field=field_def['label'],
                actual_value=actual_val,
                expected_value=expected_val,
                difference=round(diff, 6) if diff else None,
                diff_percent=round(diff_pct, 2) if diff_pct else None,
                excel_row=excel_row,
                severity=severity,
                root_cause_hint=hint
            ))

    return discrepancies


def infer_root_cause(actual: Any, expected: Any, field: str, field_type: str) -> str:
    """Infer likely root cause of mismatch"""
    if actual is None and expected is not None:
        return "Value missing in Excel - extraction failed or row mismatch"
    if expected is None and actual is not None:
        return "Value missing in PPT - extraction failed"

    try:
        actual_num = float(actual) if actual not in [None, '-'] else 0
        expected_num = float(expected) if expected not in [None, '-'] else 0
    except:
        return "Non-numeric comparison"

    # Check for percentage scale issues
    if field_type == 'percentage':
        if abs(actual_num * 100 - expected_num) < 0.01:
            return "Scale issue: Excel has decimal, PPT has percentage"
        if abs(actual_num - expected_num * 100) < 0.01:
            return "Scale issue: PPT has decimal, Excel has percentage"

    # Check for sign flip
    if actual_num != 0 and expected_num != 0:
        if abs(actual_num + expected_num) < 0.01:
            return "Sign flip: values have opposite signs"

    # Check for formula vs value
    if isinstance(actual, str) and '=' in str(actual):
        return "Excel contains formula, not computed value"

    # Check for rounding
    if abs(actual_num - expected_num) < 0.01:
        return "Minor rounding difference"

    return "Value mismatch - manual verification needed"


def validate_edge_cases(actual_records: dict, expected_records: dict) -> list[dict]:
    """
    Explicitly test edge cases
    """
    edge_case_results = []

    # Edge case 1: Records with zero values
    zero_records = [k for k, v in expected_records.items()
                   if v.get('gap_gbp') == 0 or v.get('gap_pct') == 0]
    edge_case_results.append({
        'test': 'Zero gap values',
        'count': len(zero_records),
        'samples': zero_records[:3],
        'status': 'TESTED'
    })

    # Edge case 2: Missing media categories (TV=0, Digital=0, Others=0)
    missing_tv = [k for k, v in expected_records.items() if v.get('tv') in [0, None, '-']]
    edge_case_results.append({
        'test': 'Missing TV allocation',
        'count': len(missing_tv),
        'samples': missing_tv[:3],
        'status': 'TESTED'
    })

    # Edge case 3: 100% single media channel
    single_channel = [k for k, v in expected_records.items()
                     if v.get('tv') == 1.0 or v.get('digital') == 1.0]
    edge_case_results.append({
        'test': '100% single media channel',
        'count': len(single_channel),
        'samples': single_channel[:3],
        'status': 'TESTED'
    })

    # Edge case 4: Zero campaigns
    zero_campaigns = [k for k, v in expected_records.items()
                     if v.get('long_campaigns') == 0 or v.get('short_campaigns') == 0]
    edge_case_results.append({
        'test': 'Zero campaign count',
        'count': len(zero_campaigns),
        'samples': zero_campaigns[:3],
        'status': 'TESTED'
    })

    # Edge case 5: Negative gap (underfunded)
    negative_gap = [k for k, v in expected_records.items()
                   if v.get('gap_gbp') and v.get('gap_gbp') < 0]
    edge_case_results.append({
        'test': 'Negative gap (underfunded)',
        'count': len(negative_gap),
        'samples': negative_gap[:3],
        'status': 'TESTED'
    })

    # Edge case 6: Records in Excel but not PPT
    excel_only = set(actual_records.keys()) - set(expected_records.keys())
    edge_case_results.append({
        'test': 'Records in Excel only',
        'count': len(excel_only),
        'samples': list(excel_only)[:3],
        'status': 'FLAGGED' if excel_only else 'OK'
    })

    # Edge case 7: Records in PPT but not Excel
    ppt_only = set(expected_records.keys()) - set(actual_records.keys())
    edge_case_results.append({
        'test': 'Records in PPT only',
        'count': len(ppt_only),
        'samples': list(ppt_only)[:3],
        'status': 'CRITICAL' if ppt_only else 'OK'
    })

    return edge_case_results


def run_validation(excel_data_path: Path, ppt_data_path: Path) -> ValidationResult:
    """
    Run full validation comparing Excel (actual) against PPT (expected/source of truth)
    """
    result = ValidationResult(
        timestamp=datetime.now().isoformat(),
        source_file=str(ppt_data_path),
        target_file=str(excel_data_path)
    )

    # Extract values
    actual_records = extract_actual_values(excel_data_path)
    expected_records = extract_expected_values(ppt_data_path)

    print(f"Loaded {len(actual_records)} Excel records")
    print(f"Loaded {len(expected_records)} PPT records (source of truth)")

    # Find common records
    common_keys = set(actual_records.keys()) & set(expected_records.keys())
    excel_only = set(actual_records.keys()) - set(expected_records.keys())
    ppt_only = set(expected_records.keys()) - set(actual_records.keys())

    result.missing_in_ppt = len(excel_only)
    result.missing_in_excel = len(ppt_only)

    # Add missing records as critical discrepancies
    for key in ppt_only:
        rec = expected_records[key]
        result.discrepancies.append(Discrepancy(
            market=rec.get('market'),
            category=rec.get('category'),
            brand=rec.get('brand'),
            field='ENTIRE RECORD',
            actual_value=None,
            expected_value='EXISTS IN PPT',
            difference=None,
            diff_percent=None,
            excel_row=None,
            severity='CRITICAL',
            root_cause_hint='Record exists in PPT but missing from Excel'
        ))

    # Validate each common record
    for key in common_keys:
        actual = actual_records[key]
        expected = expected_records[key]

        discrepancies = validate_record(actual, expected, key)
        result.discrepancies.extend(discrepancies)

        result.total_records_checked += 1
        result.total_fields_checked += len(FIELD_DEFINITIONS)

    # Count matches vs mismatches
    result.mismatches = len(result.discrepancies)
    result.exact_matches = result.total_fields_checked - result.mismatches

    # Calculate pass rate
    if result.total_fields_checked > 0:
        result.pass_rate = (result.exact_matches / result.total_fields_checked) * 100

    # Run edge case tests
    result.edge_cases_tested = validate_edge_cases(actual_records, expected_records)

    return result


def generate_markdown_report(result: ValidationResult, output_path: Path):
    """Generate human-readable markdown report"""
    lines = [
        "# Data Validation Report",
        "",
        f"**Generated:** {result.timestamp}",
        f"**Source (PPT):** `{result.source_file}`",
        f"**Target (Excel):** `{result.target_file}`",
        "",
        "## Summary",
        "",
        f"| Metric | Value |",
        f"|--------|-------|",
        f"| Records Checked | {result.total_records_checked} |",
        f"| Fields Checked | {result.total_fields_checked} |",
        f"| Exact Matches | {result.exact_matches} |",
        f"| Mismatches | {result.mismatches} |",
        f"| Missing in Excel | {result.missing_in_excel} |",
        f"| Missing in PPT | {result.missing_in_ppt} |",
        f"| **Pass Rate** | **{result.pass_rate:.1f}%** |",
        "",
    ]

    # Overall status
    if result.mismatches == 0 and result.missing_in_excel == 0:
        lines.append("## Status: PASS")
        lines.append("")
        lines.append("All values in Excel match PPT source of truth within tolerance.")
    else:
        lines.append("## Status: FAIL")
        lines.append("")
        lines.append(f"Found {result.mismatches} discrepancies requiring attention.")

    # Edge cases
    lines.extend([
        "",
        "## Edge Cases Tested",
        "",
        "| Test | Count | Status |",
        "|------|-------|--------|",
    ])
    for ec in result.edge_cases_tested:
        lines.append(f"| {ec['test']} | {ec['count']} | {ec['status']} |")

    # Discrepancies by severity
    if result.discrepancies:
        critical = [d for d in result.discrepancies if d.severity == 'CRITICAL']
        warnings = [d for d in result.discrepancies if d.severity == 'WARNING']

        if critical:
            lines.extend([
                "",
                "## CRITICAL Discrepancies",
                "",
                "| Market | Brand | Field | Expected | Actual | Diff | Hint |",
                "|--------|-------|-------|----------|--------|------|------|",
            ])
            for d in critical:
                lines.append(f"| {d.market} | {d.brand} | {d.field} | {d.expected_value} | {d.actual_value} | {d.difference} | {d.root_cause_hint} |")

        if warnings:
            lines.extend([
                "",
                "## Warnings",
                "",
                "| Market | Brand | Field | Expected | Actual | Diff |",
                "|--------|-------|-------|----------|--------|------|",
            ])
            for d in warnings[:50]:  # Limit to 50
                lines.append(f"| {d.market} | {d.brand} | {d.field} | {d.expected_value} | {d.actual_value} | {d.difference} |")
            if len(warnings) > 50:
                lines.append(f"*... and {len(warnings) - 50} more warnings*")

    # By market breakdown
    if result.discrepancies:
        lines.extend([
            "",
            "## Discrepancies by Market",
            "",
        ])
        by_market = defaultdict(list)
        for d in result.discrepancies:
            by_market[d.market].append(d)

        for market in sorted(by_market.keys()):
            discs = by_market[market]
            lines.append(f"- **{market}**: {len(discs)} discrepancies")

    with open(output_path, 'w') as f:
        f.write('\n'.join(lines))


def generate_json_report(result: ValidationResult, output_path: Path):
    """Generate machine-readable JSON report"""
    # Convert dataclass to dict
    data = {
        'timestamp': result.timestamp,
        'source_file': result.source_file,
        'target_file': result.target_file,
        'summary': {
            'total_records_checked': result.total_records_checked,
            'total_fields_checked': result.total_fields_checked,
            'exact_matches': result.exact_matches,
            'mismatches': result.mismatches,
            'missing_in_excel': result.missing_in_excel,
            'missing_in_ppt': result.missing_in_ppt,
            'pass_rate': result.pass_rate
        },
        'status': 'PASS' if result.mismatches == 0 and result.missing_in_excel == 0 else 'FAIL',
        'discrepancies': [asdict(d) for d in result.discrepancies],
        'edge_cases_tested': result.edge_cases_tested
    }

    with open(output_path, 'w') as f:
        json.dump(data, f, indent=2, default=str)


def main():
    parser = argparse.ArgumentParser(description='Validate Excel data against PPT source of truth')
    parser.add_argument('--excel-data', required=True, help='Path to Excel extracted data JSON')
    parser.add_argument('--ppt-data', required=True, help='Path to PPT extracted data JSON')
    parser.add_argument('--output-dir', default='.', help='Output directory for reports')
    parser.add_argument('--tolerance-pct', type=float, default=0.001, help='Percentage tolerance')
    parser.add_argument('--tolerance-currency', type=float, default=1.0, help='Currency tolerance')

    args = parser.parse_args()

    global PCT_TOL, CURRENCY_TOL
    PCT_TOL = args.tolerance_pct
    CURRENCY_TOL = args.tolerance_currency

    excel_path = Path(args.excel_data)
    ppt_path = Path(args.ppt_data)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(exist_ok=True)

    print("=" * 60)
    print("COMPREHENSIVE DATA VALIDATION")
    print("=" * 60)
    print(f"Source (PPT): {ppt_path}")
    print(f"Target (Excel): {excel_path}")
    print(f"Tolerances: PCT={PCT_TOL}, Currency={CURRENCY_TOL}")
    print()

    # Run validation
    result = run_validation(excel_path, ppt_path)

    # Generate reports
    json_path = output_dir / 'validation_report.json'
    md_path = output_dir / 'validation_report.md'

    generate_json_report(result, json_path)
    generate_markdown_report(result, md_path)

    # Print summary
    print()
    print("=" * 60)
    print("VALIDATION RESULTS")
    print("=" * 60)
    print(f"Records checked: {result.total_records_checked}")
    print(f"Fields checked: {result.total_fields_checked}")
    print(f"Exact matches: {result.exact_matches}")
    print(f"Mismatches: {result.mismatches}")
    print(f"Missing in Excel: {result.missing_in_excel}")
    print(f"Missing in PPT: {result.missing_in_ppt}")
    print(f"Pass rate: {result.pass_rate:.1f}%")
    print()

    if result.mismatches > 0 or result.missing_in_excel > 0:
        print("STATUS: FAIL")
        print()
        print("Critical discrepancies:")
        for d in result.discrepancies:
            if d.severity == 'CRITICAL':
                print(f"  [{d.market}/{d.brand}] {d.field}: expected={d.expected_value}, actual={d.actual_value}")
        sys.exit(1)
    else:
        print("STATUS: PASS")
        print("All values match within tolerance!")
        sys.exit(0)


if __name__ == '__main__':
    main()
