#!/usr/bin/env python3
"""
Sampling Strategy Tests
Validates at least 5 distinct combinations across markets/brands
Includes edge case scenarios as required by the spec.
"""

import json
from pathlib import Path
from dataclasses import dataclass
from typing import Optional
import sys

sys.path.insert(0, str(Path(__file__).parent))
from validator import compare_values, FIELD_DEFINITIONS, normalize_key


@dataclass
class SampleTest:
    """A single sample validation test"""
    name: str
    market: str
    category: str
    brand: str
    scenario: str  # 'all_present', 'category_missing', 'all_zeros', 'single_record', 'large_values'
    fields_to_check: list
    expected_result: str  # 'PASS' or 'FAIL'


def load_data(base_dir: Path):
    """Load extracted data"""
    with open(base_dir / 'output' / 'data' / 'updated_excel_extracted.json') as f:
        excel_data = json.load(f)
    with open(base_dir / 'output' / 'data' / 'ppt_extracted_data.json') as f:
        ppt_data = json.load(f)
    return excel_data, ppt_data


def build_lookup(data: dict) -> dict:
    """Build lookup by normalized key"""
    lookup = {}
    for rec in data['records']:
        if rec.get('is_total'):
            continue
        key = normalize_key(rec.get('market'), rec.get('category'), rec.get('brand'))
        lookup[key] = rec
    return lookup


def validate_sample(excel_rec: dict, ppt_rec: dict, fields: list) -> tuple[bool, list]:
    """Validate specific fields between Excel and PPT records"""
    errors = []
    for field in fields:
        if field not in FIELD_DEFINITIONS:
            continue

        field_type = FIELD_DEFINITIONS[field]['type']
        excel_val = excel_rec.get(field)
        ppt_val = ppt_rec.get(field)

        is_match, diff, match_type = compare_values(excel_val, ppt_val, field_type)

        if not is_match:
            errors.append({
                'field': field,
                'excel': excel_val,
                'ppt': ppt_val,
                'diff': diff
            })

    return len(errors) == 0, errors


def run_sampling_tests():
    """Run comprehensive sampling validation"""
    base_dir = Path(__file__).parent.parent
    excel_data, ppt_data = load_data(base_dir)

    excel_lookup = build_lookup(excel_data)
    ppt_lookup = build_lookup(ppt_data)

    # Define 5+ sample tests covering different scenarios
    sample_tests = [
        # Scenario 1: All categories present (typical record with all fields)
        SampleTest(
            name="KSA Sensodyne - All Fields Present",
            market="KSA",
            category="OHC",
            brand="Sensodyne",
            scenario="all_present",
            fields_to_check=list(FIELD_DEFINITIONS.keys()),
            expected_result="PASS"
        ),

        # Scenario 2: Record with zero/missing campaign data
        SampleTest(
            name="South Africa Med-Lemon - Zero Long Campaigns",
            market="SOUTH AFRICA",
            category="Self-Care",
            brand="Med-Lemon",
            scenario="category_missing",
            fields_to_check=['tv', 'digital', 'others', 'long_campaigns', 'short_campaigns'],
            expected_result="PASS"
        ),

        # Scenario 3: Record with negative gap (underfunded)
        SampleTest(
            name="Turkey Sensodyne - Negative Gap",
            market="TURKEY",
            category="OHC",
            brand="Sensodyne",
            scenario="negative_values",
            fields_to_check=['gap_gbp', 'gap_pct', 'budget_2026', 'sufficient_2026'],
            expected_result="PASS"
        ),

        # Scenario 4: Single record market (Morocco has only 1 brand)
        SampleTest(
            name="Morocco Sensodyne - Single Market Record",
            market="MOROCCO",
            category="OHC",
            brand="Sensodyne",
            scenario="single_record",
            fields_to_check=list(FIELD_DEFINITIONS.keys()),
            expected_result="PASS"
        ),

        # Scenario 5: Large budget values (formatting stress)
        SampleTest(
            name="GNE Sensodyne - Large Budget Values",
            market="GNE",
            category="OHC",
            brand="Sensodyne",
            scenario="large_values",
            fields_to_check=['budget_2026', 'sufficient_2026', 'gap_gbp'],
            expected_result="PASS"
        ),

        # Scenario 6: 100% digital allocation
        SampleTest(
            name="Turkey Corega - High Digital Allocation",
            market="TURKEY",
            category="OHC",
            brand="Corega",
            scenario="single_channel",
            fields_to_check=['tv', 'digital', 'others'],
            expected_result="PASS"
        ),

        # Scenario 7: CEJ allocation validation
        SampleTest(
            name="Pakistan Sensodyne - CEJ Percentages",
            market="PAKISTAN",
            category="OHC",
            brand="Sensodyne",
            scenario="all_present",
            fields_to_check=['awa', 'con', 'pur'],
            expected_result="PASS"
        ),

        # Scenario 8: Campaign split validation
        SampleTest(
            name="South Africa Eno - Campaign Distribution",
            market="SOUTH AFRICA",
            category="Wellness",
            brand="Eno",
            scenario="campaign_split",
            fields_to_check=['long_campaigns', 'short_campaigns', 'long_pct'],
            expected_result="PASS"
        ),
    ]

    print("=" * 70)
    print("SAMPLING STRATEGY VALIDATION")
    print("=" * 70)
    print(f"\nRunning {len(sample_tests)} sample tests...\n")

    results = []
    passed = 0
    failed = 0

    for test in sample_tests:
        key = normalize_key(test.market, test.category, test.brand)

        excel_rec = excel_lookup.get(key)
        ppt_rec = ppt_lookup.get(key)

        if not excel_rec:
            print(f"  [SKIP] {test.name} - Not found in Excel")
            results.append({'test': test.name, 'status': 'SKIP', 'reason': 'Not in Excel'})
            continue

        if not ppt_rec:
            print(f"  [SKIP] {test.name} - Not found in PPT")
            results.append({'test': test.name, 'status': 'SKIP', 'reason': 'Not in PPT'})
            continue

        is_valid, errors = validate_sample(excel_rec, ppt_rec, test.fields_to_check)

        if is_valid:
            print(f"  [PASS] {test.name}")
            print(f"         Scenario: {test.scenario}, Fields: {len(test.fields_to_check)}")
            passed += 1
            results.append({
                'test': test.name,
                'status': 'PASS',
                'scenario': test.scenario,
                'fields_checked': test.fields_to_check,
                'market': test.market,
                'brand': test.brand
            })
        else:
            print(f"  [FAIL] {test.name}")
            print(f"         Errors: {errors}")
            failed += 1
            results.append({
                'test': test.name,
                'status': 'FAIL',
                'scenario': test.scenario,
                'errors': errors
            })

    print()
    print("=" * 70)
    print(f"SAMPLING RESULTS: {passed} passed, {failed} failed, {len(sample_tests) - passed - failed} skipped")
    print("=" * 70)

    # Additional validation: Check coverage
    print("\n--- Coverage Analysis ---")

    markets_tested = set(t.market for t in sample_tests)
    all_markets = set(r.get('market') for r in ppt_data['records'] if not r.get('is_total'))

    print(f"Markets tested: {len(markets_tested)}/{len(all_markets)}")
    print(f"  Tested: {sorted(markets_tested)}")
    print(f"  Not tested: {sorted(all_markets - markets_tested)}")

    # Check edge case coverage
    edge_cases_covered = set(t.scenario for t in sample_tests)
    required_edge_cases = {'all_present', 'category_missing', 'single_record', 'large_values', 'negative_values'}
    print(f"\nEdge cases covered: {len(edge_cases_covered & required_edge_cases)}/{len(required_edge_cases)}")
    print(f"  Covered: {sorted(edge_cases_covered & required_edge_cases)}")

    missing = required_edge_cases - edge_cases_covered
    if missing:
        print(f"  Missing: {sorted(missing)}")

    # Save results
    output = {
        'summary': {
            'total_tests': len(sample_tests),
            'passed': passed,
            'failed': failed,
            'skipped': len(sample_tests) - passed - failed
        },
        'coverage': {
            'markets_tested': list(markets_tested),
            'markets_not_tested': list(all_markets - markets_tested),
            'edge_cases_covered': list(edge_cases_covered)
        },
        'tests': results
    }

    output_path = base_dir / 'output' / 'reports' / 'sampling_test_results.json'
    with open(output_path, 'w') as f:
        json.dump(output, f, indent=2)

    print(f"\nResults saved to: {output_path}")

    return failed == 0


if __name__ == '__main__':
    success = run_sampling_tests()
    sys.exit(0 if success else 1)
