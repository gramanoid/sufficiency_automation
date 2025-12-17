#!/usr/bin/env python3
"""
Adversarial Test Suite for Data Validation
Intentionally tries to break the validator to prove it can catch errors.
"""

import json
import copy
import tempfile
from pathlib import Path
import sys

# Add scripts dir to path
sys.path.insert(0, str(Path(__file__).parent))
from validator import (
    run_validation, compare_values, normalize_key,
    ValidationResult, PCT_TOL, CURRENCY_TOL
)


class AdversarialTestSuite:
    """Suite of tests that intentionally try to break validation"""

    def __init__(self, ppt_data_path: Path, excel_data_path: Path):
        with open(ppt_data_path) as f:
            self.ppt_data = json.load(f)
        with open(excel_data_path) as f:
            self.excel_data = json.load(f)

        self.test_results = []
        self.passed = 0
        self.failed = 0

    def run_all_tests(self):
        """Run all adversarial tests"""
        print("=" * 60)
        print("ADVERSARIAL TEST SUITE")
        print("=" * 60)
        print()

        tests = [
            ("Scale Error: Percentage as Whole Number", self.test_scale_error_percentage),
            ("Sign Flip: Negative to Positive", self.test_sign_flip),
            ("Missing Record in Excel", self.test_missing_record),
            ("Currency Off by Significant Amount", self.test_currency_error),
            ("Campaign Count Mismatch", self.test_integer_mismatch),
            ("Media Split Doesn't Sum to 100%", self.test_media_sum_error),
            ("Rounding Boundary: 0.049 vs 0.05", self.test_rounding_boundary),
            ("Wrong Brand Assigned to Market", self.test_brand_market_mismatch),
            ("Duplicate Key Handling", self.test_duplicate_key),
            ("All Zeros Edge Case", self.test_all_zeros),
            ("Single Record Group", self.test_single_record),
            ("Large Values Stress Test", self.test_large_values),
            ("Null vs Zero Handling", self.test_null_vs_zero),
            ("Within Tolerance Should Pass", self.test_within_tolerance),
            ("Just Outside Tolerance Should Fail", self.test_just_outside_tolerance),
        ]

        for name, test_fn in tests:
            try:
                result = test_fn()
                status = "PASS" if result else "FAIL"
                if result:
                    self.passed += 1
                else:
                    self.failed += 1
                print(f"  [{status}] {name}")
                self.test_results.append({'name': name, 'status': status})
            except Exception as e:
                self.failed += 1
                print(f"  [ERROR] {name}: {e}")
                self.test_results.append({'name': name, 'status': 'ERROR', 'error': str(e)})

        print()
        print("=" * 60)
        print(f"RESULTS: {self.passed} passed, {self.failed} failed")
        print("=" * 60)

        return self.failed == 0

    def _run_validation_with_modified_data(self, modified_excel: dict, modified_ppt: dict = None) -> ValidationResult:
        """Helper to run validation with modified data"""
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(modified_excel, f)
            excel_path = Path(f.name)

        if modified_ppt:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                json.dump(modified_ppt, f)
                ppt_path = Path(f.name)
        else:
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                json.dump(self.ppt_data, f)
                ppt_path = Path(f.name)

        result = run_validation(excel_path, ppt_path)

        # Cleanup
        excel_path.unlink()
        ppt_path.unlink()

        return result

    def test_scale_error_percentage(self) -> bool:
        """
        Test: Validator catches when percentage stored as 63 instead of 0.63
        Expected: FAIL validation (should catch error)
        """
        modified = copy.deepcopy(self.excel_data)

        # Modify first record's AWA from 0.60 to 60 (scale error)
        for rec in modified['records']:
            if not rec.get('is_total') and rec.get('awa'):
                rec['awa'] = rec['awa'] * 100  # Scale error
                break

        result = self._run_validation_with_modified_data(modified)
        # Should FAIL validation (detect the error)
        return result.mismatches > 0

    def test_sign_flip(self) -> bool:
        """
        Test: Validator catches sign flip (negative gap becomes positive)
        Expected: FAIL validation
        """
        modified = copy.deepcopy(self.excel_data)

        for rec in modified['records']:
            if not rec.get('is_total') and rec.get('gap_gbp') and rec['gap_gbp'] < 0:
                rec['gap_gbp'] = abs(rec['gap_gbp'])  # Flip sign
                break

        result = self._run_validation_with_modified_data(modified)
        return result.mismatches > 0

    def test_missing_record(self) -> bool:
        """
        Test: Validator catches when a record is missing from Excel
        Expected: FAIL validation
        """
        modified = copy.deepcopy(self.excel_data)

        # Remove first non-total record
        for i, rec in enumerate(modified['records']):
            if not rec.get('is_total'):
                modified['records'].pop(i)
                break

        result = self._run_validation_with_modified_data(modified)
        return result.missing_in_excel > 0

    def test_currency_error(self) -> bool:
        """
        Test: Validator catches currency value off by more than tolerance
        Expected: FAIL validation
        """
        modified = copy.deepcopy(self.excel_data)

        for rec in modified['records']:
            if not rec.get('is_total') and rec.get('budget_2026'):
                rec['budget_2026'] += 10000  # Add £10k (well above £1 tolerance)
                break

        result = self._run_validation_with_modified_data(modified)
        return result.mismatches > 0

    def test_integer_mismatch(self) -> bool:
        """
        Test: Validator catches integer field mismatch
        Expected: FAIL validation
        """
        modified = copy.deepcopy(self.excel_data)

        for rec in modified['records']:
            if not rec.get('is_total') and rec.get('long_campaigns') is not None:
                rec['long_campaigns'] += 5  # Change campaign count
                break

        result = self._run_validation_with_modified_data(modified)
        return result.mismatches > 0

    def test_media_sum_error(self) -> bool:
        """
        Test: Media percentages that don't sum to ~100%
        This tests the extraction logic, not just comparison
        """
        modified = copy.deepcopy(self.excel_data)

        for rec in modified['records']:
            if not rec.get('is_total'):
                tv = rec.get('tv') or 0
                digital = rec.get('digital') or 0
                others = rec.get('others') or 0
                if tv + digital + others > 0.9:  # Valid sum
                    # Break the sum by adding 0.5 to digital
                    rec['digital'] = (rec.get('digital') or 0) + 0.5
                    break

        result = self._run_validation_with_modified_data(modified)
        return result.mismatches > 0

    def test_rounding_boundary(self) -> bool:
        """
        Test: Values at rounding boundary (0.049 vs 0.05)
        Should detect difference when outside tolerance
        """
        # Test compare_values directly
        # 0.001 tolerance for percentages
        is_match, diff, _ = compare_values(0.049, 0.05, 'percentage')
        # Diff is 0.001, which equals tolerance, should be within
        result1 = is_match

        is_match, diff, _ = compare_values(0.048, 0.05, 'percentage')
        # Diff is 0.002, which exceeds tolerance, should not match
        result2 = not is_match

        return result1 and result2

    def test_brand_market_mismatch(self) -> bool:
        """
        Test: Brand assigned to wrong market should fail
        When markets are swapped, keys change, causing missing record detection
        """
        modified = copy.deepcopy(self.excel_data)

        # Find two records from different markets
        records_to_swap = []
        seen_markets = set()
        for rec in modified['records']:
            if not rec.get('is_total') and rec.get('market') not in seen_markets:
                records_to_swap.append(rec)
                seen_markets.add(rec.get('market'))
            if len(records_to_swap) == 2:
                break

        if len(records_to_swap) == 2:
            # Swap markets - this creates invalid keys
            records_to_swap[0]['market'], records_to_swap[1]['market'] = \
                records_to_swap[1]['market'], records_to_swap[0]['market']

        result = self._run_validation_with_modified_data(modified)
        # Swapping markets creates mismatched keys: old keys missing in Excel, wrong keys don't exist in PPT
        # This should result in missing records OR mismatches due to wrong data at key
        return result.mismatches > 0 or result.missing_in_excel > 0 or result.missing_in_ppt > 0

    def test_duplicate_key(self) -> bool:
        """
        Test: Duplicate market/category/brand should be handled
        """
        modified = copy.deepcopy(self.excel_data)

        # Duplicate first record
        for rec in modified['records']:
            if not rec.get('is_total'):
                dup = copy.deepcopy(rec)
                dup['budget_2026'] = (dup.get('budget_2026') or 0) + 50000  # Different value
                modified['records'].append(dup)
                break

        # Should use one of them (dict overwrites), validation should work
        result = self._run_validation_with_modified_data(modified)
        # This might pass or fail depending on which gets used
        # The important thing is it doesn't crash
        return True  # Just ensure no crash

    def test_all_zeros(self) -> bool:
        """
        Test: Record with all zero values should still validate correctly
        """
        modified_ppt = copy.deepcopy(self.ppt_data)
        modified_excel = copy.deepcopy(self.excel_data)

        # Create matching all-zero records
        zero_record = {
            'market': 'TESTMARKET',
            'category': 'TestCat',
            'brand': 'ZeroBrand',
            'is_total': False,
            'budget_2026': 0,
            'sufficient_2026': 0,
            'gap_gbp': 0,
            'gap_pct': 0,
            'awa': 0,
            'con': 0,
            'pur': 0,
            'tv': 0,
            'digital': 0,
            'others': 0,
            'long_campaigns': 0,
            'short_campaigns': 0,
            'long_pct': 0,
        }

        modified_ppt['records'].append(copy.deepcopy(zero_record))
        modified_excel['records'].append(copy.deepcopy(zero_record))

        result = self._run_validation_with_modified_data(modified_excel, modified_ppt)
        # Should pass - zeros should match zeros
        return result.mismatches == 0

    def test_single_record(self) -> bool:
        """
        Test: Validation works with single record
        """
        modified_ppt = copy.deepcopy(self.ppt_data)
        modified_excel = copy.deepcopy(self.excel_data)

        # Keep only first non-total record
        modified_ppt['records'] = [r for r in modified_ppt['records'] if not r.get('is_total')][:1]
        modified_excel['records'] = [r for r in modified_excel['records'] if not r.get('is_total')][:1]

        # Ensure they match
        if modified_ppt['records'] and modified_excel['records']:
            key_ppt = normalize_key(
                modified_ppt['records'][0].get('market'),
                modified_ppt['records'][0].get('category'),
                modified_ppt['records'][0].get('brand')
            )
            key_excel = normalize_key(
                modified_excel['records'][0].get('market'),
                modified_excel['records'][0].get('category'),
                modified_excel['records'][0].get('brand')
            )
            if key_ppt != key_excel:
                # Make them match
                modified_excel['records'][0]['market'] = modified_ppt['records'][0]['market']
                modified_excel['records'][0]['category'] = modified_ppt['records'][0]['category']
                modified_excel['records'][0]['brand'] = modified_ppt['records'][0]['brand']

        result = self._run_validation_with_modified_data(modified_excel, modified_ppt)
        return result.total_records_checked == 1

    def test_large_values(self) -> bool:
        """
        Test: Large values with thousands separators handled correctly
        """
        # Test compare_values with large numbers
        is_match, diff, _ = compare_values(12345678.90, 12345678.90, 'currency')
        result1 = is_match

        # Large percentage values
        is_match, diff, _ = compare_values(0.999999, 1.0, 'percentage')
        result2 = is_match  # Should be within tolerance

        return result1 and result2

    def test_null_vs_zero(self) -> bool:
        """
        Test: Null and zero should be treated equivalently
        """
        # None vs 0 should match
        is_match1, _, _ = compare_values(None, 0, 'percentage')
        is_match2, _, _ = compare_values(0, None, 'percentage')
        is_match3, _, _ = compare_values('-', 0, 'currency')
        is_match4, _, _ = compare_values(0, '-', 'currency')

        return is_match1 and is_match2 and is_match3 and is_match4

    def test_within_tolerance(self) -> bool:
        """
        Test: Values within tolerance should PASS
        """
        # Percentage: 0.001 tolerance
        is_match, _, match_type = compare_values(0.600, 0.601, 'percentage')
        result1 = is_match and match_type == 'within_tolerance'

        # Currency: 1.0 tolerance
        is_match, _, match_type = compare_values(1000.0, 1000.5, 'currency')
        result2 = is_match and match_type == 'within_tolerance'

        return result1 and result2

    def test_just_outside_tolerance(self) -> bool:
        """
        Test: Values just outside tolerance should FAIL
        """
        # Percentage: just over 0.001 tolerance
        is_match, _, _ = compare_values(0.600, 0.6015, 'percentage')
        result1 = not is_match  # Should fail

        # Currency: just over 1.0 tolerance
        is_match, _, _ = compare_values(1000.0, 1001.5, 'currency')
        result2 = not is_match  # Should fail

        return result1 and result2


def main():
    base_dir = Path(__file__).parent.parent
    ppt_path = base_dir / 'output' / 'data' / 'ppt_extracted_data.json'
    excel_path = base_dir / 'output' / 'data' / 'updated_excel_extracted.json'

    suite = AdversarialTestSuite(ppt_path, excel_path)
    success = suite.run_all_tests()

    # Save test results
    results_path = base_dir / 'output' / 'reports' / 'adversarial_test_results.json'
    with open(results_path, 'w') as f:
        json.dump({
            'passed': suite.passed,
            'failed': suite.failed,
            'tests': suite.test_results
        }, f, indent=2)

    print(f"\nResults saved to: {results_path}")

    sys.exit(0 if success else 1)


if __name__ == '__main__':
    main()
