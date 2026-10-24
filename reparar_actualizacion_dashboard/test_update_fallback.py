import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import update_data as updater


class FiscalFallbackTests(unittest.TestCase):
    def test_missing_primary_keeps_chart_and_allows_next_update(self):
        previous_chart = {"id": "fiscal", "code": "FISC_RESULTS", "data": [{"date": "2026-08-01", "value": 0.01}]}
        previous = {"fiscal": previous_chart}
        result, errors = [], []
        updater.append_cached_chart(result, previous, errors, "FISC_RESULTS",
                                    lambda: updater.make_fiscal_gdp_chart({}, {}))
        updater.append_cached_chart(result, previous, errors, "OTHER",
                                    lambda: {"code": "OTHER", "data": [123]})
        self.assertEqual(result[0], previous_chart)
        self.assertIsNot(result[0], previous_chart)
        self.assertEqual(result[1]["data"], [123])
        self.assertIn("FISC_PRIMARY", errors[0])

    def test_unavailable_chart_without_backup_is_explicit_failure(self):
        with self.assertRaisesRegex(RuntimeError, "no existe respaldo"):
            updater.append_cached_chart([], {}, [], "FISC_RESULTS",
                                        lambda: updater.make_fiscal_gdp_chart({}, {}))

    def test_success_replaces_old_chart_without_warning(self):
        result, errors = [], []
        updater.append_cached_chart(result, {}, errors, "FISC_RESULTS",
                                    lambda: {"code": "FISC_RESULTS", "data": [456]})
        self.assertEqual(result[0]["data"], [456])
        self.assertEqual(errors, [])


if __name__ == "__main__":
    unittest.main()
