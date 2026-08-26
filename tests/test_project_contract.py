import tempfile
import unittest
from pathlib import Path

from scripts.project_contract import schedule_summary, validate_contract


class ProjectContractTests(unittest.TestCase):
    def base_contract(self):
        return {
            "schema_version": 2,
            "project": {"name": "测试工程"},
            "document": {
                "template": "template.docx",
                "output_docx": "output.docx",
                "body_start": "一、编制依据及工程概况",
                "body_occurrence": 2,
                "profile": "standard",
            },
            "schedule": {
                "basis": "fixed_dates",
                "start_date": "2026-08-10",
                "finish_date": "2026-08-12",
                "duration_days": 3,
            },
            "coverage": {
                "project_terms": [f"项目{i}" for i in range(6)],
                "method_terms": [f"工法{i}" for i in range(8)],
            },
        }

    def test_inclusive_duration(self):
        self.assertEqual(schedule_summary(self.base_contract())["duration_days"], 3)

    def test_duration_mismatch_is_rejected(self):
        contract = self.base_contract()
        contract["schedule"]["duration_days"] = 2
        with self.assertRaisesRegex(ValueError, "inclusive duration"):
            schedule_summary(contract)

    def test_day_index_requires_basis_note(self):
        contract = self.base_contract()
        contract["schedule"] = {"basis": "day_index", "duration_days": 10}
        with self.assertRaisesRegex(ValueError, "basis_note"):
            schedule_summary(contract)

    def test_profile_coverage_minimums(self):
        contract = self.base_contract()
        contract["coverage"]["method_terms"] = ["拆除"]
        with tempfile.TemporaryDirectory() as directory:
            failures = validate_contract(Path(directory) / "contract.json", contract, require_files=False)
        self.assertTrue(any("method_terms" in failure for failure in failures))


if __name__ == "__main__":
    unittest.main()
