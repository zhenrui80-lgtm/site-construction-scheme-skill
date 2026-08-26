import tempfile
import unittest
from pathlib import Path

from docx import Document

from scripts.audit_schedule import audit


class ScheduleAuditTests(unittest.TestCase):
    def make_docx(self, text):
        directory = tempfile.TemporaryDirectory()
        path = Path(directory.name) / "scheme.docx"
        document = Document()
        document.add_paragraph(text)
        document.save(path)
        return directory, path

    def test_fixed_dates_pass(self):
        directory, path = self.make_docx("计划2026年8月10日至2026年8月12日施工，共3日历天。")
        self.addCleanup(directory.cleanup)
        result = audit(path, "fixed_dates", 3, "2026-08-10", "2026-08-12")
        self.assertEqual(result["failures"], [])

    def test_day_index_requires_final_day(self):
        directory, path = self.make_docx("本工程共10日历天，但未列日序。")
        self.addCleanup(directory.cleanup)
        result = audit(path, "day_index", 10)
        self.assertTrue(any("D10" in failure for failure in result["failures"]))


if __name__ == "__main__":
    unittest.main()
