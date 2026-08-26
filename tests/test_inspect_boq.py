import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "inspect_boq.py"


class InspectBoqTests(unittest.TestCase):
    def test_formula_and_cached_columns_are_emitted(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook_path = Path(directory) / "boq.xlsx"
            output = Path(directory) / "boq.json"
            workbook = Workbook()
            sheet = workbook.active
            sheet["A1"] = 1
            sheet["B1"] = 2
            sheet["C1"] = "=SUM(A1:B1)"
            sheet.row_dimensions[2].hidden = True
            workbook.save(workbook_path)
            result = subprocess.run(
                [sys.executable, str(SCRIPT), str(workbook_path), "--output", str(output)],
                capture_output=True,
                text=True,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
            payload = json.loads(output.read_text(encoding="utf-8"))["sheets"][0]
            self.assertEqual(payload["formula_cells"], 1)
            self.assertIn("C1", payload["formulas_without_cache"])
            self.assertIn("cached_values", payload["rows"][0])
            self.assertIn(2, payload["hidden_rows"])


if __name__ == "__main__":
    unittest.main()
