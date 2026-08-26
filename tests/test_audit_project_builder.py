import sys
import unittest
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from audit_project_builder import audit_source


class ProjectBuilderAuditTests(unittest.TestCase):
    def test_stable_helper_import_passes(self):
        source = "from docx_builder import add_body_paragraph\n"
        self.assertEqual(audit_source(source), [])

    def test_previous_project_loader_fails(self):
        source = '''
import importlib.util
PREVIOUS_UTILS = r"C:\\repo\\work\\old-project\\build_scheme.py"
spec = importlib.util.spec_from_file_location("old", PREVIOUS_UTILS)
'''
        failures = audit_source(source)
        self.assertTrue(any("PREVIOUS_UTILS" in item for item in failures))
        self.assertTrue(any("build_scheme.py" in item for item in failures))


if __name__ == "__main__":
    unittest.main()
