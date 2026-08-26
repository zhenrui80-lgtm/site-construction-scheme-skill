import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
ARCHIVE = ROOT / "scripts" / "archive_standard.py"
VERIFY = ROOT / "scripts" / "verify_standard_library.py"


class ArchiveStandardTests(unittest.TestCase):
    def test_archive_deduplicates_by_hash_and_merges_projects(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pdf = root / "GB99999-2026_test.pdf"
            writer = canvas.Canvas(str(pdf))
            writer.drawString(80, 700, "standard")
            writer.save()
            library_root = root / "standards"
            base = [
                sys.executable,
                str(ARCHIVE),
                str(pdf),
                "--number", "GB 99999-2026",
                "--title", "测试标准",
                "--source-url", "https://example.com/standard.pdf",
                "--source-kind", "official",
                "--library-root", str(library_root),
            ]
            first = subprocess.run([*base, "--project", "项目A"], capture_output=True, text=True)
            second = subprocess.run([*base, "--project", "项目B"], capture_output=True, text=True)
            self.assertEqual(first.returncode, 0, first.stderr)
            self.assertEqual(second.returncode, 0, second.stderr)
            index = json.loads((library_root / "index.json").read_text(encoding="utf-8"))
            self.assertEqual(len(index["standards"]), 1)
            self.assertEqual(index["standards"][0]["projects"], ["项目A", "项目B"])
            self.assertEqual(index["standards"][0]["sources"][0]["kind"], "official")
            verified = subprocess.run(
                [sys.executable, str(VERIFY), "--library-root", str(library_root)],
                capture_output=True,
                text=True,
            )
            self.assertEqual(verified.returncode, 0, verified.stdout + verified.stderr)


if __name__ == "__main__":
    unittest.main()
