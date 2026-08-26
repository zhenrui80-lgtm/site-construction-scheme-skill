import tempfile
import unittest
from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas

from scripts.render_pdf_qa import audit_pdf


class RenderPdfQaTests(unittest.TestCase):
    def make_pdf(self, path):
        writer = canvas.Canvas(str(path), pagesize=A4)
        for line in range(8):
            writer.drawString(100, 700 - line * 24, "Cover title and approval information")
        writer.showPage()
        for line in range(35):
            writer.drawString(60, 800 - line * 20, "Technical content " * 4)
        writer.showPage()
        writer.setPageSize(landscape(A4))
        for line in range(10):
            writer.drawString(60, 500 - line * 25, "Schedule row " * 8)
        writer.save()

    def test_sparse_allowlist_and_landscape_last_page(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            pdf = root / "qa.pdf"
            self.make_pdf(pdf)
            result = audit_pdf(
                pdf,
                root / "render",
                allowed_sparse_pages=[1, 3],
                sparse_ink_ratio=0.025,
                expect_landscape_last_page=True,
            )
            self.assertEqual(result["failures"], [])
            self.assertEqual(result["page_count"], 3)
            self.assertEqual(result["page_metrics"][-1]["orientation"], "landscape")


if __name__ == "__main__":
    unittest.main()
