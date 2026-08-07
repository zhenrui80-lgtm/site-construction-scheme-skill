from pathlib import Path

from docx import Document


path = Path(__file__).resolve().parent / "output" / "质检中心食堂气改电施工方案.docx"
document = Document(path)
paragraphs = document.paragraphs
print("paragraphs", len(paragraphs), "sections", len(document.sections))
for index, paragraph in enumerate(paragraphs[-18:], len(paragraphs) - 18):
    page_breaks = len(paragraph._p.xpath('.//w:br[@w:type="page"]'))
    has_section = (
        paragraph._p.pPr is not None
        and paragraph._p.pPr.sectPr is not None
    )
    drawings = len(paragraph._p.xpath(".//w:drawing"))
    print(index, repr(paragraph.text), f"page_breaks={page_breaks}", f"section={has_section}", f"drawings={drawings}")
