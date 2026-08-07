from pathlib import Path

from pypdf import PdfReader


BASE = Path(__file__).resolve().parent / "standards"
KEYWORDS = [
    "热熔",
    "玻璃钢",
    "环氧",
    "受限空间",
    "吊装作业",
    "临时用电",
    "起重",
    "电缆桥架",
    "接地",
    "高处作业",
    "防腐蚀",
]


for pdf in sorted(BASE.glob("*.pdf")):
    try:
        reader = PdfReader(str(pdf))
    except Exception as exc:
        print(f"\n## {pdf.name}\nERROR: {exc}")
        continue

    print(f"\n## {pdf.name}\npages={len(reader.pages)}")
    for number, page in enumerate(reader.pages, start=1):
        try:
            text = (page.extract_text() or "").replace("\x00", "")
        except Exception:
            continue
        hits = [keyword for keyword in KEYWORDS if keyword in text]
        if not hits:
            continue
        compact = " ".join(text.split())
        print(f"page={number} hits={','.join(hits)} text={compact[:260]}")
