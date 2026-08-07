from pathlib import Path

from PIL import Image, ImageChops
import pypdfium2 as pdfium


ROOT = Path(__file__).resolve().parent
STANDARDS = ROOT / "standards"
OUT = ROOT / "assets" / "standards"
OUT.mkdir(parents=True, exist_ok=True)

PAGES = [
    ("GB30871-2022_危险化学品企业特殊作业安全规范.pdf", 9, "GB30871_p09_介质处理与能量隔离.png"),
    ("GB30871-2022_危险化学品企业特殊作业安全规范.pdf", 13, "GB30871_p13_受限空间检测.png"),
    ("GB30871-2022_危险化学品企业特殊作业安全规范.pdf", 16, "GB30871_p16_吊装作业要求.png"),
    ("GB50212-2014_建筑防腐蚀工程施工规范.pdf", 16, "GB50212_p16_混凝土基层要求.png"),
    ("GB50212-2014_建筑防腐蚀工程施工规范.pdf", 28, "GB50212_p28_纤维增强塑料施工.png"),
    ("GB50212-2014_建筑防腐蚀工程施工规范.pdf", 31, "GB50212_p31_养护与质量检查.png"),
    ("GB50168-2018_电缆线路施工及验收标准.pdf", 19, "GB50168_p19_电缆支架桥架安装.png"),
    ("GB50168-2018_电缆线路施工及验收标准.pdf", 38, "GB50168_p38_电缆防火封堵.png"),
]


def crop_white(image: Image.Image) -> Image.Image:
    rgb = image.convert("RGB")
    bg = Image.new("RGB", rgb.size, "white")
    diff = ImageChops.difference(rgb, bg).convert("L")
    diff = diff.point(lambda value: 0 if value < 10 else 255)
    bbox = diff.getbbox()
    if not bbox:
        return rgb
    left, top, right, bottom = bbox
    pad = 36
    return rgb.crop(
        (
            max(0, left - pad),
            max(0, top - pad),
            min(rgb.width, right + pad),
            min(rgb.height, bottom + pad),
        )
    )


for pdf_name, page_number, out_name in PAGES:
    pdf_path = STANDARDS / pdf_name
    document = pdfium.PdfDocument(str(pdf_path))
    page = document[page_number - 1]
    bitmap = page.render(scale=2.6, rotation=0)
    image = crop_white(bitmap.to_pil())
    target = OUT / out_name
    image.save(target, format="PNG", optimize=True)
    print(f"{pdf_name} page {page_number} -> {target} {image.size}")
