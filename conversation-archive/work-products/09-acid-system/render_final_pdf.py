from pathlib import Path

import pypdfium2 as pdfium
from PIL import Image, ImageDraw, ImageFont


PDF = Path(r"C:\Users\Administrator\Desktop\Metallurgical_BOQ_AgentSkill_Fixed\agents\_work_quote_scheme\final_qa\101机组酸循环系统和酸雾系统综合改造施工方案_QA.pdf")
OUT = PDF.parent / "pages"
CONTACTS = PDF.parent / "contact_sheets"
OUT.mkdir(parents=True, exist_ok=True)
CONTACTS.mkdir(parents=True, exist_ok=True)

pdf = pdfium.PdfDocument(str(PDF))
rendered = []
for index in range(len(pdf)):
    page = pdf[index]
    bitmap = page.render(scale=1.35, rotation=0)
    image = bitmap.to_pil().convert("RGB")
    path = OUT / f"page-{index + 1:03d}.png"
    image.save(path, optimize=True)
    rendered.append(path)

font = ImageFont.load_default()
thumb_w = 240
thumb_h = 340
label_h = 20
cols = 4
rows = 4
per_sheet = cols * rows
for sheet_index in range((len(rendered) + per_sheet - 1) // per_sheet):
    subset = rendered[sheet_index * per_sheet : (sheet_index + 1) * per_sheet]
    canvas = Image.new("RGB", (cols * thumb_w, rows * (thumb_h + label_h)), "#d6d6d6")
    draw = ImageDraw.Draw(canvas)
    for local_index, path in enumerate(subset):
        image = Image.open(path).convert("RGB")
        image.thumbnail((thumb_w - 8, thumb_h - 8), Image.Resampling.LANCZOS)
        col = local_index % cols
        row = local_index // cols
        x = col * thumb_w + (thumb_w - image.width) // 2
        y = row * (thumb_h + label_h) + 4
        canvas.paste(image, (x, y))
        page_number = sheet_index * per_sheet + local_index + 1
        draw.text((col * thumb_w + 8, row * (thumb_h + label_h) + thumb_h), f"Page {page_number}", fill="black", font=font)
    canvas.save(CONTACTS / f"contact-{sheet_index + 1:02d}.png", optimize=True)

print(f"pages={len(rendered)} contacts={(len(rendered) + per_sheet - 1) // per_sheet}")
