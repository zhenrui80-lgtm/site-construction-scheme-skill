from pathlib import Path
import sys

import pypdfium2 as pdfium
from PIL import Image, ImageDraw


pdf_path = Path(sys.argv[1])
output_dir = Path(sys.argv[2])
output_dir.mkdir(parents=True, exist_ok=True)
rendered = output_dir / "rendered"
rendered.mkdir(exist_ok=True)

doc = pdfium.PdfDocument(str(pdf_path))
pages = []
for index in range(len(doc)):
    image = doc[index].render(scale=1.4).to_pil().convert("RGB")
    path = rendered / f"page-{index + 1:03d}.png"
    image.save(path)
    pages.append(path)

contacts = []
for batch_index, start in enumerate(range(0, len(pages), 20), start=1):
    batch = pages[start : start + 20]
    tile_w, tile_h = 260, 370
    sheet = Image.new("RGB", (tile_w * 5, tile_h * 4), "#d6d6d6")
    draw = ImageDraw.Draw(sheet)
    for local_index, path in enumerate(batch):
        image = Image.open(path).convert("RGB")
        image.thumbnail((tile_w - 12, tile_h - 28), Image.Resampling.LANCZOS)
        x = (local_index % 5) * tile_w + (tile_w - image.width) // 2
        y = (local_index // 5) * tile_h + 22
        sheet.paste(image, (x, y))
        draw.text(
            (local_index % 5 * tile_w + 8, local_index // 5 * tile_h + 4),
            f"P{start + local_index + 1}",
            fill="black",
        )
    output = output_dir / f"contact_{batch_index}.png"
    sheet.save(output)
    contacts.append(output)

print(f"pages={len(pages)}")
print("contacts=" + ";".join(str(path) for path in contacts))
