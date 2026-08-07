from pathlib import Path
import sys

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
RENDER = ROOT / (sys.argv[1] if len(sys.argv) > 1 else "source_render")
PAGE_PREFIX = sys.argv[2] if len(sys.argv) > 2 else "page"
CONTACT_PREFIX = sys.argv[3] if len(sys.argv) > 3 else "contact"
PAGES = sorted(
    RENDER.glob(f"{PAGE_PREFIX}-*.png"),
    key=lambda path: int(path.stem.rsplit("-", 1)[-1]),
)

columns = 3
rows = 3
thumb_width = 520
label_height = 34
gap = 18

for sheet_index in range(0, len(PAGES), columns * rows):
    batch = PAGES[sheet_index : sheet_index + columns * rows]
    thumbs = []
    for page in batch:
        with Image.open(page) as image:
            ratio = thumb_width / image.width
            thumb = image.resize(
                (thumb_width, round(image.height * ratio)), Image.Resampling.LANCZOS
            )
            thumbs.append((page, thumb.copy()))

    max_height = max(image.height for _, image in thumbs)
    canvas = Image.new(
        "RGB",
        (
            columns * thumb_width + (columns + 1) * gap,
            rows * (max_height + label_height) + (rows + 1) * gap,
        ),
        "white",
    )
    draw = ImageDraw.Draw(canvas)
    font = ImageFont.load_default()
    for index, (page, image) in enumerate(thumbs):
        row, column = divmod(index, columns)
        x = gap + column * (thumb_width + gap)
        y = gap + row * (max_height + label_height + gap)
        canvas.paste(image, (x, y + label_height))
        number = int(page.stem.rsplit("-", 1)[-1])
        draw.text((x, y + 8), f"Page {number}", fill="black", font=font)

    output = RENDER / f"{CONTACT_PREFIX}-{sheet_index // 9 + 1:02d}.png"
    canvas.save(output, quality=92)
    print(output)
