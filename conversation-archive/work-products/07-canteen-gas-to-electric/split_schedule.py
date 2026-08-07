from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT / "spreadsheet" / "renders" / "施工进度计划.png"
OUTPUT_DIR = ROOT / "assets"

with Image.open(SOURCE) as source:
    source = source.convert("RGB")
    width, height = source.size

    # The fixed information columns end near x=1297. Repeat them on both
    # appendix sheets so every time segment remains readable in isolation.
    fixed_end = 1297
    timeline_split = 1960

    first = source.crop((0, 0, timeline_split, height))
    first.save(OUTPUT_DIR / "施工进度计划_8月至10月.png", quality=95)

    fixed = source.crop((0, 0, fixed_end, height))
    late_timeline = source.crop((timeline_split, 0, width, height))
    second = Image.new("RGB", (fixed.width + late_timeline.width, height), "white")
    second.paste(fixed, (0, 0))
    second.paste(late_timeline, (fixed.width, 0))
    second.save(OUTPUT_DIR / "施工进度计划_10月至12月.png", quality=95)

print(OUTPUT_DIR / "施工进度计划_8月至10月.png")
print(OUTPUT_DIR / "施工进度计划_10月至12月.png")
