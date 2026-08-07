from pathlib import Path

from PIL import Image, ImageOps


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"

CROPS = {
    "raw_20801.png": (145, 205, 1295, 1065, "规范查新_GBT20801.1-2025.png"),
    "raw_tsg31.png": (210, 245, 1230, 1090, "规范查新_TSG31-2025.png"),
    "raw_tsgz6002.png": (210, 245, 1230, 1060, "规范查新_TSGZ6002-2026.png"),
    "raw_50235.png": (150, 190, 1010, 570, "规范查新_GB50235-2010.png"),
    "raw_50205.png": (150, 190, 1010, 510, "规范查新_GB50205-2020.png"),
    "raw_jgjt46.png": (150, 190, 1010, 610, "规范查新_JGJT46-2024.png"),
    "raw_50168.png": (455, 10, 1225, 1090, "规范封面_GB50168-2018.png"),
}

for source_name, (left, top, right, bottom, output_name) in CROPS.items():
    with Image.open(ASSETS / source_name).convert("RGB") as image:
        cropped = image.crop((left, top, right, bottom))
        bordered = ImageOps.expand(cropped, border=2, fill="#B7B7B7")
        output = ASSETS / output_name
        bordered.save(output, quality=94)
        print(f"{output.name}|{bordered.width}x{bordered.height}")
