from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parent
NORMS = ROOT / "norms"
ASSETS = ROOT / "assets"


def crop(source_name, output_name, box):
    with Image.open(NORMS / source_name) as source:
        source.convert("RGB").crop(box).save(ASSETS / output_name, quality=95)


# Retain only the official announcement body and its implementation statement.
crop(
    "GB55024-2022_effective.png",
    "规范查新_GB55024-2022_正文.png",
    (70, 65, 930, 720),
)

# Retain the standard title, status badges and drafting information; omit the
# navigation header, sharing toolbar and the unused lower page area.
crop(
    "GB6222-2025_status.png",
    "规范查新_GB6222-2025_现行状态.png",
    (55, 205, 1185, 505),
)

print(ASSETS / "规范查新_GB55024-2022_正文.png")
print(ASSETS / "规范查新_GB6222-2025_现行状态.png")
