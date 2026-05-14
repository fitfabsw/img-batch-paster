"""產生煙霧測試用的圖片：xxx_1..3, yyy_1..3, zzz_1..3"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).parent / "fixtures" / "images"
OUT.mkdir(parents=True, exist_ok=True)

COLORS = {"xxx": (220, 80, 80), "yyy": (80, 180, 80), "zzz": (80, 120, 220)}

for group, color in COLORS.items():
    for n in range(1, 4):
        img = Image.new("RGB", (400, 300), color)
        d = ImageDraw.Draw(img)
        d.text((20, 20), f"{group}_{n}", fill="white")
        img.save(OUT / f"{group}_{n}.png")

print(f"wrote 9 images to {OUT}")
