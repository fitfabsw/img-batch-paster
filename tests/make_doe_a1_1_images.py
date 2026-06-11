"""產生 AI DOE-A1-1(延伸) 測試用 dummy 圖。

依範例的檔名規則 {前三字}-{字母後綴}（含範例中唯一不同副檔名 2JG-K.jpeg）。
圖內標註 group-idx，且「同欄同色相、列用明暗區分」，方便核對每張是否落到正確格子。

跑法：.venv/bin/python tests/make_doe_a1_1_images.py
輸出：tests/fixtures/doe-a1-1/
"""
import colorsys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).parent / "fixtures" / "doe-a1-1"
OUT.mkdir(parents=True, exist_ok=True)

GROUPS = ["2EG", "2JG"]                                   # 前三字 → 列 (SN)
IDXS = ["C", "D", "BC", "E", "CCD", "G", "H", "TC", "I", "K", "TE"]  # 字母後綴 → 欄（依範本表頭順序）
EXT = {("2JG", "K"): "jpeg"}                              # 範例中唯一的 .jpeg

W, H = 480, 320


def font(size):
    for p in ("/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/Supplemental/Arial.ttf"):
        try:
            return ImageFont.truetype(p, size)
        except Exception:
            pass
    return ImageFont.load_default()


F_BIG, F_SMALL = font(96), font(40)

n = 0
for gi, group in enumerate(GROUPS):
    for ci, idx in enumerate(IDXS):
        hue = ci / len(IDXS)                     # 每欄一個色相
        light = 0.55 if gi == 0 else 0.38        # 列用明暗區分（2EG 亮 / 2JG 暗）
        r, g, b = (int(x * 255) for x in colorsys.hls_to_rgb(hue, light, 0.65))
        img = Image.new("RGB", (W, H), (r, g, b))
        d = ImageDraw.Draw(img)
        name = f"{group}-{idx}"
        d.text((W / 2, H / 2), idx, fill="white", anchor="mm", font=F_BIG)
        d.text((W / 2, 32), name, fill="white", anchor="mm", font=F_SMALL)
        ext = EXT.get((group, idx), "jpg")
        img.save(OUT / f"{name}.{ext}", "JPEG", quality=85)
        n += 1

print(f"wrote {n} images to {OUT}")
