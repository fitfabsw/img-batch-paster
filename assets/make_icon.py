"""產生 app icon (1024 PNG + macOS .icns)。

設計：圓角白底 + 2x2 的彩色圖卡網格 (代表批次貼圖)，左上角有「箭頭」表示「貼進去」。

Usage:
    python assets/make_icon.py
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter

OUT_DIR = Path(__file__).parent
ICON_PNG = OUT_DIR / "icon.png"
ICONSET = OUT_DIR / "icon.iconset"
ICNS = OUT_DIR / "icon.icns"

SIZE = 1024


def hex_(s: str) -> tuple[int, int, int, int]:
    s = s.lstrip("#")
    return (int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16), 255)


def make_icon() -> None:
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 圓角主體：深色背景做反差
    BG = hex_("#0F172A")          # slate-900
    radius = 220
    draw.rounded_rectangle([(0, 0), (SIZE - 1, SIZE - 1)], radius=radius, fill=BG)

    # 2x2 彩色圖卡 (代表批次圖片)
    PALETTE = [hex_("#F472B6"), hex_("#60A5FA"), hex_("#34D399"), hex_("#FBBF24")]  # pink / blue / green / amber
    margin = 150
    gap = 56
    card = (SIZE - 2 * margin - gap) // 2

    for r in range(2):
        for c in range(2):
            x = margin + c * (card + gap)
            y = margin + r * (card + gap)
            color = PALETTE[r * 2 + c]
            # 圖卡本身
            draw.rounded_rectangle([(x, y), (x + card, y + card)], radius=38, fill=color)
            # 卡內模擬「圖片」: 上半淺色、下半深色 + 小圓 (太陽)
            inset = 30
            ix1, iy1 = x + inset, y + inset
            ix2, iy2 = x + card - inset, y + card - inset
            mid_y = iy1 + (iy2 - iy1) * 6 // 10
            # 天空
            draw.rectangle([(ix1, iy1), (ix2, mid_y)], fill=(255, 255, 255, 220))
            # 地面 (用 color 加深一點)
            ground = tuple(max(0, c - 60) for c in color[:3]) + (255,)
            draw.rectangle([(ix1, mid_y), (ix2, iy2)], fill=ground)
            # 太陽 (右上角)
            sun_r = (ix2 - ix1) // 7
            sx = ix2 - sun_r * 2 - 12
            sy = iy1 + 12
            draw.ellipse([(sx, sy), (sx + sun_r * 2, sy + sun_r * 2)], fill=(255, 250, 200, 255))

    img.save(ICON_PNG)
    print(f"  PNG: {ICON_PNG}")


def build_icns() -> None:
    if sys.platform != "darwin":
        print("  Skip .icns build (not on macOS)")
        return
    if ICONSET.exists():
        subprocess.run(["rm", "-rf", str(ICONSET)], check=True)
    ICONSET.mkdir(parents=True)
    sizes = [
        ("16x16",   16),
        ("16x16@2x", 32),
        ("32x32",   32),
        ("32x32@2x", 64),
        ("128x128", 128),
        ("128x128@2x", 256),
        ("256x256", 256),
        ("256x256@2x", 512),
        ("512x512", 512),
        ("512x512@2x", 1024),
    ]
    for name, px in sizes:
        out = ICONSET / f"icon_{name}.png"
        subprocess.run(["sips", "-z", str(px), str(px), str(ICON_PNG), "--out", str(out)],
                       check=True, stdout=subprocess.DEVNULL)
    subprocess.run(["iconutil", "-c", "icns", str(ICONSET), "-o", str(ICNS)], check=True)
    subprocess.run(["rm", "-rf", str(ICONSET)], check=True)
    print(f"  ICNS: {ICNS}")


if __name__ == "__main__":
    print("Building icon...")
    make_icon()
    build_icns()
    print("✓ Done")
