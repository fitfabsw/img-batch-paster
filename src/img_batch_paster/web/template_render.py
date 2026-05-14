"""Render the first slide of a .pptx to a PNG via LibreOffice (soffice)."""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

CACHE_DIR = Path(tempfile.gettempdir()) / "img-batch-paster-cache"
CACHE_DIR.mkdir(exist_ok=True)


def _cache_key(path: Path) -> str:
    h = hashlib.sha1(f"{path.resolve()}|{path.stat().st_mtime_ns}".encode()).hexdigest()[:16]
    return h


def slide_size_cm(pptx_path: Path) -> tuple[float, float]:
    prs = Presentation(str(pptx_path))
    return (Emu(prs.slide_width).cm, Emu(prs.slide_height).cm)


def render_first_slide(pptx_path: Path) -> Path:
    """Return path to a cached PNG for the first slide of pptx_path."""
    key = _cache_key(pptx_path)
    out_png = CACHE_DIR / f"{key}.png"
    if out_png.exists():
        return out_png

    soffice = shutil.which("soffice") or "/usr/local/bin/soffice"
    if not Path(soffice).exists():
        raise RuntimeError("找不到 soffice (LibreOffice)，無法渲染範本預覽")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "png", "--outdir", str(tmp_path), str(pptx_path)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(f"soffice 失敗: {result.stderr or result.stdout}")
        produced = list(tmp_path.glob("*.png"))
        if not produced:
            raise RuntimeError("soffice 未產出 PNG")
        shutil.move(str(produced[0]), out_png)
    return out_png
