"""Render the first slide of a .pptx to a PNG via LibreOffice (soffice)."""
from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
import threading
from pathlib import Path

from pptx import Presentation
from pptx.util import Emu

CACHE_DIR = Path(tempfile.gettempdir()) / "img-batch-paster-cache"
CACHE_DIR.mkdir(exist_ok=True)

# LibreOffice 持久 user profile：第一次冷啟仍慢，但後續呼叫快很多
_LO_PROFILE = Path(tempfile.gettempdir()) / "img-batch-paster-lo-profile"
_warm_started = False
_warm_lock = threading.Lock()


def _cache_key(path: Path) -> str:
    h = hashlib.sha1(f"{path.resolve()}|{path.stat().st_mtime_ns}".encode()).hexdigest()[:16]
    return h


def slide_size_cm(pptx_path: Path) -> tuple[float, float]:
    prs = Presentation(str(pptx_path))
    return (Emu(prs.slide_width).cm, Emu(prs.slide_height).cm)


def has_soffice() -> bool:
    if shutil.which("soffice"):
        return True
    return Path("/usr/local/bin/soffice").exists() or Path("/opt/homebrew/bin/soffice").exists() \
        or Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists()


def render_first_slide(pptx_path: Path) -> Path:
    """Return path to a cached PNG for the first slide of pptx_path."""
    import sys as _sys
    key = _cache_key(pptx_path)
    out_png = CACHE_DIR / f"{key}.png"
    if out_png.exists():
        return out_png

    # macOS: 優先用 Keynote (daemon-warm 後可以 ~1-2s)
    if _sys.platform == "darwin":
        try:
            from ..keynote_export import render_pptx_via_keynote
            kn_png = render_pptx_via_keynote(pptx_path)
            if kn_png and kn_png.is_file():
                shutil.move(str(kn_png), out_png)
                return out_png
        except Exception as _e:
            print(f"[render] Keynote fallback → LO: {_e}", file=_sys.stderr)

    # fallback: LibreOffice
    soffice = (
        shutil.which("soffice")
        or ("/usr/local/bin/soffice" if Path("/usr/local/bin/soffice").exists() else None)
        or ("/opt/homebrew/bin/soffice" if Path("/opt/homebrew/bin/soffice").exists() else None)
        or ("/Applications/LibreOffice.app/Contents/MacOS/soffice"
            if Path("/Applications/LibreOffice.app/Contents/MacOS/soffice").exists() else None)
    )
    if not soffice or not Path(soffice).exists():
        raise RuntimeError("找不到 soffice (LibreOffice)，無法渲染範本預覽")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "png",
             "--outdir", str(tmp_path),
             f"-env:UserInstallation=file://{_LO_PROFILE}",
             str(pptx_path)],
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode != 0:
            raise RuntimeError(f"soffice 失敗: {result.stderr or result.stdout}")
        produced = list(tmp_path.glob("*.png"))
        if not produced:
            raise RuntimeError("soffice 未產出 PNG")
        shutil.move(str(produced[0]), out_png)
    return out_png


def prewarm_libreoffice(default_pptx: Path, periodic_interval_sec: int = 300) -> None:
    """背景啟動 LibreOffice 把 profile 暖起來，並啟動定期 prewarm 保持 OS page cache 熱。
    macOS 同時也 prewarm Keynote (daemon)，因為 render_first_slide 會優先用 Keynote。
    第一次：建 profile + 暖 OS cache
    後續：每 periodic_interval_sec (預設 5 分鐘) 跑一次 dummy render
    """
    global _warm_started
    with _warm_lock:
        if _warm_started:
            return
        _warm_started = True

    # macOS: 啟動 Keynote daemon (背景 alive)
    import sys as _sys
    if _sys.platform == "darwin":
        try:
            from ..keynote_export import prewarm_keynote
            print("[prewarm] starting Keynote daemon (open -ga Keynote)", flush=True)
            prewarm_keynote()
            print("[prewarm] Keynote prewarm call returned", flush=True)
        except Exception as _e:
            print(f"[prewarm] Keynote prewarm failed: {_e}", flush=True)

    if not has_soffice() or not default_pptx.is_file():
        return

    soffice_path = _find_soffice()

    def _initial():
        # 第一次：透過 render_first_slide 順便把 default.pptx 的 PNG 也 cache 起來
        try:
            render_first_slide(default_pptx)
        except Exception:
            pass

    def _periodic():
        # 持續跑 throwaway render，繞過自家 PNG cache (path+mtime 命中會 early-return)
        # 用 default.pptx 但 render 到丟棄的 tmp dir 而不是經過 render_first_slide
        if not soffice_path:
            return
        import time
        while True:
            time.sleep(periodic_interval_sec)
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    subprocess.run(
                        [soffice_path, "--headless", "--convert-to", "png",
                         "--outdir", tmp,
                         f"-env:UserInstallation=file://{_LO_PROFILE}",
                         str(default_pptx)],
                        capture_output=True, timeout=60,
                    )
            except Exception:
                pass

    threading.Thread(target=_initial, daemon=True).start()
    threading.Thread(target=_periodic, daemon=True).start()


def _find_soffice() -> str | None:
    """找 soffice 執行檔路徑。重複了 render_first_slide 內的邏輯，方便分享。"""
    p = shutil.which("soffice")
    if p:
        return p
    for cand in ("/usr/local/bin/soffice", "/opt/homebrew/bin/soffice",
                 "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if Path(cand).exists():
            return cand
    return None
