from __future__ import annotations

import io
import shutil
import tempfile
import uuid
from pathlib import Path

import click
from flask import Flask, jsonify, request, send_file, send_from_directory
from PIL import Image

from ..grouper import scan_folder
from ..pptx_writer import Placement, write_pages, write_placements
from ..xlsx_writer import CellPlacement, write_xlsx
from ..keynote_export import convert_key_to_pptx, convert_pptx_to_key
from .template_render import render_first_slide, slide_size_cm

STATIC_DIR = Path(__file__).parent / "static"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "img-batch-paster-uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024 * 1024  # 2 GB per request


def _ws_dir(ws: str) -> Path | None:
    if not ws or "/" in ws or ".." in ws:
        return None
    p = UPLOAD_DIR / ws
    return p if p.is_dir() else None


@app.get("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.post("/api/scan")
def api_scan():
    data = request.get_json(force=True)
    folder = Path(data["folder"]).expanduser().resolve()
    extensions = data.get("extensions", [".png", ".jpg", ".jpeg"])
    exts = {e.lower() for e in extensions}

    if not folder.is_dir():
        return jsonify({"error": f"資料夾不存在: {folder}"}), 400

    files = []
    for entry in sorted(folder.iterdir()):
        if not entry.is_file() or entry.suffix.lower() not in exts:
            continue
        try:
            with Image.open(entry) as im:
                w, h = im.size
        except Exception:
            w, h = 0, 0
        files.append({"path": str(entry), "name": entry.name, "w": w, "h": h})

    return jsonify({"folder": str(folder), "files": files})


@app.get("/api/thumb")
def api_thumb():
    path = request.args.get("path")
    if not path:
        return "missing path", 400
    p = Path(path)
    if not p.is_file():
        return "not found", 404
    size = int(request.args.get("size", 240))
    img = Image.open(p)
    img.thumbnail((size, size))
    buf = io.BytesIO()
    fmt = "PNG" if img.mode in ("RGBA", "P") else "JPEG"
    img.convert("RGB" if fmt == "JPEG" else img.mode).save(buf, fmt)
    buf.seek(0)
    return send_file(buf, mimetype=f"image/{fmt.lower()}")


@app.get("/api/version")
def api_version():
    import subprocess
    from .. import __version__
    project_root = Path(__file__).resolve().parents[3]
    def _git(*args):
        try:
            return subprocess.check_output(
                ["git", *args], cwd=project_root, stderr=subprocess.DEVNULL, text=True
            ).strip()
        except Exception:
            return ""
    return jsonify({
        "version": __version__,
        "commit": _git("rev-parse", "--short", "HEAD") or "?",
        "date": _git("log", "-1", "--format=%cs") or "?",
    })


@app.post("/api/workspace")
def api_workspace():
    """Create a fresh per-browser workspace under /tmp/img-batch-paster-uploads/."""
    ws = uuid.uuid4().hex[:12]
    (UPLOAD_DIR / ws / "images").mkdir(parents=True, exist_ok=True)
    return jsonify({"workspace": ws})


@app.post("/api/upload/template")
def api_upload_template():
    import sys as _sys
    ws = request.form.get("workspace", "")
    base = _ws_dir(ws)
    if not base:
        return jsonify({"error": "invalid workspace"}), 400
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "no file"}), 400

    ext = Path(f.filename).suffix.lower()
    dest_pptx = base / "template.pptx"
    dest_xlsx = base / "template.xlsx"

    if ext == ".pptx":
        f.save(str(dest_pptx))
        return jsonify({"path": str(dest_pptx), "mode": "slides"})

    if ext == ".xlsx":
        f.save(str(dest_xlsx))
        return jsonify({"path": str(dest_xlsx), "mode": "excel"})

    if ext == ".key":
        if _sys.platform != "darwin":
            return jsonify({"error": ".key 範本僅在 macOS server 可轉檔；請改傳 .pptx"}), 501
        tmp_key = base / "uploaded.key"
        f.save(str(tmp_key))
        try:
            convert_key_to_pptx(tmp_key, dest_pptx)
        except Exception as e:
            return jsonify({"error": f"Keynote → pptx 轉檔失敗: {e}"}), 500
        finally:
            if tmp_key.exists():
                tmp_key.unlink()
        return jsonify({"path": str(dest_pptx), "mode": "slides"})

    return jsonify({"error": f"不支援的副檔名: {ext}（請用 .pptx / .key / .xlsx）"}), 400


DEFAULT_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "default.pptx"


@app.post("/api/template/default")
def api_template_default():
    """Copy the bundled default template into the workspace and return its path."""
    if not DEFAULT_TEMPLATE.is_file():
        return jsonify({"error": "預設範本不存在"}), 500
    data = request.get_json(force=True, silent=True) or {}
    ws = data.get("workspace", "")
    base = _ws_dir(ws)
    if not base:
        return jsonify({"error": "invalid workspace"}), 400
    dest = base / "template.pptx"
    shutil.copy2(DEFAULT_TEMPLATE, dest)
    return jsonify({"path": str(dest), "name": "default.pptx"})


@app.post("/api/upload/images")
def api_upload_images():
    ws = request.form.get("workspace", "")
    base = _ws_dir(ws)
    if not base:
        return jsonify({"error": "invalid workspace"}), 400
    folder = base / "images"
    # Replace any previous uploads in this workspace
    if folder.exists():
        shutil.rmtree(folder)
    folder.mkdir(parents=True, exist_ok=True)

    files = request.files.getlist("files")
    saved = 0
    for f in files:
        if not f.filename:
            continue
        # 只保留檔名，剝掉路徑成分；保留 unicode（避免 secure_filename 過濾中文）
        name = Path(f.filename).name
        if not name:
            continue
        f.save(str(folder / name))
        saved += 1
    return jsonify({"folder": str(folder), "count": saved})


@app.get("/api/download/<ws>/<path:filename>")
def api_download(ws, filename):
    base = _ws_dir(ws)
    if not base:
        return "not found", 404
    p = base / filename
    if not p.is_file() or ".." in filename:
        return "not found", 404
    return send_file(str(p), as_attachment=True, download_name=Path(filename).name)


@app.post("/api/pick")
def api_pick():
    """用 macOS osascript 開原生檔案/資料夾選擇器。"""
    import subprocess
    import sys
    if sys.platform != "darwin":
        return jsonify({"error": "檔案選擇器僅在 macOS 主機可用；請手動輸入路徑"}), 501
    data = request.get_json(force=True) or {}
    kind = data.get("kind", "folder")  # "folder" | "file"
    prompt = data.get("prompt", "請選擇")

    # 起始位置：以欄位現有路徑為基準（找最近存在的祖先目錄）
    default = data.get("default", "")
    default_clause = ""
    if default:
        start = Path(default).expanduser()
        if not start.is_dir():
            start = start.parent
        while start != start.parent and not start.is_dir():
            start = start.parent
        if start.is_dir():
            default_clause = f' default location POSIX file "{start}"'

    if kind == "folder":
        script = f'POSIX path of (choose folder with prompt "{prompt}"{default_clause})'
    else:
        ext_filter = data.get("extensions")
        if ext_filter:
            of = "{" + ", ".join(f'"{e.lstrip(".")}"' for e in ext_filter) + "}"
            script = f'POSIX path of (choose file with prompt "{prompt}" of type {of}{default_clause})'
        else:
            script = f'POSIX path of (choose file with prompt "{prompt}"{default_clause})'

    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=300,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if result.returncode != 0:
        # 使用者取消會回 1；其它錯誤同樣回空字串
        return jsonify({"path": None, "cancelled": True})
    path = result.stdout.strip().rstrip("/")
    return jsonify({"path": path})


@app.post("/api/template/load")
def api_template_load():
    data = request.get_json(force=True)
    path = Path(data["path"]).expanduser().resolve()
    if not path.is_file():
        return jsonify({"error": f"檔案不存在: {path}"}), 400
    try:
        w_cm, h_cm = slide_size_cm(path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    # 渲染預覽 (需要 LibreOffice)；失敗就無背景，但仍回傳尺寸供 layout 使用
    preview_url = None
    try:
        png = render_first_slide(path)
        preview_url = f"/api/template/preview?key={png.stem}"
    except Exception:
        pass
    return jsonify({
        "path": str(path),
        "width_cm": w_cm,
        "height_cm": h_cm,
        "preview_url": preview_url,
    })


@app.post("/api/template/excel-grid")
def api_template_excel_grid():
    """Return cell-grid structure of an .xlsx template for frontend preview."""
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    data = request.get_json(force=True)
    path = Path(data["path"]).expanduser().resolve()
    if not path.is_file():
        return jsonify({"error": f"檔案不存在: {path}"}), 400
    try:
        wb = load_workbook(str(path))
    except Exception as e:
        return jsonify({"error": f"無法讀取: {e}"}), 500
    ws = wb.active
    max_col = max(ws.max_column or 1, 12)
    max_row = max(ws.max_row or 1, 20)
    DEFAULT_W = 8.43
    DEFAULT_H = 15.0

    # 收集 range-based 欄寬：openpyxl 可能用 <col min=A max=B width=W> 一次涵蓋多欄
    col_widths: dict[int, float] = {}
    for cd in ws.column_dimensions.values():
        if cd.min is None or cd.max is None or cd.width is None:
            continue
        for i in range(cd.min, cd.max + 1):
            col_widths[i] = cd.width

    # 偵測預設字體大小調整 MDW (Max Digit Width)
    try:
        sz = float(wb._fonts[0].sz or 11)
    except Exception:
        sz = 11.0
    mdw = 7.0 if sz <= 11 else 7.0 + (sz - 11) * 0.85

    cols = []
    for c in range(1, max_col + 4):
        letter = get_column_letter(c)
        w = col_widths.get(c, DEFAULT_W)
        cols.append({"letter": letter, "w_px": w * mdw + 5})
    rows = []
    for r in range(1, max_row + 4):
        rd = ws.row_dimensions.get(r)
        h = rd.height if rd is not None and rd.height is not None else DEFAULT_H
        rows.append({"r": r, "h_px": h * 4 / 3})

    cells = []
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            if cell.value is None:
                continue
            cells.append({
                "r": cell.row, "c": cell.column,
                "text": str(cell.value),
                "font_pt": float(cell.font.size or 11),
                "bold": bool(cell.font.bold),
                "h_align": cell.alignment.horizontal or "left",
                "v_align": cell.alignment.vertical or "bottom",
            })

    # Borders (簡化：只要有設定就算有；不分四邊)
    borders = []
    for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
        for cell in row:
            b = cell.border
            sides = {
                "top": bool(b.top and b.top.style),
                "right": bool(b.right and b.right.style),
                "bottom": bool(b.bottom and b.bottom.style),
                "left": bool(b.left and b.left.style),
            }
            if any(sides.values()):
                borders.append({"r": cell.row, "c": cell.column, **sides})

    return jsonify({
        "sheet": ws.title,
        "cols": cols,
        "rows": rows,
        "cells": cells,
        "borders": borders,
    })


@app.get("/api/template/preview")
def api_template_preview():
    from .template_render import CACHE_DIR
    key = request.args.get("key", "")
    p = CACHE_DIR / f"{key}.png"
    if not p.is_file():
        return "not found", 404
    return send_file(str(p), mimetype="image/png")


@app.post("/api/export")
def api_export():
    data = request.get_json(force=True)
    slide = data.get("slide", {"width_cm": 25.4, "height_cm": 14.29})

    ws = data.get("workspace") or data["output"].get("workspace")
    requested = data["output"]["path"]
    base = _ws_dir(ws) if ws else None
    if base:
        out_name = Path(requested).name or "out.pptx"
        out_path = base / out_name
    else:
        out_path = Path(requested).expanduser().resolve()

    template = data["output"].get("template")
    template_path = Path(template).expanduser().resolve() if template else None

    # Excel 分支：副檔名 .xlsx
    if out_path.suffix.lower() == ".xlsx":
        xl_placements = [
            CellPlacement(
                path=Path(p["path"]) if p.get("path") else None,
                row=int(p["row"]), col=int(p["col"]),
                span_cols=int(p.get("span_cols", 1)),
                span_rows=int(p.get("span_rows", 1)),
                text=p.get("text"),
                font_pt=float(p.get("font_pt", 12)),
            )
            for p in data.get("cells", [])
        ]
        if not xl_placements:
            return jsonify({"error": "沒有可匯出的儲存格"}), 400
        embed_in_cell = bool(data.get("embedInCell"))
        img_fit = data.get("imgFit", "cover")
        write_xlsx(xl_placements, out_path, template_path,
                   embed_in_cell=embed_in_cell, img_fit=img_fit)
        resp = {"output": str(out_path)}
        if base:
            resp["download_url"] = f"/api/download/{ws}/{out_path.name}"
        return jsonify(resp)

    def _to_pl(p):
        return Placement(
            path=Path(p["path"]) if p.get("path") else None,
            x_cm=float(p["x_cm"]), y_cm=float(p["y_cm"]),
            w_cm=float(p["w_cm"]), h_cm=float(p["h_cm"]),
            text=p.get("text"),
            font_pt=float(p.get("font_pt", 18)),
            bold=bool(p.get("bold", True)),
            align=p.get("align", "center"),
        )

    if "pages" in data:
        pages = [[_to_pl(p) for p in page] for page in data["pages"]]
    else:
        pages = [[_to_pl(p) for p in data.get("placements", [])]]

    if not pages or all(len(p) == 0 for p in pages):
        return jsonify({"error": "沒有可匯出的圖片"}), 400

    # 副檔名 .key → 先輸出 .pptx，再經由 Keynote 轉成 .key
    want_key = out_path.suffix.lower() == ".key"
    pptx_out = out_path.with_suffix(".pptx") if want_key else out_path
    write_pages(
        float(slide["width_cm"]), float(slide["height_cm"]),
        pages, pptx_out, template_path,
    )

    if want_key:
        try:
            convert_pptx_to_key(pptx_out, out_path)
        except Exception as e:
            return jsonify({
                "error": f".key 轉檔失敗，但 .pptx 已輸出至 {pptx_out}: {e}",
                "output": str(pptx_out), "pages": len(pages),
            }), 500
        # 保留中介 .pptx 供 debug；要刪可改為 pptx_out.unlink(missing_ok=True)

    resp = {"output": str(out_path), "pages": len(pages)}
    if base:
        resp["download_url"] = f"/api/download/{ws}/{out_path.name}"
    return jsonify(resp)


@click.command()
@click.option("--host", default="127.0.0.1")
@click.option("--port", default=5050, type=int)
@click.option("--desktop", is_flag=True, help="以 pywebview 桌面視窗開啟")
def run(host: str, port: int, desktop: bool) -> None:
    """啟動 Web 預覽伺服器。"""
    if desktop:
        try:
            import webview  # type: ignore
        except ImportError:
            raise SystemExit("請先安裝桌面依賴: pip install -e '.[desktop]'")
        import threading

        threading.Thread(
            target=lambda: app.run(host=host, port=port, debug=False, use_reloader=False),
            daemon=True,
        ).start()
        webview.create_window("img-batch-paster", f"http://{host}:{port}", width=1200, height=820)
        webview.start()
    else:
        click.echo(f" * 開啟 http://{host}:{port}")
        app.run(host=host, port=port, debug=True, use_reloader=False)


if __name__ == "__main__":
    run()
