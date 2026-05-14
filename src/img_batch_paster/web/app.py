from __future__ import annotations

import io
from pathlib import Path

import click
from flask import Flask, jsonify, request, send_file, send_from_directory
from PIL import Image

from ..grouper import scan_folder
from ..pptx_writer import Placement, write_placements
from .template_render import render_first_slide, slide_size_cm

STATIC_DIR = Path(__file__).parent / "static"

app = Flask(__name__, static_folder=str(STATIC_DIR), static_url_path="/static")


@app.get("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.post("/api/scan")
def api_scan():
    data = request.get_json(force=True)
    folder = Path(data["folder"]).expanduser().resolve()
    pattern = data.get("pattern", "{group}_{n}")
    extensions = data.get("extensions", [".png", ".jpg", ".jpeg"])
    cols = int(data.get("cols", 3))

    if not folder.is_dir():
        return jsonify({"error": f"資料夾不存在: {folder}"}), 400

    grouped = scan_folder(folder, pattern, extensions, cols)

    def _cell(p):
        if p is None:
            return None
        try:
            with Image.open(p) as im:
                w, h = im.size
        except Exception:
            w, h = 0, 0
        return {"path": str(p), "w": w, "h": h}

    rows = [
        {"group": name, "cells": [_cell(p) for p in row]}
        for name, row in grouped.rows
    ]
    return jsonify({"folder": str(folder), "cols": grouped.cols, "rows": rows})


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


@app.post("/api/template/load")
def api_template_load():
    data = request.get_json(force=True)
    path = Path(data["path"]).expanduser().resolve()
    if not path.is_file():
        return jsonify({"error": f"檔案不存在: {path}"}), 400
    try:
        w_cm, h_cm = slide_size_cm(path)
        png = render_first_slide(path)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    return jsonify({
        "path": str(path),
        "width_cm": w_cm,
        "height_cm": h_cm,
        "preview_url": f"/api/template/preview?key={png.stem}",
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
    slide = data["slide"]
    out_path = Path(data["output"]["path"]).expanduser().resolve()
    template = data["output"].get("template")
    template_path = Path(template).expanduser().resolve() if template else None

    placements = [
        Placement(
            path=Path(p["path"]),
            x_cm=float(p["x_cm"]),
            y_cm=float(p["y_cm"]),
            w_cm=float(p["w_cm"]),
            h_cm=float(p["h_cm"]),
        )
        for p in data["placements"]
    ]
    if not placements:
        return jsonify({"error": "沒有可匯出的圖片"}), 400

    out = write_placements(
        float(slide["width_cm"]), float(slide["height_cm"]),
        placements, out_path, template_path,
    )
    return jsonify({"output": str(out)})


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
