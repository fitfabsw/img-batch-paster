from __future__ import annotations

import io
from pathlib import Path

import click
from flask import Flask, jsonify, request, send_file, send_from_directory
from PIL import Image

from ..grouper import scan_folder
from ..pptx_writer import Placement, write_pages, write_placements
from ..keynote_export import convert_pptx_to_key
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

    def _to_pl(p):
        return Placement(
            path=Path(p["path"]),
            x_cm=float(p["x_cm"]), y_cm=float(p["y_cm"]),
            w_cm=float(p["w_cm"]), h_cm=float(p["h_cm"]),
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

    return jsonify({"output": str(out_path), "pages": len(pages)})


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
