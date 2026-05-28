"""Stateless entry: replay UI-prepared configs against new image folders.

Designed for headless callers (MCP server, CLI scripts, CI). Two entry points:

- run_paste_job: simple grid mode with explicit params (no .ibp needed)
- run_paste_job_ibp: full .ibp replay (pptx/key/xlsx, 依檔名 順序 + idx 對位 modes)

Out of scope: 依範本 SN 多 source 模式 (needs different signature for multiple
input folders + per-source crop) — still needs web UI.
"""
from __future__ import annotations

import json
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image

from .config import Config, GridConfig, InputConfig, OutputConfig, Point, Size, SlideConfig
from .grouper import scan_folder
from .pptx_writer import Placement, write_pages, write_pptx
from .xlsx_writer import CellPlacement, write_xlsx


class IbpModeUnsupported(ValueError):
    """Raised when the .ibp config uses a mode this entry point can't replay."""


# ---------- pattern + grouping (mirror frontend extractGroupIdx/buildRows) ----------

_PLACEHOLDER_RE = re.compile(r"(\{group\}|\{idx\}|\{skip\}|\{n\})")


def _pattern_to_regex(pattern: str) -> re.Pattern[str]:
    """{group}/{idx}/{n} are named-captured; {skip} is non-capturing."""
    parts = _PLACEHOLDER_RE.split(pattern)
    out: list[str] = []
    for p in parts:
        if p == "{group}":
            out.append(r"(?P<group>.+?)")
        elif p in ("{idx}", "{n}"):
            out.append(r"(?P<idx>.+?)")
        elif p == "{skip}":
            out.append(r".+?")
        else:
            out.append(re.escape(p))
    return re.compile("^" + "".join(out) + "$")


def _extract_group_idx(filename: str, pattern: str) -> tuple[str, str | None]:
    stem = Path(filename).stem
    try:
        m = _pattern_to_regex(pattern).match(stem)
    except re.error:
        return (stem, None)
    if not m:
        return (stem, None)
    gd = m.groupdict()
    return (gd.get("group") or stem, gd.get("idx"))


def _detect_idx_list(
    files: list[Path], pattern: str,
    sort: str = "auto", custom_order: list[str] | None = None,
) -> list[str]:
    seen_order: list[str] = []
    seen: set[str] = set()
    for f in files:
        _, idx = _extract_group_idx(f.name, pattern)
        if idx is None or idx in seen:
            continue
        seen.add(idx)
        seen_order.append(idx)
    if sort == "custom":
        return list(custom_order or [])
    if sort == "alpha":
        return sorted(seen_order)
    if sort == "num":
        def _key(x: str):
            try:
                return (0, float(x))
            except ValueError:
                return (1, x)
        return sorted(seen_order, key=_key)
    # auto: 全字母 → alpha, 全數字 → num, 混合 → 首見順序
    def _is_num(s: str) -> bool:
        try:
            float(s)
            return True
        except ValueError:
            return False
    if seen_order and all(_is_num(s) for s in seen_order):
        return sorted(seen_order, key=float)
    if seen_order and all(s.isalpha() for s in seen_order):
        return sorted(seen_order)
    return seen_order


def _build_rows(
    files: list[Path], pattern: str,
    cols_or_idx_list, auto_align: bool,
) -> list[tuple[str, list[Path | None]]]:
    if auto_align:
        idx_list = list(cols_or_idx_list) if not isinstance(cols_or_idx_list, int) else []
        cols = len(idx_list)
        idx_to_col = {s: i for i, s in enumerate(idx_list)}
        order: list[str] = []
        rows_map: dict[str, list[Path | None]] = {}
        for f in files:
            group, idx = _extract_group_idx(f.name, pattern)
            if group not in rows_map:
                rows_map[group] = [None] * cols
                order.append(group)
            col = idx_to_col.get(idx) if idx is not None else None
            if col is None:
                col = next((i for i, x in enumerate(rows_map[group]) if x is None), 0)
            rows_map[group][col] = f
        return [(g, rows_map[g]) for g in order]

    cols = cols_or_idx_list if isinstance(cols_or_idx_list, int) else len(cols_or_idx_list)
    rows: list[tuple[str, list[Path | None]]] = []
    for i in range(0, len(files), cols):
        slice_ = files[i:i + cols]
        cells: list[Path | None] = list(slice_) + [None] * (cols - len(slice_))
        anchor = next((x for x in slice_ if x), None)
        group = _extract_group_idx(anchor.name, pattern)[0] if anchor else ""
        rows.append((group, cells))
    return rows


def _list_image_files(folder: Path, extensions: list[str]) -> list[Path]:
    exts = {e.lower() for e in extensions}
    return [
        p for p in sorted(folder.iterdir())
        if p.is_file() and p.suffix.lower() in exts
    ]


def _parse_excel_cell(ref: str) -> tuple[int, int]:
    """'B5' → (col=2, row=5). Both 1-based."""
    m = re.match(r"^([A-Za-z]+)(\d+)$", (ref or "").strip())
    if not m:
        return (1, 1)
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - 64)
    return (col, int(m.group(2)))


def _col_letter_to_idx(letter: str) -> int:
    s = (letter or "").upper().strip()
    if not re.match(r"^[A-Z]+$", s):
        return 0
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


# ---------- simple grid mode (no .ibp) ----------

def run_paste_job(
    image_folder: str | Path,
    output_path: str | Path,
    template: str | Path | None = None,
    pattern: str = "{group}_{n}",
    cols: int = 3,
    cell_w_cm: float = 6.0,
    cell_h_cm: float = 4.0,
    origin_x_cm: float = 2.0,
    origin_y_cm: float = 2.0,
    gap_x_cm: float = 0.3,
    gap_y_cm: float = 0.3,
    slide_w_cm: float = 25.4,
    slide_h_cm: float = 14.29,
    extensions: tuple[str, ...] = (".png", ".jpg", ".jpeg"),
) -> Path:
    """Paste images from a folder onto a slide grid and return the output path."""
    folder = Path(image_folder).expanduser().resolve()
    out = Path(output_path).expanduser().resolve()
    tpl = Path(template).expanduser().resolve() if template else None

    cfg = Config(
        slide=SlideConfig(width_cm=slide_w_cm, height_cm=slide_h_cm),
        grid=GridConfig(
            origin=Point(x_cm=origin_x_cm, y_cm=origin_y_cm),
            cell=Size(w_cm=cell_w_cm, h_cm=cell_h_cm),
            gap=Point(x_cm=gap_x_cm, y_cm=gap_y_cm),
            cols=cols,
        ),
        input=InputConfig(folder=folder, pattern=pattern, extensions=list(extensions)),
        output=OutputConfig(path=out, template=tpl),
    )
    grouped = scan_folder(cfg.input.folder, cfg.input.pattern, cfg.input.extensions, cfg.grid.cols)
    if not grouped.rows:
        raise FileNotFoundError(f"No images matching pattern '{pattern}' in {folder}")
    return write_pptx(cfg, grouped)


# ---------- .ibp replay (pptx/key/xlsx, 順序 + idx 對位) ----------

def run_paste_job_ibp(
    ibp_path: str | Path,
    image_folder: str | Path,
    output_path: str | Path,
) -> Path:
    """Replay an .ibp config bundle against a new image folder.

    Supported:
      mode  : 依檔名 + 依順序 (autoAlign=False), 依檔名 idx 對位 (autoAlign=True)
      output: .pptx, .key (macOS only), .xlsx

    Not supported (raises IbpModeUnsupported):
      mode  : 依範本 SN 多 source 模式 (needs different signature for multi-folder input)
    """
    ibp = Path(ibp_path).expanduser().resolve()
    if not ibp.is_file():
        raise FileNotFoundError(f".ibp not found: {ibp}")
    out = Path(output_path).expanduser().resolve()
    out_ext = out.suffix.lower()
    if out_ext not in (".pptx", ".key", ".xlsx"):
        raise IbpModeUnsupported(
            f"Output must be .pptx / .key / .xlsx, got '{out_ext}'."
        )
    if out_ext == ".key" and sys.platform != "darwin":
        raise IbpModeUnsupported(".key output requires macOS + Keynote.app")

    with zipfile.ZipFile(ibp, "r") as zf:
        try:
            manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
        except KeyError as e:
            raise ValueError(f"{ibp.name} missing manifest.json") from e

        mode = manifest.get("mode", {})
        if mode.get("snMatchMode"):
            raise IbpModeUnsupported(
                "依範本 SN 多 source 模式需要不同的工具 — 目前只能走 web UI。"
            )
        auto_align = bool(mode.get("autoAlign"))

        tmp_dir = Path(tempfile.mkdtemp(prefix="ibp-replay-"))
        template_path: Path | None = None
        for name in zf.namelist():
            if name.startswith("template."):
                template_path = tmp_dir / name
                with zf.open(name) as src, open(template_path, "wb") as dst:
                    dst.write(src.read())
                break

    try:
        label = manifest.get("label") or {}
        pattern = label.get("pattern") or "{group}_{n}"
        extensions = [".png", ".jpg", ".jpeg"]

        folder = Path(image_folder).expanduser().resolve()
        all_files = _list_image_files(folder, extensions)
        if not all_files:
            raise FileNotFoundError(f"No images in {folder}")

        if auto_align:
            idx_list = _detect_idx_list(
                all_files, pattern,
                sort=str(label.get("idxSort", "auto")),
                custom_order=label.get("idxOrder") or [],
            )
            if not idx_list:
                raise ValueError(
                    f"autoAlign=True but no idx captured from pattern '{pattern}' — "
                    f"pattern should contain {{idx}}."
                )
            rows = _build_rows(all_files, pattern, idx_list, auto_align=True)
        else:
            if out_ext == ".xlsx":
                excel_cfg = manifest.get("excel") or {}
                cols = max(1, int(excel_cfg.get("imgsPerRow", 4)))
            else:
                cols = max(1, int((manifest.get("grid") or {}).get("cols", 3)))
            rows = _build_rows(all_files, pattern, cols, auto_align=False)

        if not rows:
            raise FileNotFoundError(f"No images matched pattern '{pattern}' in {folder}")

        if out_ext == ".xlsx":
            return _emit_xlsx(rows, manifest, out, template_path)

        pptx_out = out if out_ext == ".pptx" else out.with_suffix(".pptx")
        _emit_pptx(rows, manifest, pptx_out, template_path)
        if out_ext == ".key":
            from .keynote_export import convert_pptx_to_key
            convert_pptx_to_key(pptx_out, out)
            return out
        return pptx_out
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _emit_pptx(
    rows: list[tuple[str, list[Path | None]]],
    manifest: dict,
    out_path: Path,
    template_path: Path | None,
) -> Path:
    grid = manifest.get("grid") or {}
    slide = manifest.get("slide") or {}
    label = manifest.get("label") or {}
    origin = grid.get("origin") or {}
    gap = grid.get("gap") or {}
    rows_per_page = max(1, int(grid.get("rows", 3)))

    slide_w = float(slide.get("width_cm", 25.4))
    slide_h = float(slide.get("height_cm", 14.29))
    ox = slide_w * float(origin.get("x", 8)) / 100.0
    oy = slide_h * float(origin.get("y", 15)) / 100.0
    cell_w = slide_w * float(grid.get("width", 25)) / 100.0
    gx = slide_w * float(gap.get("x", 2)) / 100.0
    gy = slide_h * float(gap.get("y", 3)) / 100.0

    label_enabled = bool(label.get("enabled"))
    label_x = slide_w * float(label.get("x", 2)) / 100.0
    label_w = slide_w * float(label.get("width", 12)) / 100.0
    label_font_pt = float(label.get("font_pt", 18))

    pages: list[list[Placement]] = []
    for start in range(0, len(rows), rows_per_page):
        page_rows = rows[start:start + rows_per_page]
        page_placements: list[Placement] = []
        cur_y = oy
        for group, cells in page_rows:
            anchor = next((c for c in cells if c is not None), None)
            if anchor is None:
                continue
            with Image.open(anchor) as im:
                iw, ih = im.size
            aspect = (ih / iw) if iw else 0.75
            row_h = cell_w * aspect

            if label_enabled:
                page_placements.append(Placement(
                    path=None, text=group,
                    x_cm=label_x, y_cm=cur_y, w_cm=label_w, h_cm=row_h,
                    font_pt=label_font_pt, bold=True, align="center",
                ))
            for ci, f in enumerate(cells):
                if f is None:
                    continue
                page_placements.append(Placement(
                    path=f,
                    x_cm=ox + ci * (cell_w + gx), y_cm=cur_y,
                    w_cm=cell_w, h_cm=row_h,
                ))
            cur_y += row_h + gy
        pages.append(page_placements)

    return write_pages(slide_w, slide_h, pages, out_path, template=template_path)


def _emit_xlsx(
    rows: list[tuple[str, list[Path | None]]],
    manifest: dict,
    out_path: Path,
    template_path: Path | None,
) -> Path:
    excel_cfg = manifest.get("excel") or {}
    label = manifest.get("label") or {}

    start_col, start_row = _parse_excel_cell(str(excel_cfg.get("startCell", "B5")))
    sn_col = _col_letter_to_idx(str(excel_cfg.get("snCol", "")))
    cell_cols = max(1, int(excel_cfg.get("cellCols", 1)))
    cell_rows = max(1, int(excel_cfg.get("cellRows", 1)))
    gap_rows = max(0, int(excel_cfg.get("gapRows", 0)))
    step_row = cell_rows + gap_rows

    label_enabled = bool(label.get("enabled"))
    label_font_pt = float(label.get("font_pt", 12))

    placements: list[CellPlacement] = []
    for r_idx, (group, cells) in enumerate(rows):
        target_row = start_row + r_idx * step_row
        if label_enabled and sn_col > 0:
            placements.append(CellPlacement(
                path=None, row=target_row, col=sn_col,
                span_cols=1, span_rows=cell_rows,
                text=group, font_pt=label_font_pt,
            ))
        for ci, f in enumerate(cells):
            if f is None:
                continue
            placements.append(CellPlacement(
                path=f, row=target_row,
                col=start_col + ci * cell_cols,
                span_cols=cell_cols, span_rows=cell_rows,
            ))

    img_fit = str(excel_cfg.get("imgFit", "contain"))
    embed_in_cell = bool(excel_cfg.get("embedInCell", False))
    contain_inset_pct = float(excel_cfg.get("containInset", 5))
    crop = manifest.get("crop") or None
    if crop and not crop.get("enabled"):
        crop = None
    return write_xlsx(
        placements, out_path, template_path,
        embed_in_cell=embed_in_cell,
        img_fit=img_fit,
        contain_inset=contain_inset_pct / 100.0,
        crop=crop,
    )
