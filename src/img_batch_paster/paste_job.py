"""Stateless entry: image folder + grid params → .pptx file.

Designed for headless callers (MCP server, CLI scripts, CI). Wraps the
existing scan_folder + write_pptx pipeline so callers don't need to build
a YAML Config first.
"""
from __future__ import annotations

import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from .config import Config, GridConfig, InputConfig, OutputConfig, Point, Size, SlideConfig
from .grouper import GroupedImages, scan_folder
from .pptx_writer import write_pptx


class IbpModeUnsupported(ValueError):
    """Raised when the .ibp config uses a mode this entry point can't replay."""


def _ui_pattern_to_regex(pattern: str) -> re.Pattern[str]:
    """Compile a web UI-style pattern with {group}/{idx}/{skip} placeholders.

    Mirrors the frontend's extractGroup logic so .ibp files exported from the
    UI work here. {idx} and {skip} are captured-but-ignored — only {group}
    matters for bucketing.
    """
    parts = re.split(r"(\{group\}|\{idx\}|\{skip\}|\{n\})", pattern)
    out: list[str] = []
    for p in parts:
        if p == "{group}":
            out.append(r"(?P<group>.+?)")
        elif p in ("{idx}", "{skip}", "{n}"):
            out.append(r".+?")
        else:
            out.append(re.escape(p))
    return re.compile("^" + "".join(out) + "$")


def _scan_sequential(folder: Path, pattern: str, extensions: list[str], cols: int) -> GroupedImages:
    """Group by {group}, fill cols left-to-right by alphabetic filename order.

    For the web UI's "依檔名 + 依順序" mode: cell position is encounter order
    within the group, not a number parsed from the filename.
    """
    regex = _ui_pattern_to_regex(pattern)
    exts = {e.lower() for e in extensions}

    buckets: dict[str, list[Path]] = {}
    order: list[str] = []
    for entry in sorted(folder.iterdir()):
        if not entry.is_file() or entry.suffix.lower() not in exts:
            continue
        m = regex.match(entry.stem)
        if not m:
            continue
        try:
            group = m.group("group")
        except IndexError:
            group = entry.stem
        if group not in buckets:
            buckets[group] = []
            order.append(group)
        buckets[group].append(entry)

    rows: list[tuple[str, list[Path | None]]] = []
    for g in order:
        files = buckets[g][:cols]
        row: list[Path | None] = list(files) + [None] * (cols - len(files))
        rows.append((g, row))
    return GroupedImages(rows=rows, cols=cols)


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
        raise FileNotFoundError(
            f"No images matching pattern '{pattern}' in {folder}"
        )
    return write_pptx(cfg, grouped)


def run_paste_job_ibp(
    ibp_path: str | Path,
    image_folder: str | Path,
    output_path: str | Path,
) -> Path:
    """Replay an .ibp config bundle against a new image folder → pptx.

    Only supports the "依檔名 / 依順序" mode (autoAlign=false, snMatchMode=false)
    with .pptx output. Other modes (依檔名 idx, 依範本 SN, .xlsx, .key) are
    intentionally out of scope — use the web UI for those.
    """
    ibp = Path(ibp_path).expanduser().resolve()
    if not ibp.is_file():
        raise FileNotFoundError(f".ibp not found: {ibp}")
    out = Path(output_path).expanduser().resolve()
    if out.suffix.lower() != ".pptx":
        raise IbpModeUnsupported(
            f"Only .pptx output is supported here, got '{out.suffix}'. "
            "Use the web UI for .xlsx or .key."
        )

    with zipfile.ZipFile(ibp, "r") as zf:
        try:
            manifest = json.loads(zf.read("manifest.json").decode("utf-8"))
        except KeyError as e:
            raise ValueError(f"{ibp.name} missing manifest.json") from e

        mode = manifest.get("mode", {})
        if mode.get("autoAlign") or mode.get("snMatchMode"):
            raise IbpModeUnsupported(
                "Only basic mode (依檔名 + 依順序) is supported via MCP. "
                "Config uses autoAlign=%s, snMatchMode=%s — please run via web UI."
                % (mode.get("autoAlign"), mode.get("snMatchMode"))
            )

        tmp_dir = Path(tempfile.mkdtemp(prefix="ibp-replay-"))
        template_path: Path | None = None
        for name in zf.namelist():
            if name.startswith("template."):
                template_path = tmp_dir / name
                with zf.open(name) as src, open(template_path, "wb") as dst:
                    dst.write(src.read())
                break

    try:
        grid = manifest.get("grid") or {}
        slide = manifest.get("slide") or {}
        label = manifest.get("label") or {}
        # Frontend stores `grid.origin.x` (no _cm suffix) and `grid.width` (no cell.w_cm)
        origin = grid.get("origin") or {}
        gap = grid.get("gap") or {}
        cols = int(grid.get("cols", 3))
        pattern = label.get("pattern") or "{group}_{n}"
        extensions = [".png", ".jpg", ".jpeg"]

        folder = Path(image_folder).expanduser().resolve()
        grouped = _scan_sequential(folder, pattern, extensions, cols)
        if not grouped.rows:
            raise FileNotFoundError(
                f"No images matching pattern '{pattern}' in {folder}"
            )

        cfg = Config(
            slide=SlideConfig(
                width_cm=float(slide.get("width_cm", 25.4)),
                height_cm=float(slide.get("height_cm", 14.29)),
            ),
            grid=GridConfig(
                origin=Point(
                    x_cm=float(origin.get("x", origin.get("x_cm", 2.0))),
                    y_cm=float(origin.get("y", origin.get("y_cm", 2.0))),
                ),
                cell=Size(
                    w_cm=float(grid.get("width", 6.0)),
                    h_cm=0.0,  # height computed by aspect in placements_from_config
                ),
                gap=Point(
                    x_cm=float(gap.get("x", gap.get("x_cm", 0.3))),
                    y_cm=float(gap.get("y", gap.get("y_cm", 0.3))),
                ),
                cols=cols,
            ),
            input=InputConfig(folder=folder, pattern=pattern, extensions=extensions),
            output=OutputConfig(path=out, template=template_path),
        )
        return write_pptx(cfg, grouped)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
