"""Stateless entry: image folder + grid params → .pptx file.

Designed for headless callers (MCP server, CLI scripts, CI). Wraps the
existing scan_folder + write_pptx pipeline so callers don't need to build
a YAML Config first.
"""
from __future__ import annotations

from pathlib import Path

from .config import Config, GridConfig, InputConfig, OutputConfig, Point, Size, SlideConfig
from .grouper import scan_folder
from .pptx_writer import write_pptx


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
