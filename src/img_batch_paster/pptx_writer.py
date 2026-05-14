from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Cm

from .config import Config
from .grouper import GroupedImages


@dataclass
class Placement:
    path: Path
    x_cm: float
    y_cm: float
    w_cm: float
    h_cm: float


def write_placements(
    slide_w_cm: float,
    slide_h_cm: float,
    placements: list[Placement],
    out_path: Path,
    template: Path | None = None,
) -> Path:
    if template:
        prs = Presentation(str(template))
        # 不要新增頁，直接在範本既有的第一張上加圖
        if len(prs.slides) == 0:
            blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
            slide = prs.slides.add_slide(blank_layout)
        else:
            slide = prs.slides[0]
    else:
        prs = Presentation()
        prs.slide_width = Cm(slide_w_cm)
        prs.slide_height = Cm(slide_h_cm)
        blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
        slide = prs.slides.add_slide(blank_layout)

    for pl in placements:
        slide.shapes.add_picture(
            str(pl.path),
            Cm(pl.x_cm),
            Cm(pl.y_cm),
            width=Cm(pl.w_cm),
            height=Cm(pl.h_cm),
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def _image_aspect(path: Path) -> float:
    """height / width"""
    with Image.open(path) as im:
        w, h = im.size
    return h / w if w else 1.0


def placements_from_config(cfg: Config, grouped: GroupedImages) -> list[Placement]:
    """Compute placements: each image keeps its natural aspect.

    cfg.grid.cell.w_cm = each image width
    cfg.grid.origin = first image top-left
    cfg.grid.gap = horizontal/vertical spacing between images
    Row height for row r = cell.w_cm * aspect(first non-null image of row r)
    """
    g = cfg.grid
    placements: list[Placement] = []
    cur_y = g.origin.y_cm

    for _group, row in grouped.rows:
        anchor = next((p for p in row if p is not None), None)
        if anchor is None:
            continue
        row_h = g.cell.w_cm * _image_aspect(anchor)
        for ci, p in enumerate(row):
            if p is None:
                continue
            x = g.origin.x_cm + ci * (g.cell.w_cm + g.gap.x_cm)
            placements.append(Placement(
                path=p, x_cm=x, y_cm=cur_y, w_cm=g.cell.w_cm, h_cm=row_h,
            ))
        cur_y += row_h + g.gap.y_cm
    return placements


def write_pptx(cfg: Config, grouped: GroupedImages) -> Path:
    """Backwards-compatible entry: computes placements from cfg, then writes."""
    placements = placements_from_config(cfg, grouped)
    return write_placements(
        cfg.slide.width_cm, cfg.slide.height_cm, placements, cfg.output.path, cfg.output.template,
    )
