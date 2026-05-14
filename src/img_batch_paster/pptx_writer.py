from __future__ import annotations

from pathlib import Path

from pptx import Presentation
from pptx.util import Cm

from .config import Config
from .grouper import GroupedImages


def write_pptx(cfg: Config, grouped: GroupedImages) -> Path:
    prs = Presentation(str(cfg.output.template)) if cfg.output.template else Presentation()
    prs.slide_width = Cm(cfg.slide.width_cm)
    prs.slide_height = Cm(cfg.slide.height_cm)

    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank_layout)

    g = cfg.grid
    for row_idx, (_group, paths) in enumerate(grouped.rows):
        for col_idx, img_path in enumerate(paths):
            if img_path is None:
                continue
            x = g.origin.x_cm + col_idx * (g.cell.w_cm + g.gap.x_cm)
            y = g.origin.y_cm + row_idx * (g.cell.h_cm + g.gap.y_cm)
            slide.shapes.add_picture(
                str(img_path),
                Cm(x),
                Cm(y),
                width=Cm(g.cell.w_cm),
                height=Cm(g.cell.h_cm),
            )

    cfg.output.path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(cfg.output.path))
    return cfg.output.path
