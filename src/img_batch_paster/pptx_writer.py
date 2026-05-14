from __future__ import annotations

import copy
from dataclasses import dataclass
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Cm

from .config import Config
from .grouper import GroupedImages


_R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _remap_rids_in_tree(elem, rid_map: dict[str, str]) -> None:
    """Walk an XML subtree and rewrite r:* attributes whose value appears in rid_map."""
    for el in elem.iter():
        for attr, val in list(el.attrib.items()):
            if attr.startswith(_R_NS) and val in rid_map:
                el.attrib[attr] = rid_map[val]


def _duplicate_slide(prs, source_slide):
    """Duplicate an existing slide: copy shapes and their relationships, remapping rIds.

    python-pptx assigns its own rIds when we add relationships to the new slide, but the
    deep-copied shape XML still references the source's rIds. We build a src→dst rId map
    and rewrite the copied XML so all references resolve. Otherwise PowerPoint complains
    about needing repair (dangling r:embed / r:link references).
    """
    new_slide = prs.slides.add_slide(source_slide.slide_layout)
    # Drop default placeholders coming from the layout
    for shp in list(new_slide.shapes):
        sp = shp._element
        sp.getparent().remove(sp)

    src_rels = source_slide.part.rels
    dst_rels = new_slide.part.rels

    # Build rId map by ensuring each source relationship exists in the new slide
    rid_map: dict[str, str] = {}
    for rel in src_rels.values():
        if "notesSlide" in rel.reltype:
            continue
        if rel.is_external:
            new_rid = dst_rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
        else:
            new_rid = dst_rels.get_or_add(rel.reltype, rel.target_part)
        if new_rid != rel.rId:
            rid_map[rel.rId] = new_rid

    # Deep-copy shapes and remap their rId references
    for shp in source_slide.shapes:
        new_el = copy.deepcopy(shp._element)
        if rid_map:
            _remap_rids_in_tree(new_el, rid_map)
        new_slide.shapes._spTree.insert_element_before(new_el, "p:extLst")

    return new_slide


@dataclass
class Placement:
    path: Path
    x_cm: float
    y_cm: float
    w_cm: float
    h_cm: float


def _add_placements_to_slide(slide, placements: list[Placement]) -> None:
    for pl in placements:
        slide.shapes.add_picture(
            str(pl.path), Cm(pl.x_cm), Cm(pl.y_cm),
            width=Cm(pl.w_cm), height=Cm(pl.h_cm),
        )


def write_pages(
    slide_w_cm: float,
    slide_h_cm: float,
    pages: list[list[Placement]],
    out_path: Path,
    template: Path | None = None,
) -> Path:
    """Write a multi-page pptx. Each page = a list of Placements.

    With a template: every page is a deep-copy of the template's first slide.
    Without a template: every page is a fresh blank slide.
    """
    if not pages:
        raise ValueError("pages 不可為空")

    if template:
        prs = Presentation(str(template))
        if len(prs.slides) == 0:
            blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
            base = prs.slides.add_slide(blank)
        else:
            base = prs.slides[0]
        # IMPORTANT: duplicate FIRST (while base is still pristine), THEN add images.
        # If we mutated base first, the duplicates would inherit those mutations.
        page_slides = [base]
        for _ in range(len(pages) - 1):
            page_slides.append(_duplicate_slide(prs, base))
        for slide, page_pls in zip(page_slides, pages):
            _add_placements_to_slide(slide, page_pls)
    else:
        prs = Presentation()
        prs.slide_width = Cm(slide_w_cm)
        prs.slide_height = Cm(slide_h_cm)
        blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
        for page_pls in pages:
            slide = prs.slides.add_slide(blank)
            _add_placements_to_slide(slide, page_pls)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def write_placements(
    slide_w_cm: float,
    slide_h_cm: float,
    placements: list[Placement],
    out_path: Path,
    template: Path | None = None,
) -> Path:
    """Single-page convenience wrapper around write_pages."""
    return write_pages(slide_w_cm, slide_h_cm, [placements], out_path, template)


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
