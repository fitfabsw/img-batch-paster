from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml


@dataclass
class Point:
    x_cm: float
    y_cm: float


@dataclass
class Size:
    w_cm: float
    h_cm: float


@dataclass
class GridConfig:
    origin: Point
    cell: Size
    gap: Point
    cols: int


@dataclass
class SlideConfig:
    width_cm: float
    height_cm: float


@dataclass
class InputConfig:
    folder: Path
    pattern: str
    extensions: list[str]


@dataclass
class OutputConfig:
    path: Path
    template: Path | None


@dataclass
class Config:
    slide: SlideConfig
    grid: GridConfig
    input: InputConfig
    output: OutputConfig


def load_config(path: str | Path) -> Config:
    raw = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    base = Path(path).resolve().parent

    s = raw["slide"]
    g = raw["grid"]
    i = raw["input"]
    o = raw["output"]

    def _resolve(p: str) -> Path:
        pp = Path(p)
        return pp if pp.is_absolute() else (base / pp).resolve()

    return Config(
        slide=SlideConfig(width_cm=float(s["width_cm"]), height_cm=float(s["height_cm"])),
        grid=GridConfig(
            origin=Point(**g["origin"]),
            cell=Size(**g["cell"]),
            gap=Point(**g["gap"]),
            cols=int(g["cols"]),
        ),
        input=InputConfig(
            folder=_resolve(i["folder"]),
            pattern=i["pattern"],
            extensions=[e.lower() for e in i["extensions"]],
        ),
        output=OutputConfig(
            path=_resolve(o["path"]),
            template=_resolve(o["template"]) if o.get("template") else None,
        ),
    )
