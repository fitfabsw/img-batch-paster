from __future__ import annotations

import sys
from pathlib import Path

import click

from .config import load_config
from .grouper import scan_folder
from .pptx_writer import write_pptx


@click.command()
@click.option("-c", "--config", "config_path", required=True, type=click.Path(exists=True, dir_okay=False),
              help="YAML 設定檔路徑")
@click.option("--folder", type=click.Path(exists=True, file_okay=False), default=None,
              help="覆蓋 config 內的 input.folder")
@click.option("-o", "--output", type=click.Path(dir_okay=False), default=None,
              help="覆蓋 config 內的 output.path")
def main(config_path: str, folder: str | None, output: str | None) -> None:
    """把資料夾內的圖片依檔名規則貼到簡報的格子中。"""
    cfg = load_config(config_path)
    if folder:
        cfg.input.folder = Path(folder).resolve()
    if output:
        cfg.output.path = Path(output).resolve()

    grouped = scan_folder(cfg.input.folder, cfg.input.pattern, cfg.input.extensions, cfg.grid.cols)
    if not grouped.rows:
        click.echo(f"找不到符合 pattern '{cfg.input.pattern}' 的圖片於 {cfg.input.folder}", err=True)
        sys.exit(1)

    click.echo(f"偵測到 {len(grouped.rows)} 列 x {grouped.cols} 欄：")
    for name, row in grouped.rows:
        filled = sum(1 for p in row if p is not None)
        click.echo(f"  - {name}: {filled}/{grouped.cols}")

    out = write_pptx(cfg, grouped)
    click.echo(f"已輸出: {out}")


if __name__ == "__main__":
    main()
