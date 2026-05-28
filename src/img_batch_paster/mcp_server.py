"""MCP server exposing img-batch-paster as a tool for AI clients.

Run via the `img-batch-paster-mcp` entrypoint. Uses stdio transport, which is
the most compatible with Claude Code, Claude Desktop, AnythingLLM, etc.

Install with the mcp extra:
    pip install -e ".[mcp]"
"""
from __future__ import annotations

from pathlib import Path

try:
    from mcp.server.fastmcp import FastMCP
except ImportError as e:
    raise SystemExit(
        "mcp package not installed. Run: pip install -e \".[mcp]\""
    ) from e

from .paste_job import run_paste_job, run_paste_job_ibp


mcp = FastMCP("img-batch-paster")


@mcp.tool()
def paste_images_to_pptx(
    image_folder: str,
    output_path: str,
    template: str | None = None,
    pattern: str = "{group}_{n}",
    cols: int = 3,
    cell_w_cm: float = 6.0,
    cell_h_cm: float = 4.0,
    gap_x_cm: float = 0.3,
    gap_y_cm: float = 0.3,
) -> str:
    """Paste images from a folder into a .pptx grid layout.

    Filenames must follow `pattern` so they can be grouped into rows. For
    example pattern `"{group}_{n}"` matches `cat_1.png`, `cat_2.png`,
    `dog_1.png` — grouping `cat` and `dog` as two rows, columns by `n`.

    Args:
        image_folder: Absolute path to a folder of images.
        output_path: Absolute path where the .pptx will be written.
        template: Optional .pptx template to clone for each page.
        pattern: Filename pattern with placeholders {group} and {n}.
        cols: Number of columns in the grid (rows derived from filenames).
        cell_w_cm: Width of each image cell in cm.
        cell_h_cm: Hint for cell height in cm (actual height keeps aspect).
        gap_x_cm: Horizontal gap between cells in cm.
        gap_y_cm: Vertical gap between cells in cm.

    Returns:
        Absolute path string of the generated .pptx.
    """
    out = run_paste_job(
        image_folder=image_folder,
        output_path=output_path,
        template=template,
        pattern=pattern,
        cols=cols,
        cell_w_cm=cell_w_cm,
        cell_h_cm=cell_h_cm,
        gap_x_cm=gap_x_cm,
        gap_y_cm=gap_y_cm,
    )
    return str(out)


@mcp.tool()
def paste_with_ibp(
    ibp_path: str,
    image_folder: str,
    output_path: str,
) -> str:
    """Replay a saved .ibp config bundle against a new folder of images.

    The .ibp is a zip created from the web UI ("📁 配置 → 匯出 .ibp"), containing
    the recipe (grid, pattern, slide dims) and optionally a template file.

    Supported modes:
      - 依檔名 + 依順序 (autoAlign=False) — files chunked linearly into rows
      - 依檔名 idx 對位 (autoAlign=True) — idx in filename picks the column

    Supported outputs (decided by output_path extension):
      - .pptx
      - .key  (macOS only; conversion via Keynote.app)
      - .xlsx

    Not supported (raises a clear error):
      - 依範本 SN 多 source 模式 (snMatchMode=True) — needs different signature
        to pass multiple source folders + per-source crop. Use the web UI.

    Args:
        ibp_path: Absolute path to a .ibp file.
        image_folder: Absolute path to a folder of images to paste.
        output_path: Absolute path where the output (.pptx/.key/.xlsx) is written.

    Returns:
        Absolute path string of the generated file.
    """
    return str(run_paste_job_ibp(ibp_path, image_folder, output_path))


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
