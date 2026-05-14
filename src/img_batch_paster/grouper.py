from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path


@dataclass
class GroupedImages:
    """rows: list of (group_name, [path_col1, path_col2, ...]) preserving first-seen order."""
    rows: list[tuple[str, list[Path | None]]]
    cols: int


def _pattern_to_regex(pattern: str) -> re.Pattern[str]:
    # Pattern uses placeholders {group} and {n}. Escape the literal parts.
    parts = re.split(r"(\{group\}|\{n\})", pattern)
    out = []
    for p in parts:
        if p == "{group}":
            out.append(r"(?P<group>.+?)")
        elif p == "{n}":
            out.append(r"(?P<n>\d+)")
        else:
            out.append(re.escape(p))
    return re.compile("^" + "".join(out) + "$")


def scan_folder(folder: Path, pattern: str, extensions: list[str], cols: int) -> GroupedImages:
    regex = _pattern_to_regex(pattern)
    exts = {e.lower() for e in extensions}

    # group_name -> {col_idx (0-based) -> path}
    buckets: dict[str, dict[int, Path]] = {}
    order: list[str] = []

    for entry in sorted(folder.iterdir()):
        if not entry.is_file() or entry.suffix.lower() not in exts:
            continue
        m = regex.match(entry.stem)
        if not m:
            continue
        group = m.group("group")
        n = int(m.group("n"))
        if n < 1 or n > cols:
            continue
        if group not in buckets:
            buckets[group] = {}
            order.append(group)
        buckets[group][n - 1] = entry

    rows: list[tuple[str, list[Path | None]]] = []
    for g in order:
        row: list[Path | None] = [None] * cols
        for idx, p in buckets[g].items():
            row[idx] = p
        rows.append((g, row))

    return GroupedImages(rows=rows, cols=cols)
