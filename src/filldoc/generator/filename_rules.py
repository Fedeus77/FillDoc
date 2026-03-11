from __future__ import annotations

import re


_bad_chars = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
_ws = re.compile(r"\s+")


def safe_filename(name: str) -> str:
    s = (name or "").strip()
    s = _bad_chars.sub(" ", s)
    s = _ws.sub(" ", s).strip()
    if not s:
        return "document"
    return s


def ensure_unique_path(path):
    from pathlib import Path

    p = Path(path)
    if not p.exists():
        return p
    stem = p.stem
    suf = p.suffix
    parent = p.parent
    i = 2
    while True:
        cand = parent / f"{stem} ({i}){suf}"
        if not cand.exists():
            return cand
        i += 1

