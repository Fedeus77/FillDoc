from __future__ import annotations

import re


_bad_chars = re.compile(r'[<>:"/\\|?*\x00-\x1F]')
_ws = re.compile(r"\s+")
_field_token = re.compile(r"\{([^{}]+)\}")


def safe_filename(name: str) -> str:
    s = (name or "").strip()
    s = _bad_chars.sub(" ", s)
    s = _ws.sub(" ", s).strip()
    if not s:
        return "document"
    return s


def apply_output_name_rule(
    rule: str,
    filename_stem: str,
    fields: dict[str, str],
) -> str:
    """
    Applies a template output-name rule.

    Supported tokens:
    - {%filename%}: template file name without extension
    - {Field name}: value from the selected project fields
    """
    result = (rule or "").replace("{%filename%}", filename_stem)

    def replace_field(match: re.Match) -> str:
        key = match.group(1).strip()
        return (fields.get(key) or "").strip()

    result = _field_token.sub(replace_field, result)
    result = re.sub(r"(\s*-\s*){2,}", " - ", result)
    result = re.sub(r"^\s*-\s*|\s*-\s*$", "", result)
    return result.strip() or filename_stem


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

