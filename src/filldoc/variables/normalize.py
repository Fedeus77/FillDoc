from __future__ import annotations

import re


_ws_re = re.compile(r"\s+")


def normalize_var_name(raw: str) -> str:
    """
    Нормализация MVP:
    - убрать фигурные скобки
    - привести пробелы
    - привести к нижнему регистру
    - убрать точки в сокращениях ("Юр. адрес" -> "Юр адрес")
    """
    s = raw.strip()
    if s.startswith("{") and s.endswith("}"):
        s = s[1:-1].strip()
    s = s.replace(".", " ")
    s = _ws_re.sub(" ", s).strip().lower()
    return s

