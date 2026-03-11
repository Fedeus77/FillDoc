from __future__ import annotations

from pathlib import Path
import re

from docx import Document

from filldoc.core.errors import GenerationError
from filldoc.generator.filename_rules import safe_filename, ensure_unique_path


def _replace_in_paragraph(paragraph, mapping: dict[str, str]) -> None:
    """
    Аккуратная замена по runs, чтобы сохранить форматирование.

    Предполагаем, что:
    - в шаблоне переменные имеют вид {Имя};
    - в mapping ключи — это внутреннее имя без фигурных скобок.
    """
    if not mapping or not paragraph.runs:
        return

    # Строим один общий regex по всем ключам: \{(КЛЮЧ1|КЛЮЧ2|...)\}
    keys = [k for k in mapping.keys() if k]
    if not keys:
        return

    pattern = re.compile(
        "(" + "|".join(r"\{" + re.escape(k) + r"\}" for k in keys) + ")"
    )

    while True:
        runs = list(paragraph.runs)
        full_text = "".join(r.text or "" for r in runs)
        if not full_text:
            return

        m = pattern.search(full_text)
        if not m:
            return

        start, end = m.span()
        placeholder = m.group(0)  # например "{Заказчик}"
        inner = placeholder[1:-1]
        replacement = mapping.get(inner, "")

        # Построим карту позиций runs в общем тексте
        positions = []
        pos = 0
        for r in runs:
            t = r.text or ""
            positions.append((r, pos, pos + len(t)))
            pos += len(t)

        overlapping = [
            (r, s, e)
            for (r, s, e) in positions
            if not (e <= start or s >= end)
        ]
        if not overlapping:
            # На всякий случай, чтобы не уйти в бесконечный цикл
            return

        first_run, first_s, first_e = overlapping[0]
        last_run, last_s, last_e = overlapping[-1]

        for r, s, e in overlapping:
            text = r.text or ""
            if r is first_run and r is last_run:
                # Вся переменная внутри одного run
                local_start = max(start - s, 0)
                local_end = min(end - s, len(text))
                before = text[:local_start]
                after = text[local_end:]
                r.text = before + replacement + after
            elif r is first_run:
                local_start = max(start - s, 0)
                before = text[:local_start]
                r.text = before + replacement
            elif r is last_run:
                local_end = min(end - s, len(text))
                after = text[local_end:]
                r.text = after
            else:
                # промежуточные куски переменной очищаем
                r.text = ""


def _replace_in_cell(cell, mapping: dict[str, str]) -> None:
    for p in cell.paragraphs:
        _replace_in_paragraph(p, mapping)


def generate_docx_from_template(
    template_path: str,
    output_dir: str,
    output_name: str,
    mapping: dict[str, str],
) -> str:
    try:
        doc = Document(template_path)
    except Exception as e:  # noqa: BLE001
        raise GenerationError(f"Не удалось открыть шаблон для генерации: {e}") from e

    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                _replace_in_cell(cell, mapping)

    out_dir = Path(output_dir)
    if not out_dir.exists():
        raise GenerationError("Папка выгрузки недоступна или не существует.")

    filename = safe_filename(output_name) + ".docx"
    out_path = ensure_unique_path(out_dir / filename)

    try:
        doc.save(str(out_path))
        return str(out_path)
    except Exception as e:  # noqa: BLE001
        raise GenerationError(f"Не удалось сохранить готовый документ: {e}") from e

