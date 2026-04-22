"""Тесты извлечения переменных из .docx шаблонов."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.templates.vars_extractor import extract_docx_variables


def _make_docx(tmp_path: Path, text: str) -> Path:
    """Создаёт минимальный .docx с заданным текстом в первом параграфе."""
    from docx import Document
    doc = Document()
    doc.add_paragraph(text)
    p = tmp_path / "template.docx"
    doc.save(str(p))
    return p


class TestExtractDocxVariables:
    def test_extracts_single_variable(self, tmp_path: Path) -> None:
        p = _make_docx(tmp_path, "Должник: {Должник}")
        _, uniq = extract_docx_variables(str(p))
        assert "Должник" in uniq

    def test_deduplicates_variables(self, tmp_path: Path) -> None:
        p = _make_docx(tmp_path, "{Должник} и ещё раз {Должник}")
        _, uniq = extract_docx_variables(str(p))
        assert uniq.count("Должник") == 1

    def test_preserves_order(self, tmp_path: Path) -> None:
        p = _make_docx(tmp_path, "{A} {B} {A} {C}")
        ordered, uniq = extract_docx_variables(str(p))
        assert ordered == ["A", "B", "A", "C"]
        assert uniq == ["A", "B", "C"]

    def test_empty_document(self, tmp_path: Path) -> None:
        p = _make_docx(tmp_path, "Без переменных")
        _, uniq = extract_docx_variables(str(p))
        assert uniq == []

    def test_preserves_inner_spaces_for_quality_checks(self, tmp_path: Path) -> None:
        p = _make_docx(tmp_path, "Должник: { Должник }")
        ordered, uniq = extract_docx_variables(str(p))
        assert ordered == [" Должник "]
        assert uniq == [" Должник "]
