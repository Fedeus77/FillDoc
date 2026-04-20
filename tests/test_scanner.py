from __future__ import annotations

import json
import sys
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.templates.scanner import TemplateLibrary


def _make_docx(path: Path, text: str = "Value {field}") -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    doc.add_paragraph(text)
    doc.save(str(path))
    return path


def test_scan_finds_templates_in_nested_folders(tmp_path: Path) -> None:
    _make_docx(tmp_path / "root.docx")
    _make_docx(tmp_path / "claims" / "court" / "nested.docx")

    cards = TemplateLibrary(str(tmp_path)).scan()

    assert [(card.category, card.name) for card in cards] == [
        ("", "root"),
        (str(Path("claims") / "court"), "nested"),
    ]


def test_scan_ignores_temporary_word_files(tmp_path: Path) -> None:
    _make_docx(tmp_path / "real.docx")
    (tmp_path / "~$real.docx").write_bytes(b"temporary lock file")

    cards = TemplateLibrary(str(tmp_path)).scan()

    assert [card.name for card in cards] == ["real"]


def test_scan_saves_template_card_to_filldoc_cache(tmp_path: Path) -> None:
    _make_docx(tmp_path / "folder" / "template.docx", "Hello {client}")

    [card] = TemplateLibrary(str(tmp_path)).scan()

    cache_path = tmp_path / ".filldoc" / "folder__template.docx.json"
    data = json.loads(cache_path.read_text(encoding="utf-8"))
    assert data["name"] == "template"
    assert data["path"] == card.path
    assert data["category"] == "folder"
    assert data["variables_unique"] == ["client"]


def test_scan_preserves_manual_output_name_rule(tmp_path: Path) -> None:
    _make_docx(tmp_path / "template.docx")
    library = TemplateLibrary(str(tmp_path))
    library.scan()
    cache_path = tmp_path / ".filldoc" / "template.docx.json"
    data = json.loads(cache_path.read_text(encoding="utf-8"))
    data["output_name_rule"] = "{%filename%} custom {client}"
    cache_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    [card] = library.scan()

    assert card.output_name_rule == "{%filename%} custom {client}"
    cached_again = json.loads(cache_path.read_text(encoding="utf-8"))
    assert cached_again["output_name_rule"] == "{%filename%} custom {client}"
