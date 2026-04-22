from __future__ import annotations

from pathlib import Path

import pytest

from filldoc.excel.models import Project
from filldoc.projects.json_import import (
    JsonImportError,
    merge_fields_into_project,
    project_from_json_fields,
    read_json_fields,
)


def test_read_json_fields_normalizes_keys_and_values(tmp_path: Path) -> None:
    path = tmp_path / "project.json"
    path.write_text('{"A": 1, "B": null, "3": true}', encoding="utf-8")

    assert read_json_fields(path) == {"A": "1", "B": "", "3": "True"}


def test_read_json_fields_rejects_non_object(tmp_path: Path) -> None:
    path = tmp_path / "project.json"
    path.write_text("[1, 2, 3]", encoding="utf-8")

    with pytest.raises(JsonImportError):
        read_json_fields(path)


def test_project_from_json_fields_uses_case_number_or_file_name(tmp_path: Path) -> None:
    path = tmp_path / "fallback-name.json"

    assert project_from_json_fields(path, {"Номер дела": "A-10"}).project_id == "A-10"
    assert project_from_json_fields(path, {"Field": "Value"}).project_id == "fallback-name"


def test_merge_fields_into_project_reports_counts_and_updates_headers() -> None:
    project = Project(
        project_id="p1",
        fields={"A": "old", "B": "", "C": "same", "D": "keep"},
        headers=["A"],
    )
    calls: list[tuple[str, str, str]] = []

    def should_replace(key: str, old_value: str, new_value: str) -> bool:
        calls.append((key, old_value, new_value))
        return key == "A"

    result = merge_fields_into_project(
        project,
        {"A": "new", "B": "fill", "C": "same", "D": "json", "E": "", "F": "value"},
        should_replace=should_replace,
    )

    assert result.added_count == 3
    assert result.replaced_count == 1
    assert result.kept_count == 2
    assert calls == [("A", "old", "new"), ("D", "keep", "json")]
    assert project.fields == {
        "A": "new",
        "B": "fill",
        "C": "same",
        "D": "keep",
        "E": "",
        "F": "value",
    }
    assert project.headers == ["A", "B", "C", "D", "E", "F"]
