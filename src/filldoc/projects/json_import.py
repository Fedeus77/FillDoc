from __future__ import annotations

from collections.abc import Callable, Mapping
from dataclasses import dataclass
import json
from pathlib import Path

from filldoc.excel.models import Project


class JsonImportError(ValueError):
    """Raised when a JSON file cannot be imported as project fields."""


@dataclass(frozen=True)
class JsonMergeResult:
    added_count: int = 0
    replaced_count: int = 0
    kept_count: int = 0


ReplacementResolver = Callable[[str, str, str], bool]


def read_json_fields(path: str | Path) -> dict[str, str]:
    """Read a JSON object and normalize keys/values to strings."""
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:  # noqa: BLE001
        raise JsonImportError(f"Не удалось прочитать файл:\n{e}") from e

    if not isinstance(data, dict):
        raise JsonImportError("Файл должен содержать JSON-объект (словарь полей).")

    return {str(k): "" if v is None else str(v) for k, v in data.items()}


def project_from_json_fields(path: str | Path, fields: Mapping[str, str]) -> Project:
    """Build a new Project from imported fields."""
    path = Path(path)
    case_number = str(fields.get("№ дела", "") or fields.get("Номер дела", "")).strip()
    project_id = case_number or path.stem
    return Project(project_id=project_id, fields=dict(fields), headers=list(fields.keys()))


def merge_fields_into_project(
    project: Project,
    fields: Mapping[str, str],
    *,
    should_replace: ReplacementResolver | None = None,
) -> JsonMergeResult:
    """Merge imported fields into a project without UI dependencies."""
    added_count = 0
    replaced_count = 0
    kept_count = 0

    for key, new_value_raw in fields.items():
        new_value = new_value_raw.strip()
        old_value = str(project.fields.get(key, "") or "").strip()

        if key not in project.fields or old_value == "":
            project.fields[key] = new_value_raw
            added_count += 1
            continue

        if new_value == "" or old_value == new_value:
            kept_count += 1
            continue

        replace = should_replace(key, old_value, new_value) if should_replace is not None else False
        if replace:
            project.fields[key] = new_value_raw
            replaced_count += 1
        else:
            kept_count += 1

    headers = list(project.headers or [])
    for key in fields:
        if key and key not in headers:
            headers.append(key)
    if headers:
        project.headers = headers

    return JsonMergeResult(
        added_count=added_count,
        replaced_count=replaced_count,
        kept_count=kept_count,
    )

