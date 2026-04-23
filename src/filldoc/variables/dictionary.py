from __future__ import annotations

from dataclasses import dataclass, field
from importlib import resources
import json
from pathlib import Path
from typing import Any, Iterable

from .normalize import normalize_var_name


FIELD_TYPES = ("text", "multiline", "select", "date", "number", "boolean")


@dataclass(frozen=True)
class VariableEntry:
    technical_name: str
    display_name: str
    variants: set[str] = field(default_factory=set)
    field_type: str = "text"
    group: str = "project"
    required: bool = False
    comment: str = ""


class VariableDictionary:
    """
    Словарь переменных: сопоставляет техническое имя, видимое имя и варианты
    написания с одной записью VariableEntry.
    """

    def __init__(self, entries: Iterable[VariableEntry] | None = None) -> None:
        self._entries: dict[str, VariableEntry] = {}
        self._by_norm: dict[str, VariableEntry] = {}
        for entry in entries or ():
            self.add(entry)

    def add(self, entry: VariableEntry) -> None:
        key = normalize_var_name(entry.technical_name)
        if not key:
            raise ValueError("technical_name must not be empty")
        self._entries[key] = entry
        self._reindex()

    def resolve(self, raw_name: str) -> VariableEntry | None:
        return self._by_norm.get(normalize_var_name(raw_name))

    def all_entries(self) -> list[VariableEntry]:
        return sorted(self._entries.values(), key=lambda e: e.display_name.lower())

    def _reindex(self) -> None:
        self._by_norm.clear()
        for entry in self._entries.values():
            names = [entry.technical_name, entry.display_name, *entry.variants]
            for name in names:
                norm = normalize_var_name(name)
                if norm:
                    self._by_norm[norm] = entry


def variables_dir() -> Path:
    return Path.home() / ".filldoc"


def user_variables_path() -> Path:
    return variables_dir() / "variables.json"


def _default_variables_text() -> str:
    try:
        return (
            resources.files("filldoc.data")
            .joinpath("default_variables.json")
            .read_text(encoding="utf-8")
        )
    except Exception:
        path = Path(__file__).resolve().parents[1] / "data" / "default_variables.json"
        return path.read_text(encoding="utf-8")


def _raw_variables(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, dict):
        payload = payload.get("variables", [])
    if not isinstance(payload, list):
        raise ValueError("variables.json must contain a list or an object with a variables list")
    rows: list[dict[str, Any]] = []
    for item in payload:
        if isinstance(item, dict):
            rows.append(item)
    return rows


def _as_variants(value: Any) -> set[str]:
    if isinstance(value, str):
        return {value.strip()} if value.strip() else set()
    if not isinstance(value, (list, tuple, set)):
        return set()
    return {str(v).strip() for v in value if str(v).strip()}


def entry_from_dict(data: dict[str, Any]) -> VariableEntry:
    technical_name = str(data.get("technical_name", "")).strip()
    display_name = str(data.get("display_name", "")).strip() or technical_name
    field_type = str(data.get("field_type", "text")).strip()
    if field_type not in FIELD_TYPES:
        field_type = "text"
    return VariableEntry(
        technical_name=technical_name,
        display_name=display_name,
        variants=_as_variants(data.get("variants", [])),
        field_type=field_type,
        group=str(data.get("group", "project")).strip() or "project",
        required=bool(data.get("required", False)),
        comment=str(data.get("comment", "")).strip(),
    )


def entry_to_dict(entry: VariableEntry) -> dict[str, Any]:
    return {
        "technical_name": entry.technical_name,
        "display_name": entry.display_name,
        "variants": sorted(entry.variants, key=str.lower),
        "field_type": entry.field_type,
        "required": entry.required,
        "group": entry.group,
        "comment": entry.comment,
    }


def load_entries_from_file(path: str | Path) -> list[VariableEntry]:
    payload = json.loads(Path(path).read_text(encoding="utf-8"))
    entries: list[VariableEntry] = []
    for row in _raw_variables(payload):
        entry = entry_from_dict(row)
        if entry.technical_name:
            entries.append(entry)
    return entries


def save_entries_to_file(entries: Iterable[VariableEntry], path: str | Path) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "version": 1,
        "variables": [entry_to_dict(entry) for entry in entries],
    }
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def default_entries() -> list[VariableEntry]:
    payload = json.loads(_default_variables_text())
    return [entry_from_dict(row) for row in _raw_variables(payload)]


def ensure_user_variables_file() -> Path:
    path = user_variables_path()
    if not path.exists():
        save_entries_to_file(default_entries(), path)
    return path


def load_variable_entries(
    path: str | Path | None = None,
    *,
    create_user_file: bool = True,
) -> list[VariableEntry]:
    if path is None:
        try:
            target = ensure_user_variables_file() if create_user_file else user_variables_path()
        except OSError:
            return default_entries()
        if not target.exists():
            return default_entries()
        try:
            return load_entries_from_file(target)
        except Exception:
            return default_entries()

    return load_entries_from_file(path)


def load_variable_dictionary(
    path: str | Path | None = None,
    *,
    create_user_file: bool = True,
) -> VariableDictionary:
    return VariableDictionary(load_variable_entries(path, create_user_file=create_user_file))


def default_dictionary() -> VariableDictionary:
    return VariableDictionary(default_entries())
