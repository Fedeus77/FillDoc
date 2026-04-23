"""Тесты словаря переменных и нормализации."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.variables.dictionary import (
    VariableEntry,
    default_dictionary,
    load_entries_from_file,
    load_variable_dictionary,
    save_entries_to_file,
)
from filldoc.variables.normalize import normalize_var_name


class TestNormalize:
    def test_lowercases(self) -> None:
        assert normalize_var_name("ДОЛЖНИК") == normalize_var_name("должник")

    def test_strips_spaces(self) -> None:
        assert normalize_var_name("  Должник  ") == normalize_var_name("Должник")


class TestVariableDictionary:
    def test_resolve_by_variant(self) -> None:
        d = default_dictionary()
        entry = d.resolve("ИНН ДОЛЖНИКА")
        assert entry is not None
        assert entry.technical_name == "ИНН должника"

    def test_resolve_unknown(self) -> None:
        d = default_dictionary()
        assert d.resolve("НесуществующееПоле") is None

    def test_all_entries_unique(self) -> None:
        d = default_dictionary()
        entries = d.all_entries()
        names = [e.technical_name for e in entries]
        assert len(names) == len(set(names))

    def test_save_and_load_entries_from_json(self, tmp_path: Path) -> None:
        path = tmp_path / "variables.json"
        save_entries_to_file(
            [
                VariableEntry(
                    technical_name="inn_debtor",
                    display_name="ИНН должника",
                    variants={"ИНН ДОЛЖНИКА"},
                    field_type="text",
                    required=True,
                    group="debtor",
                    comment="ИНН юрлица или ИП",
                )
            ],
            path,
        )

        entries = load_entries_from_file(path)

        assert entries == [
            VariableEntry(
                technical_name="inn_debtor",
                display_name="ИНН должника",
                variants={"ИНН ДОЛЖНИКА"},
                field_type="text",
                required=True,
                group="debtor",
                comment="ИНН юрлица или ИП",
            )
        ]

    def test_load_dictionary_from_json_resolves_alias(self, tmp_path: Path) -> None:
        path = tmp_path / "variables.json"
        path.write_text(
            """
{
  "version": 1,
  "variables": [
    {
      "technical_name": "inn_debtor",
      "display_name": "ИНН должника",
      "variants": ["ИНН ДОЛЖНИКА", "ИНН Должник"],
      "field_type": "text",
      "required": true,
      "group": "debtor",
      "comment": "ИНН юрлица или ИП"
    }
  ]
}
""",
            encoding="utf-8",
        )

        entry = load_variable_dictionary(path).resolve("ИНН Должник")

        assert entry is not None
        assert entry.technical_name == "inn_debtor"
        assert entry.required is True

    def test_invalid_field_type_falls_back_to_text(self, tmp_path: Path) -> None:
        path = tmp_path / "variables.json"
        path.write_text(
            """
[
  {
    "technical_name": "custom",
    "display_name": "Custom",
    "variants": [],
    "field_type": "unexpected"
  }
]
""",
            encoding="utf-8",
        )

        entries = load_entries_from_file(path)

        assert entries[0].field_type == "text"
