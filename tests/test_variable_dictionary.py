"""Тесты словаря переменных и нормализации."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.variables.dictionary import default_dictionary
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
