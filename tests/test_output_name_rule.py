"""Тесты правила нейминга выходного файла."""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.generator.filename_rules import apply_output_name_rule


class TestApplyOutputNameRule:
    def test_filename_token(self) -> None:
        result = apply_output_name_rule("{%filename%}", "Заявление", {})
        assert result == "Заявление"

    def test_debtor_token(self) -> None:
        result = apply_output_name_rule(
            "{%filename%} - {ДОЛЖНИК}",
            "Заявление",
            {"ДОЛЖНИК": "ООО Ромашка"},
        )
        assert result == "Заявление - ООО Ромашка"

    def test_empty_field_collapses_separator(self) -> None:
        result = apply_output_name_rule(
            "{%filename%} - {ДОЛЖНИК}",
            "Заявление",
            {"ДОЛЖНИК": ""},
        )
        # Пустой ДОЛЖНИК → разделитель схлопывается
        assert result == "Заявление"

    def test_fallback_when_rule_is_empty(self) -> None:
        result = apply_output_name_rule("", "Шаблон", {})
        assert result == "Шаблон"

    def test_custom_field(self) -> None:
        result = apply_output_name_rule(
            "{%filename%} - {Заказчик} - {ДОЛЖНИК}",
            "Акт",
            {"Заказчик": "Иванов", "ДОЛЖНИК": "ООО Рога"},
        )
        assert result == "Акт - Иванов - ООО Рога"

    def test_field_tokens_tolerate_inner_spaces(self) -> None:
        result = apply_output_name_rule(
            "{%filename%} - { ДОЛЖНИК }",
            "Акт",
            {"ДОЛЖНИК": "ООО Рога"},
        )
        assert result == "Акт - ООО Рога"
