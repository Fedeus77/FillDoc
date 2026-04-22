from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from filldoc.templates.models import TemplateCard
from filldoc.templates.quality import analyze_template_quality
from filldoc.variables.dictionary import default_dictionary


def _card(*variables: str) -> TemplateCard:
    return TemplateCard(
        name="Template",
        path="template.docx",
        category="",
        variables_in_order=list(variables),
        variables_unique=list(dict.fromkeys(variables)),
    )


def test_quality_detects_extra_spaces() -> None:
    issues = analyze_template_quality(_card(" Должник "), default_dictionary())

    assert any(issue.code == "extra_spaces" for issue in issues)


def test_quality_detects_variables_missing_from_dictionary() -> None:
    issues = analyze_template_quality(_card("CustomField"), default_dictionary())

    assert any(issue.code == "not_in_dictionary" for issue in issues)


def test_quality_detects_unknown_variables_against_project() -> None:
    issues = analyze_template_quality(
        _card("CustomField"),
        default_dictionary(),
        project_fields={"Должник": "ООО Ромашка"},
    )

    assert any(issue.code == "unknown_variable" for issue in issues)


def test_quality_does_not_mark_project_field_as_unknown() -> None:
    issues = analyze_template_quality(
        _card("CustomField"),
        default_dictionary(),
        project_fields={"customfield": "value"},
    )

    assert not any(issue.code == "unknown_variable" for issue in issues)


def test_quality_detects_case_only_differences() -> None:
    issues = analyze_template_quality(_card("Должник", "должник"), default_dictionary())

    assert any(issue.code == "case_only_difference" for issue in issues)
