from __future__ import annotations

from dataclasses import dataclass
import re

from filldoc.templates.models import TemplateCard
from filldoc.variables.dictionary import VariableDictionary
from filldoc.variables.normalize import normalize_var_name


_inner_ws = re.compile(r"\s{2,}")


@dataclass(frozen=True)
class TemplateQualityIssue:
    code: str
    variable: str
    message: str


def _project_has_field(raw: str, project_fields: dict[str, str] | None) -> bool:
    if not project_fields:
        return False
    target = normalize_var_name(raw)
    return any(normalize_var_name(key) == target for key in project_fields)


def analyze_template_quality(
    card: TemplateCard,
    dictionary: VariableDictionary,
    project_fields: dict[str, str] | None = None,
) -> list[TemplateQualityIssue]:
    issues: list[TemplateQualityIssue] = []
    variables = card.variables_in_order or card.variables_unique

    for raw in variables:
        stripped = raw.strip()
        if not stripped:
            issues.append(
                TemplateQualityIssue(
                    code="unknown_variable",
                    variable=raw,
                    message="Пустое имя переменной.",
                )
            )
            continue

        if raw != stripped or _inner_ws.search(stripped):
            issues.append(
                TemplateQualityIssue(
                    code="extra_spaces",
                    variable=raw,
                    message=f"Лишние пробелы: {{{raw}}}. Лучше использовать {{{stripped}}}.",
                )
            )

        if dictionary.resolve(stripped) is None:
            issues.append(
                TemplateQualityIssue(
                    code="not_in_dictionary",
                    variable=raw,
                    message=f"Переменная {{{stripped}}} не найдена в словаре.",
                )
            )
            if project_fields is not None and not _project_has_field(stripped, project_fields):
                issues.append(
                    TemplateQualityIssue(
                        code="unknown_variable",
                        variable=raw,
                        message=f"Переменная {{{stripped}}} не найдена ни в словаре, ни в проекте.",
                    )
                )

    by_lower: dict[str, set[str]] = {}
    for raw in variables:
        stripped = raw.strip()
        if stripped:
            by_lower.setdefault(stripped.casefold(), set()).add(stripped)

    for variants in by_lower.values():
        if len(variants) <= 1:
            continue
        ordered = sorted(variants)
        issues.append(
            TemplateQualityIssue(
                code="case_only_difference",
                variable=", ".join(ordered),
                message="Переменные отличаются только регистром: " + ", ".join(f"{{{v}}}" for v in ordered) + ".",
            )
        )

    seen: set[tuple[str, str, str]] = set()
    unique_issues: list[TemplateQualityIssue] = []
    for issue in issues:
        key = (issue.code, issue.variable, issue.message)
        if key in seen:
            continue
        seen.add(key)
        unique_issues.append(issue)
    return unique_issues
