from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class TemplateCard:
    name: str
    path: str
    category: str
    variables_in_order: list[str] = field(default_factory=list)
    variables_unique: list[str] = field(default_factory=list)
    output_name_rule: str = "{%filename%}"
    active: bool = True

