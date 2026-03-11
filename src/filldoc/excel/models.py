from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Project:
    """
    MVP: проект = строка Excel + набор полей (ключ -> значение).
    Идентификатором считаем значение в колонке '№ дела' (если есть),
    иначе используем номер строки.
    """

    project_id: str
    fields: dict[str, str] = field(default_factory=dict)

