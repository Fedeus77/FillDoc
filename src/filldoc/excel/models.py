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
    # Порядок колонок из Excel (как в заголовке), чтобы можно было
    # отображать поля в том же порядке, что и в таблице.
    headers: list[str] | None = None
    # Номер строки в Excel-листе (1-based), чтобы можно было
    # сохранять изменения без привязки к колонке "№ дела".
    row_index: int | None = None

