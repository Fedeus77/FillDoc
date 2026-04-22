from __future__ import annotations

from dataclasses import dataclass, field
import uuid


FILLDOC_ID_FIELD = "FillDoc ID"


@dataclass
class Project:
    """
    MVP: проект = строка Excel + набор полей (ключ -> значение).
    Идентификатором считаем значение в колонке '№ дела' (если есть),
    иначе используем номер строки.
    """

    project_id: str
    fields: dict[str, str] = field(default_factory=dict)
    internal_id: str = field(default_factory=lambda: uuid.uuid4().hex)
    # Порядок колонок из Excel (как в заголовке), чтобы можно было
    # отображать поля в том же порядке, что и в таблице.
    headers: list[str] | None = None
    # Номер строки в Excel-листе (1-based), чтобы можно было
    # сохранять изменения без привязки к колонке "№ дела".
    row_index: int | None = None
    loaded_snapshot: str | None = None

    def __post_init__(self) -> None:
        excel_id = str(self.fields.get(FILLDOC_ID_FIELD, "") or "").strip()
        if excel_id:
            self.internal_id = excel_id
        elif not self.internal_id:
            self.internal_id = uuid.uuid4().hex

