from __future__ import annotations

from dataclasses import dataclass

from filldoc.variables.dictionary import VariableDictionary
from filldoc.variables.normalize import normalize_var_name


@dataclass(frozen=True)
class MissingField:
    raw_template_var: str
    resolved_key: str
    display_name: str


def compute_missing_fields(
    template_vars: list[str],
    project_fields: dict[str, str],
    dictionary: VariableDictionary,
) -> tuple[list[MissingField], list[MissingField]]:
    """
    Возвращает (missing, filled) для UI:
    - missing: поле отсутствует или пустое
    - filled: поле есть и заполнено
    """
    missing: list[MissingField] = []
    filled: list[MissingField] = []

    for raw in template_vars:
        entry = dictionary.resolve(raw)  # может быть None
        key = entry.display_name if entry else raw.strip()
        disp = entry.display_name if entry else raw.strip()

        # В MVP project_fields ключи = заголовки Excel как есть.
        # Попытка сопоставить “по нормализации”: ищем по нормализованному имени.
        val = ""
        if key in project_fields:
            val = str(project_fields.get(key, "") or "").strip()
        else:
            nk = normalize_var_name(key)
            for pk, pv in project_fields.items():
                if normalize_var_name(pk) == nk:
                    val = str(pv or "").strip()
                    break

        mf = MissingField(raw_template_var=raw, resolved_key=key, display_name=disp)
        if val:
            filled.append(mf)
        else:
            missing.append(mf)

    # Убираем повторы по resolved_key
    def uniq(items: list[MissingField]) -> list[MissingField]:
        seen: set[str] = set()
        out: list[MissingField] = []
        for it in items:
            if it.resolved_key in seen:
                continue
            seen.add(it.resolved_key)
            out.append(it)
        return out

    return uniq(missing), uniq(filled)

