from __future__ import annotations

from dataclasses import dataclass, field

from .normalize import normalize_var_name


@dataclass(frozen=True)
class VariableEntry:
    technical_name: str
    display_name: str
    variants: set[str] = field(default_factory=set)
    field_type: str = "text"  # MVP: text / multiline / select
    group: str = "project"  # project / template / system
    required: bool = False
    comment: str = ""


class VariableDictionary:
    """
    MVP-реализация словаря:
    - хранит сопоставление вариантов написания -> техническое имя
    - умеет нормализовать и сопоставлять по нормализованному имени
    """

    def __init__(self) -> None:
        self._by_norm: dict[str, VariableEntry] = {}

    def add(self, entry: VariableEntry) -> None:
        key = normalize_var_name(entry.technical_name)
        self._by_norm[key] = entry
        for v in entry.variants:
            self._by_norm[normalize_var_name(v)] = entry
        self._by_norm[normalize_var_name(entry.display_name)] = entry

    def resolve(self, raw_name: str) -> VariableEntry | None:
        return self._by_norm.get(normalize_var_name(raw_name))

    def all_entries(self) -> list[VariableEntry]:
        uniq = {id(v): v for v in self._by_norm.values()}
        return sorted(uniq.values(), key=lambda e: e.display_name.lower())


def default_dictionary() -> VariableDictionary:
    """
    Минимальная база на основе ТЗ (без “умной” расширенной нормализации).
    Можно расширять во 2 этапе.
    """
    d = VariableDictionary()
    d.add(
        VariableEntry(
            technical_name="№ дела",
            display_name="№ дела",
            variants={"номер дела", "№дела"},
        )
    )
    d.add(VariableEntry(technical_name="Заказчик", display_name="Заказчик", variants={"ЗАКАЗЧИК"}))
    d.add(VariableEntry(technical_name="Должник", display_name="Должник", variants={"ДОЛЖНИК"}))
    d.add(VariableEntry(technical_name="ИНН должника", display_name="ИНН должника", variants={"ИНН Должник", "ИНН ДОЛЖНИКА"}))
    d.add(VariableEntry(technical_name="ОГРН должника", display_name="ОГРН должника", variants={"ОГРН Должник", "ОГРН Должника"}))
    d.add(VariableEntry(technical_name="Юр адрес должника", display_name="Юр. адрес должника", variants={"Юр. адрес Должник", "Юр Адрес должника"}))
    d.add(VariableEntry(technical_name="Дата решения", display_name="Дата решения", variants={"дата решения", "Дата решения"}))
    d.add(VariableEntry(technical_name="Вид услуг", display_name="Вид услуг", variants={"по оплате(вид услуг)"}))
    d.add(VariableEntry(technical_name="Цедент полное наименование", display_name="Цедент полное наименование", variants={"Цедент П.Наим"}))
    d.add(VariableEntry(technical_name="Цена уступки", display_name="Цена уступки", variants={"цена уступки"}))
    d.add(VariableEntry(technical_name="Цена прописью", display_name="Цена прописью", variants={"цена прописью"}))
    d.add(
        VariableEntry(
            technical_name="Корреспондентский счет банка заказчика",
            display_name="Корреспондентский счет банка заказчика",
            variants={"к/сч. Банка Заказчика"},
        )
    )
    d.add(VariableEntry(technical_name="Расчетный счет заказчика", display_name="Расчетный счет заказчика", variants={"р/сч. Заказчика"}))
    d.add(VariableEntry(technical_name="Юр адрес заказчика", display_name="Юр. адрес заказчика", variants={"Юр Адрес Заказчика"}))
    d.add(VariableEntry(technical_name="Юр адрес цедента", display_name="Юр. адрес цедента", variants={"Юр. адрес цедента"}))
    d.add(VariableEntry(technical_name="Номер цессии", display_name="Номер цессии", variants={"Номер Цессии"}))
    d.add(VariableEntry(technical_name="Резолютивка", display_name="Резолютивка", variants={"Резюлютивка"}))
    d.add(VariableEntry(technical_name="Резолютивка – взыскать с", display_name="Резолютивка – взыскать с"))
    return d

