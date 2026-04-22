from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path

from filldoc.core.errors import TemplateError
from .models import TemplateCard
from .vars_extractor import extract_docx_variables


def _cards_dir(templates_root: Path) -> Path:
    return templates_root / ".filldoc"


def _card_path(templates_root: Path, rel: Path) -> Path:
    safe = str(rel).replace("\\", "__").replace("/", "__")
    return _cards_dir(templates_root) / f"{safe}.json"


class TemplateLibrary:
    def __init__(self, templates_dir: str) -> None:
        self.templates_dir = templates_dir

    def scan(self) -> list[TemplateCard]:
        root = Path(self.templates_dir)
        if not root.exists():
            raise TemplateError("Папка библиотеки шаблонов недоступна или не существует.")

        cards: list[TemplateCard] = []
        for p in root.rglob("*.docx"):
            if p.name.startswith("~$"):  # временные файлы Word
                continue
            try:
                rel = p.relative_to(root)
            except Exception:  # noqa: BLE001
                rel = Path(p.name)
            category = str(rel.parent) if hasattr(rel, "parent") else ""
            variables_in_order, variables_unique = extract_docx_variables(str(p))

            # Читаем существующий кеш, чтобы сохранить ручные правки карточки.
            cached = self._load_card(root, rel if isinstance(rel, Path) else Path(p.name))

            card = TemplateCard(
                name=cached.name if cached is not None and cached.name else p.stem,
                path=str(p),
                category=(
                    cached.category
                    if cached is not None
                    else category if category != "." else ""
                ),
                variables_in_order=variables_in_order,
                variables_unique=variables_unique,
                output_name_rule=(
                    cached.output_name_rule
                    if cached is not None
                    else "{%filename%} - {ДОЛЖНИК}"
                ),
                active=cached.active if cached is not None else True,
                comment=cached.comment if cached is not None else "",
            )
            cards.append(card)
            self._save_card(root, rel if isinstance(rel, Path) else Path(p.name), card)
        return sorted(cards, key=lambda c: (c.category.lower(), c.name.lower()))

    def save_card(self, card: TemplateCard) -> None:
        root = Path(self.templates_dir)
        try:
            rel = Path(card.path).relative_to(root)
        except Exception:  # noqa: BLE001
            rel = Path(card.path).name
        self._save_card(root, Path(rel), card)

    def _load_card(self, root: Path, rel: Path) -> TemplateCard | None:
        """Загружает сохранённую карточку из JSON-кеша; возвращает None если не найдена."""
        try:
            path = _card_path(root, rel)
            if not path.exists():
                return None
            data = json.loads(path.read_text(encoding="utf-8"))
            return TemplateCard(
                name=data.get("name", ""),
                path=data.get("path", ""),
                category=data.get("category", ""),
                variables_in_order=data.get("variables_in_order", []),
                variables_unique=data.get("variables_unique", []),
                output_name_rule=data.get("output_name_rule", "{%filename%} - {ДОЛЖНИК}"),
                active=data.get("active", True),
                comment=data.get("comment", ""),
            )
        except Exception:  # noqa: BLE001
            return None

    def _save_card(self, root: Path, rel: Path, card: TemplateCard) -> None:
        try:
            d = _cards_dir(root)
            d.mkdir(parents=True, exist_ok=True)
            _card_path(root, rel).write_text(
                json.dumps(asdict(card), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:  # noqa: BLE001
            raise TemplateError(f"Не удалось сохранить карточку шаблона для '{card.name}': {e}") from e

