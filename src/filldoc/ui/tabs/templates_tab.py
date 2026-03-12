from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QMessageBox,
    QTextEdit,
    QLabel,
    QSplitter,
)

from filldoc.core.settings import AppSettings
from filldoc.templates.scanner import TemplateLibrary
from filldoc.templates.models import TemplateCard


class TemplatesTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings()
        self._cards: list[TemplateCard] = []
        self._by_path: dict[str, TemplateCard] = {}

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        root.addLayout(top)

        self.refresh_btn = QPushButton("Обновить шаблоны")
        top.addWidget(self.refresh_btn)
        top.addStretch(1)

        split = QSplitter(self)
        root.addWidget(split, 1)

        self.tree = QTreeWidget(self)
        self.tree.setHeaderLabels(["Шаблоны"])
        split.addWidget(self.tree)

        right = QWidget(self)
        rlay = QVBoxLayout(right)
        rlay.addWidget(QLabel("Карточка шаблона (MVP):"))
        self.details = QTextEdit(self)
        self.details.setReadOnly(True)
        rlay.addWidget(self.details, 1)
        split.addWidget(right)
        split.setSizes([420, 620])

        self.refresh_btn.clicked.connect(self._scan)
        self.tree.currentItemChanged.connect(self._select_item)

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    def _scan(self) -> None:
        if not self._settings.templates_dir:
            QMessageBox.warning(self, "Шаблоны", "Не указан путь к библиотеке шаблонов (см. Настройки).")
            return
        try:
            lib = TemplateLibrary(self._settings.templates_dir)
            self._cards = lib.scan()
            self._by_path = {c.path: c for c in self._cards}
            self._render_tree()
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Шаблоны", str(e))

    def _render_tree(self) -> None:
        self.tree.clear()
        cat_nodes: dict[str, QTreeWidgetItem] = {}

        def folder_label(category: str) -> str:
            if not category:
                root_name = Path(self._settings.templates_dir).name.strip()
                return root_name or "(корень)"
            return Path(category).name or category

        def get_cat_node(category: str) -> QTreeWidgetItem:
            if category not in cat_nodes:
                node = QTreeWidgetItem([folder_label(category)])
                node.setData(0, 0, None)
                cat_nodes[category] = node
                self.tree.addTopLevelItem(node)
            return cat_nodes[category]

        for c in self._cards:
            cat = get_cat_node(c.category)
            item = QTreeWidgetItem([Path(c.path).name])
            item.setData(0, 0, c.path)
            cat.addChild(item)

        self.tree.expandAll()

    def _select_item(self, current, _prev) -> None:
        if not current:
            self.details.setPlainText("")
            return
        path = current.data(0, 0)
        if not path:
            self.details.setPlainText("")
            return
        card = self._by_path.get(path)
        if not card:
            self.details.setPlainText("")
            return

        text = []
        text.append(f"Имя файла: {Path(card.path).name}")
        text.append(f"Категория: {card.category}")
        text.append(f"Путь: {card.path}")
        text.append("")
        text.append(f"Переменные (уникальные): {len(card.variables_unique)}")
        for v in card.variables_unique:
            text.append(f"- {v}")
        self.details.setPlainText("\n".join(text))

