from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QByteArray, QSize, QTimer
from PySide6.QtGui import QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project
from filldoc.fill.missing_fields import compute_missing_fields
from filldoc.generator.docx_generator import generate_docx_from_template
from filldoc.templates.models import TemplateCard
from filldoc.templates.scanner import TemplateLibrary
from filldoc.variables.dictionary import default_dictionary

# ── SVG иконки ───────────────────────────────────────────────────────────────

_SVG_REFRESH = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M1 4v6h6"/>
  <path d="M23 20v-6h-6"/>
  <path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4-4.64 4.36A9 9 0 0 1 3.51 15"/>
</svg>"""

_SVG_SAVE = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
  <polyline points="17 21 17 13 7 13 7 21"/>
  <polyline points="7 3 7 8 15 8"/>
</svg>"""

_SVG_FOLDER = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.0"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
</svg>"""

# ── Стили ────────────────────────────────────────────────────────────────────

_BTN_STYLE = """
QToolButton {{
    background-color: {bg};
    border: none;
    border-radius: 9px;
    min-width:  38px;
    min-height: 38px;
    max-width:  38px;
    max-height: 38px;
}}
QToolButton:hover {{
    background-color: {hover};
}}
QToolButton:pressed {{
    background-color: {pressed};
}}
"""

_FILL_BTN_STYLE = """
QPushButton {
    background-color: #4A90D9;
    color: #ffffff;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    font-size: 13px;
    font-weight: bold;
}
QPushButton:hover {
    background-color: #357ABD;
}
QPushButton:pressed {
    background-color: #2A6099;
}
QPushButton:disabled {
    background-color: #9BB8D4;
    color: #e0e0e0;
}
"""

_SECTION_HEADER_STYLE = """
QLabel {
    font-size: 12px;
    font-weight: bold;
    color: #1e2a38;
    padding: 6px 0px 2px 0px;
}
"""

_VAR_LABEL_STYLE = """
QLabel {
    font-size: 11px;
    color: #4a5568;
    padding: 1px 4px;
}
"""

_SUCCESS_LABEL_STYLE = """
QLabel {
    font-size: 12px;
    color: #2d8a4e;
    padding: 4px 0px;
}
"""

_PLACEHOLDER_STYLE = """
QLabel {
    font-size: 13px;
    color: #9aa5b4;
    padding: 20px;
}
"""

_TABLE_STYLE = """
QTableWidget {
    border: 1px solid #dde2ea;
    border-radius: 4px;
    gridline-color: #edf0f5;
    font-size: 12px;
}
QTableWidget::item {
    padding: 4px 6px;
}
QHeaderView::section {
    background-color: #f4f6f9;
    border: none;
    border-bottom: 1px solid #dde2ea;
    padding: 4px 6px;
    font-size: 12px;
    font-weight: bold;
    color: #3a4a5c;
}
"""

# ── Вспомогательные функции ───────────────────────────────────────────────────

def _make_icon(svg_src: str, color: str = "#ffffff", size: int = 18) -> QIcon:
    colored = svg_src.replace("currentColor", color)
    data = QByteArray(colored.encode())
    renderer = QSvgRenderer(data)
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)


def _icon_btn(
    svg: str,
    tooltip: str,
    icon_color: str,
    bg: str,
    hover: str,
    pressed: str,
) -> QToolButton:
    btn = QToolButton()
    btn.setIcon(_make_icon(svg, icon_color))
    btn.setIconSize(QSize(18, 18))
    btn.setToolTip(tooltip)
    btn.setStyleSheet(_BTN_STYLE.format(bg=bg, hover=hover, pressed=pressed))
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def _h_separator() -> QFrame:
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setFrameShadow(QFrame.Shadow.Sunken)
    line.setStyleSheet("color: #dde2ea; margin: 4px 0px;")
    return line


# ── Основной класс ────────────────────────────────────────────────────────────

class TemplatesTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self._settings = AppSettings()
        self._cards: list[TemplateCard] = []
        self._by_path: dict[str, TemplateCard] = {}
        self._projects: list[Project] = []
        self._dict = default_dictionary()
        self._current_card: TemplateCard | None = None
        self._missing_table: QTableWidget | None = None

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(1200)
        self._autosave_timer.timeout.connect(self._autosave_to_excel)

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(6)

        # ── Верхняя панель ────────────────────────────────────────────
        top = QHBoxLayout()
        top.setSpacing(6)
        root.addLayout(top)

        top.addWidget(QLabel("Проект:"))
        self.project_combo = QComboBox(self)
        self.project_combo.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        top.addWidget(self.project_combo, 1)

        top.addStretch()

        self.refresh_btn = _icon_btn(
            _SVG_REFRESH,
            "Обновить шаблоны и проекты",
            "#ffffff", "#8A9BB0", "#6B7F96", "#556477",
        )
        self.save_btn = _icon_btn(
            _SVG_SAVE,
            "Сохранить заполненные поля в Excel",
            "#ffffff", "#8A9BB0", "#6B7F96", "#556477",
        )
        top.addWidget(self.refresh_btn)
        top.addWidget(self.save_btn)

        # ── Сплиттер: дерево | карточка ───────────────────────────────
        split = QSplitter(self)
        root.addWidget(split, 1)

        # Левая часть — дерево шаблонов
        left_wrap = QWidget(self)
        left_lay = QVBoxLayout(left_wrap)
        left_lay.setContentsMargins(0, 0, 0, 0)
        left_lay.setSpacing(4)

        left_hdr = QHBoxLayout()
        left_hdr.setContentsMargins(2, 0, 2, 0)
        left_hdr.setSpacing(4)
        left_hdr.addWidget(QLabel("Шаблоны"))
        self.open_templates_dir_btn = QToolButton(self)
        self.open_templates_dir_btn.setIcon(_make_icon(_SVG_FOLDER, "#5f6e80"))
        self.open_templates_dir_btn.setIconSize(QSize(16, 16))
        self.open_templates_dir_btn.setToolTip("Открыть папку шаблонов из настроек")
        self.open_templates_dir_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.open_templates_dir_btn.setAutoRaise(True)
        self.open_templates_dir_btn.setFixedSize(20, 20)
        left_hdr.addWidget(self.open_templates_dir_btn)
        left_hdr.addStretch(1)
        left_lay.addLayout(left_hdr)

        self.tree = QTreeWidget(self)
        self.tree.setHeaderHidden(True)
        left_lay.addWidget(self.tree, 1)
        split.addWidget(left_wrap)

        # Правая часть — область карточки с прокруткой
        self.card_scroll = QScrollArea(self)
        self.card_scroll.setWidgetResizable(True)
        self.card_scroll.setFrameShape(QFrame.Shape.NoFrame)

        self._card_container = QWidget()
        self._card_layout = QVBoxLayout(self._card_container)
        self._card_layout.setContentsMargins(0, 0, 0, 0)
        self._card_layout.setSpacing(0)

        self.card_scroll.setWidget(self._card_container)
        split.addWidget(self.card_scroll)
        split.setSizes([420, 620])

        # Показываем placeholder при старте
        self._show_placeholder()

        # ── Соединения ────────────────────────────────────────────────
        self.refresh_btn.clicked.connect(self._reload_all)
        self.save_btn.clicked.connect(self._save_to_excel)
        self.open_templates_dir_btn.clicked.connect(self._open_templates_dir)
        self.tree.currentItemChanged.connect(self._on_tree_item_changed)
        self.project_combo.currentIndexChanged.connect(self._on_project_changed)

    # ── Настройки ─────────────────────────────────────────────────────────────

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    # ── Загрузка данных ───────────────────────────────────────────────────────

    def _reload_all(self) -> None:
        self._load_projects()
        self._scan_templates()

    @staticmethod
    def _project_display_name(p: Project) -> str:
        """Returns a human-readable label for the project combo box."""
        for key in ("Имя проекта", "Номер осн. дела", "Номер дела", "№ дела", "№дела"):
            val = p.fields.get(key, "").strip()
            if val:
                return val
        # Last resort: first non-empty field value
        for v in p.fields.values():
            if str(v).strip():
                return str(v).strip()
        return p.project_id

    def _load_projects(self) -> None:
        if not self._settings.excel_path:
            return
        try:
            store = ExcelProjectStore(self._settings.excel_path)
            self._projects = store.load_projects()
            prev_id = self.project_combo.currentData()
            self.project_combo.blockSignals(True)
            self.project_combo.clear()
            for p in self._projects:
                self.project_combo.addItem(self._project_display_name(p), p.project_id)
            # Restore previous selection by stored project_id
            restore_idx = -1
            for i in range(self.project_combo.count()):
                if self.project_combo.itemData(i) == prev_id:
                    restore_idx = i
                    break
            if restore_idx >= 0:
                self.project_combo.setCurrentIndex(restore_idx)
            self.project_combo.blockSignals(False)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Шаблоны", f"Не удалось загрузить проекты: {e}")

    def _scan_templates(self) -> None:
        if not self._settings.templates_dir:
            QMessageBox.warning(
                self, "Шаблоны",
                "Не указан путь к библиотеке шаблонов (см. Настройки).",
            )
            return
        try:
            lib = TemplateLibrary(self._settings.templates_dir)
            self._cards = lib.scan()
            self._by_path = {c.path: c for c in self._cards}
            self._render_tree()
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Шаблоны", str(e))

    # ── Дерево шаблонов ───────────────────────────────────────────────────────

    def _render_tree(self) -> None:
        self.tree.clear()
        cat_nodes: dict[str, QTreeWidgetItem] = {}

        for c in self._cards:
            category_key = c.category
            if category_key not in cat_nodes:
                if category_key:
                    label = Path(category_key).name or category_key
                else:
                    label = Path(self._settings.templates_dir).name or "(корень)"
                node = QTreeWidgetItem([label])
                node.setData(0, Qt.ItemDataRole.UserRole, None)
                cat_nodes[category_key] = node
                self.tree.addTopLevelItem(node)

            parent_node = cat_nodes[category_key]
            item = QTreeWidgetItem([Path(c.path).stem])
            item.setData(0, Qt.ItemDataRole.UserRole, c.path)
            parent_node.addChild(item)

        self.tree.expandAll()

    # ── Обработчики событий ───────────────────────────────────────────────────

    def _on_tree_item_changed(self, current, _prev) -> None:
        if not current:
            self._current_card = None
            self._show_placeholder()
            return
        path = current.data(0, Qt.ItemDataRole.UserRole)
        if not path:
            self._current_card = None
            self._show_placeholder()
            return
        card = self._by_path.get(path)
        if not card:
            self._current_card = None
            self._show_placeholder()
            return
        self._current_card = card
        self._build_card(card)

    def _on_project_changed(self) -> None:
        if self._current_card:
            self._build_card(self._current_card)

    # ── Карточка шаблона ─────────────────────────────────────────────────────

    def _clear_card_layout(self) -> None:
        self._missing_table = None
        while self._card_layout.count():
            child = self._card_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def _show_placeholder(self) -> None:
        self._clear_card_layout()
        lbl = QLabel("← Выберите шаблон в списке слева")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(_PLACEHOLDER_STYLE)
        self._card_layout.addWidget(lbl)

    def _build_card(self, card: TemplateCard) -> None:
        self._clear_card_layout()

        project = self._current_project()

        if project:
            missing_fields, filled_fields = compute_missing_fields(
                card.variables_unique, project.fields, self._dict
            )
        else:
            missing_fields, filled_fields = [], []

        lay = self._card_layout
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(6)

        # ── Секция: Переменные в реквизитах ──────────────────────────
        filled_hdr = QLabel(
            f"Переменные в реквизитах"
            + (f" ({len(filled_fields)})" if project else "")
            + ":"
        )
        filled_hdr.setStyleSheet(_SECTION_HEADER_STYLE)
        lay.addWidget(filled_hdr)

        if not project:
            lbl = QLabel("Выберите проект для анализа переменных")
            lbl.setStyleSheet(_PLACEHOLDER_STYLE)
            lay.addWidget(lbl)
        elif filled_fields:
            for mf in filled_fields:
                lbl = QLabel(f"• {mf.display_name}")
                lbl.setStyleSheet(_VAR_LABEL_STYLE)
                lay.addWidget(lbl)
        else:
            lbl = QLabel("(нет заполненных переменных)")
            lbl.setStyleSheet(_VAR_LABEL_STYLE)
            lay.addWidget(lbl)

        lay.addWidget(_h_separator())

        # ── Секция: Недостающие переменные ────────────────────────────
        count_str = str(len(missing_fields)) if project else "—"
        missing_hdr = QLabel(f"Недостающие переменные ({count_str}):")
        missing_hdr.setStyleSheet(_SECTION_HEADER_STYLE)
        lay.addWidget(missing_hdr)

        if project and missing_fields:
            tbl = QTableWidget(len(missing_fields), 2)
            tbl.horizontalHeader().setVisible(False)
            tbl.horizontalHeader().setSectionResizeMode(
                0, QHeaderView.ResizeMode.ResizeToContents
            )
            tbl.horizontalHeader().setSectionResizeMode(
                1, QHeaderView.ResizeMode.Stretch
            )
            tbl.verticalHeader().setVisible(False)
            tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            tbl.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
            tbl.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
            tbl.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
            tbl.setStyleSheet(_TABLE_STYLE)

            for i, mf in enumerate(missing_fields):
                name_item = QTableWidgetItem(mf.display_name)
                name_item.setFlags(
                    name_item.flags() & ~Qt.ItemFlag.ItemIsEditable
                )
                tbl.setItem(i, 0, name_item)
                tbl.setItem(i, 1, QTableWidgetItem(""))

            row_h = 26
            tbl.verticalHeader().setDefaultSectionSize(row_h)
            tbl.resizeRowsToContents()
            total_h = sum(tbl.rowHeight(r) for r in range(tbl.rowCount()))
            total_h += 6  # небольшие отступы/рамка
            tbl.setFixedHeight(total_h)

            self._missing_table = tbl
            tbl.itemChanged.connect(self._schedule_autosave)
            lay.addWidget(tbl)

        elif project:
            lbl = QLabel("✓ Все переменные шаблона заполнены!")
            lbl.setStyleSheet(_SUCCESS_LABEL_STYLE)
            lay.addWidget(lbl)
        else:
            lbl = QLabel("Выберите проект для анализа недостающих переменных")
            lbl.setStyleSheet(_VAR_LABEL_STYLE)
            lay.addWidget(lbl)

        # ── Растяжка ─────────────────────────────────────────────────
        lay.addStretch(1)

        # ── Кнопка «Заполнить шаблон» ────────────────────────────────
        fill_btn = QPushButton("Заполнить шаблон")
        fill_btn.setStyleSheet(_FILL_BTN_STYLE)
        fill_btn.setEnabled(bool(project))
        fill_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        fill_btn.clicked.connect(self._fill_template)
        lay.addWidget(fill_btn)

    # ── Вспомогательные методы ────────────────────────────────────────────────

    def _current_project(self) -> Project | None:
        pid = self.project_combo.currentData()
        for p in self._projects:
            if p.project_id == pid:
                return p
        return None

    def _collect_missing_values(self) -> dict[str, str]:
        result: dict[str, str] = {}
        if not self._missing_table:
            return result
        for r in range(self._missing_table.rowCount()):
            k_item = self._missing_table.item(r, 0)
            v_item = self._missing_table.item(r, 1)
            k = (k_item.text() if k_item else "").strip()
            v = (v_item.text() if v_item else "").strip()
            if k:
                result[k] = v
        return result

    def _schedule_autosave(self) -> None:
        self._autosave_timer.start()

    def _autosave_to_excel(self) -> None:
        project = self._current_project()
        if not project or not self._settings.excel_path:
            return
        extra = self._collect_missing_values()
        for k, v in extra.items():
            if v.strip():
                project.fields[k] = v.strip()
        try:
            from filldoc.excel.excel_store import ExcelProjectStore
            store = ExcelProjectStore(self._settings.excel_path)
            store.save_project_fields(project)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Автосохранение", str(e))

    def _open_templates_dir(self) -> None:
        templates_dir = (self._settings.templates_dir or "").strip()
        if not templates_dir:
            QMessageBox.warning(
                self,
                "Шаблоны",
                "Папка шаблонов не указана в настройках.",
            )
            return

        path = Path(templates_dir)
        if not path.exists() or not path.is_dir():
            QMessageBox.warning(
                self,
                "Шаблоны",
                "Папка шаблонов недоступна или не существует.",
            )
            return

        try:
            if sys.platform == "win32":
                os.startfile(str(path))  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])  # noqa: S603
            else:
                subprocess.Popen(["xdg-open", str(path)])  # noqa: S603
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(
                self,
                "Шаблоны",
                f"Не удалось открыть папку шаблонов: {e}",
            )

    # ── Сохранение ────────────────────────────────────────────────────────────

    def _save_to_excel(self) -> None:
        project = self._current_project()
        if not project:
            QMessageBox.warning(self, "Сохранение", "Нет выбранного проекта.")
            return
        if not self._settings.excel_path:
            QMessageBox.warning(
                self, "Сохранение",
                "Не указан путь к Excel-файлу (см. Настройки).",
            )
            return

        extra = self._collect_missing_values()
        for k, v in extra.items():
            if v.strip():
                project.fields[k] = v.strip()

        try:
            store = ExcelProjectStore(self._settings.excel_path)
            store.save_project_fields(project)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Сохранение", str(e))
            return

        if self._current_card:
            self._build_card(self._current_card)

        QMessageBox.information(self, "Сохранение", "Изменения сохранены в Excel.")

    # ── Заполнение шаблона ────────────────────────────────────────────────────

    def _fill_template(self) -> None:
        project = self._current_project()
        if not project:
            QMessageBox.warning(self, "Заполнение", "Выберите проект.")
            return
        card = self._current_card
        if not card:
            QMessageBox.warning(self, "Заполнение", "Выберите шаблон.")
            return
        if not self._settings.output_dir:
            QMessageBox.warning(
                self, "Заполнение",
                "Не указан каталог вывода (см. Настройки).",
            )
            return

        # Применяем введённые значения
        extra = self._collect_missing_values()
        for k, v in extra.items():
            if v.strip():
                project.fields[k] = v.strip()

        # Строим словарь подстановки
        mapping: dict[str, str] = {}
        for k, v in project.fields.items():
            mapping[str(k)] = str(v or "")
        for raw in card.variables_unique:
            entry = self._dict.resolve(raw)
            if entry:
                val = (
                    project.fields.get(entry.display_name)
                    or project.fields.get(entry.technical_name)
                    or ""
                )
                mapping[raw] = str(val)
            else:
                mapping[raw] = str(project.fields.get(raw, ""))

        customer = (project.fields.get("Заказчик") or "").strip()
        debtor = (project.fields.get("Должник") or "").strip()
        parts = [card.name]
        if customer:
            parts.append(customer)
        if debtor:
            parts.append(debtor)
        out_name = " - ".join(parts)

        try:
            out_path = generate_docx_from_template(
                card.path, self._settings.output_dir, out_name, mapping
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Генерация", str(e))
            return

        out_folder = str(Path(out_path).parent)
        try:
            if sys.platform == "win32":
                os.startfile(out_folder)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", out_folder])  # noqa: S603
            else:
                subprocess.Popen(["xdg-open", out_folder])  # noqa: S603
        except Exception:  # noqa: BLE001
            pass

        QMessageBox.information(self, "Готово", f"Файл создан:\n{out_path}")
