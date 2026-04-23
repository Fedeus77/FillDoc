from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QByteArray, QSize, QTimer, Signal
from PySide6.QtGui import QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.models import FILLDOC_ID_FIELD
from filldoc.projects.repository import ProjectRepository
from filldoc.templates.scanner import TemplateLibrary
from filldoc.ui.icons import SVG_ADD, SVG_REFRESH, SVG_SAVE, icon_btn, make_icon, update_icon_btn
from filldoc.ui.theme import ThemeColors, ThemeManager
from filldoc.variables.dictionary import (
    FIELD_TYPES,
    VariableDictionary,
    VariableEntry,
    load_variable_dictionary,
    load_variable_entries,
    save_entries_to_file,
    user_variables_path,
)
from filldoc.variables.normalize import normalize_var_name


_SVG_COPY = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2"
     stroke-linecap="round" stroke-linejoin="round">
  <rect x="9" y="9" width="13" height="13" rx="2" ry="2"/>
  <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>
</svg>"""

_SVG_CHECK = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.5"
     stroke-linecap="round" stroke-linejoin="round">
  <polyline points="20 6 9 17 4 12"/>
</svg>"""

_SVG_EXCEL = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <path d="M14 2v6h6"/>
  <path d="M8 12l4 6"/>
  <path d="M12 12l-4 6"/>
</svg>"""

_SVG_DELETE = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.1"
     stroke-linecap="round" stroke-linejoin="round">
  <polyline points="3 6 5 6 21 6"/>
  <path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>
  <path d="M10 11v6"/>
  <path d="M14 11v6"/>
  <path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>
</svg>"""


def _make_icon(svg_src: str, color: str, size: int = 16) -> QIcon:
    colored = svg_src.replace("currentColor", color)
    data = QByteArray(colored.encode())
    renderer = QSvgRenderer(data)
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)


def _copy_btn(c: ThemeColors) -> QToolButton:
    btn = QToolButton()
    btn.setIcon(_make_icon(_SVG_COPY, c.text_muted, 13))
    btn.setIconSize(QSize(13, 13))
    btn.setFixedSize(20, 20)
    btn.setToolTip("Копировать")
    btn.setStyleSheet(f"""
QToolButton {{ background-color: transparent; border: none; border-radius: 4px; padding: 2px; }}
QToolButton:hover {{ background-color: {c.icon_btn_ghost_hover}; }}
QToolButton:pressed {{ background-color: {c.icon_btn_pressed}; }}
""")
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def _input_style(c: ThemeColors) -> str:
    return f"""
QLineEdit, QTextEdit, QComboBox {{
    background-color: {c.bg_input};
    color: {c.text_primary};
    border: 1px solid {c.border_input};
    border-radius: 8px;
    padding: 6px 8px;
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
}}
QLineEdit:focus, QTextEdit:focus, QComboBox:focus {{
    border-color: {c.text_accent};
    background-color: {c.bg_input_focus};
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox QAbstractItemView {{
    background-color: {c.bg_panel};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
}}
"""


def _table_style(c: ThemeColors) -> str:
    return f"""
QTableWidget {{
    background-color: {c.bg_panel};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 8px;
    gridline-color: {c.border_light};
    font-size: 12px;
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
}}
QTableWidget::item {{
    padding: 5px 7px;
}}
QHeaderView::section {{
    background-color: {c.bg_header};
    border: none;
    border-bottom: 1px solid {c.border_base};
    padding: 5px 7px;
    font-size: 12px;
    font-weight: 700;
    color: {c.text_secondary};
}}
"""


def _entry_key(entry: VariableEntry) -> str:
    return normalize_var_name(entry.technical_name)


class VariablesTab(QWidget):
    dictionary_changed = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self._settings = AppSettings()
        self._headers: list[str] = []
        self._entries: list[VariableEntry] = []
        self._dict = VariableDictionary()
        self._usage_by_key: dict[str, list[str]] = {}
        self._theme_colors = ThemeManager.instance().colors
        self._updating_form = False

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(4)

        self.tabs = QTabWidget(self)
        root.addWidget(self.tabs, 1)

        self._build_excel_tab()
        self._build_dictionary_tab()

        self.refresh_btn.clicked.connect(self._reload)
        self.open_excel_btn.clicked.connect(self._open_excel_file)
        self.dict_reload_btn.clicked.connect(self._reload_dictionary)
        self.dict_save_btn.clicked.connect(self._save_dictionary)
        self.dict_add_btn.clicked.connect(self._add_variable)
        self.dict_delete_btn.clicked.connect(self._delete_variable)
        self.alias_add_btn.clicked.connect(self._add_alias)
        self.dictionary_table.currentCellChanged.connect(self._on_dictionary_row_changed)

        for widget in (
            self.technical_edit,
            self.display_edit,
            self.group_edit,
            self.aliases_edit,
            self.comment_edit,
        ):
            if isinstance(widget, QLineEdit):
                widget.textChanged.connect(self._on_form_changed)
            else:
                widget.textChanged.connect(self._on_form_changed)
        self.field_type_combo.currentTextChanged.connect(self._on_form_changed)
        self.required_check.stateChanged.connect(self._on_form_changed)

        self._show_empty("Нажмите  ↻  для загрузки переменных из Excel")
        self._reload_dictionary(show_errors=False)

    def _build_excel_tab(self) -> None:
        tab = QWidget(self)
        root = QVBoxLayout(tab)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(4)

        top = QHBoxLayout()
        top.setSpacing(6)
        self._desc_label = QLabel("Переменные из таблицы Excel:")
        top.addWidget(self._desc_label, 1)

        self.open_excel_btn = QToolButton(self)
        self.open_excel_btn.setIcon(make_icon(_SVG_EXCEL, "#2d8a4e", 18))
        self.open_excel_btn.setIconSize(QSize(18, 18))
        self.open_excel_btn.setToolTip("Открыть Excel-файл")
        self.open_excel_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        top.addWidget(self.open_excel_btn)

        self.refresh_btn = icon_btn(SVG_REFRESH, "Обновить поля Excel")
        top.addWidget(self.refresh_btn)
        root.addLayout(top)

        self._sep = QFrame()
        self._sep.setFrameShape(QFrame.Shape.HLine)
        self._sep.setFrameShadow(QFrame.Shadow.Sunken)
        root.addWidget(self._sep)

        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)

        self._list_container = QWidget()
        self._list_layout = QVBoxLayout(self._list_container)
        self._list_layout.setContentsMargins(4, 2, 4, 2)
        self._list_layout.setSpacing(0)

        self.scroll.setWidget(self._list_container)
        root.addWidget(self.scroll, 1)
        self.tabs.addTab(tab, "Excel")

    def _build_dictionary_tab(self) -> None:
        tab = QWidget(self)
        root = QVBoxLayout(tab)
        root.setContentsMargins(6, 6, 6, 6)
        root.setSpacing(6)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(6)
        self.dict_add_btn = icon_btn(SVG_ADD, "Добавить переменную")
        self.dict_save_btn = icon_btn(SVG_SAVE, "Сохранить словарь")
        self.dict_reload_btn = icon_btn(SVG_REFRESH, "Обновить словарь")
        self.dict_delete_btn = icon_btn(_SVG_DELETE, "Удалить переменную")
        toolbar.addWidget(self.dict_add_btn)
        toolbar.addWidget(self.dict_save_btn)
        toolbar.addWidget(self.dict_reload_btn)
        toolbar.addWidget(self.dict_delete_btn)
        toolbar.addStretch(1)
        self.dict_path_label = QLabel(str(user_variables_path()))
        toolbar.addWidget(self.dict_path_label)
        root.addLayout(toolbar)

        split = QSplitter(self)
        root.addWidget(split, 1)

        self.dictionary_table = QTableWidget(0, 4, self)
        self.dictionary_table.setHorizontalHeaderLabels(["Поле", "Тех. имя", "Тип", "Обяз."])
        self.dictionary_table.verticalHeader().setVisible(False)
        self.dictionary_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.dictionary_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.dictionary_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.dictionary_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.dictionary_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.dictionary_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.dictionary_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        split.addWidget(self.dictionary_table)

        form_wrap = QWidget(self)
        form_root = QVBoxLayout(form_wrap)
        form_root.setContentsMargins(12, 0, 0, 0)
        form_root.setSpacing(8)

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(6)

        self.technical_edit = QLineEdit(self)
        self.display_edit = QLineEdit(self)
        self.group_edit = QLineEdit(self)
        self.field_type_combo = QComboBox(self)
        self.field_type_combo.addItems(FIELD_TYPES)
        self.required_check = QCheckBox("Обязательная")
        self.aliases_edit = QTextEdit(self)
        self.aliases_edit.setFixedHeight(86)
        self.comment_edit = QTextEdit(self)
        self.comment_edit.setFixedHeight(72)

        alias_row = QHBoxLayout()
        alias_row.setSpacing(6)
        self.alias_input = QLineEdit(self)
        self.alias_add_btn = icon_btn(SVG_ADD, "Добавить alias", button_size=32, icon_size=16)
        alias_row.addWidget(self.alias_input, 1)
        alias_row.addWidget(self.alias_add_btn)
        alias_wrap = QWidget(self)
        alias_wrap.setLayout(alias_row)

        form.addRow("Техническое имя:", self.technical_edit)
        form.addRow("Название:", self.display_edit)
        form.addRow("Группа:", self.group_edit)
        form.addRow("Тип поля:", self.field_type_combo)
        form.addRow("Обязательность:", self.required_check)
        form.addRow("Новый alias:", alias_wrap)
        form.addRow("Aliases:", self.aliases_edit)
        form.addRow("Комментарий:", self.comment_edit)
        form_root.addLayout(form)

        self.usage_label = QLabel("Использование в шаблонах")
        self.usage_text = QTextEdit(self)
        self.usage_text.setReadOnly(True)
        self.usage_text.setFixedHeight(132)
        form_root.addWidget(self.usage_label)
        form_root.addWidget(self.usage_text)
        form_root.addStretch(1)
        split.addWidget(form_wrap)
        split.setSizes([620, 420])

        self.tabs.addTab(tab, "Словарь")

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    def apply_theme(self, c: ThemeColors) -> None:
        self._theme_colors = c
        self._desc_label.setStyleSheet(
            f"QLabel {{ font-size: 12px; color: {c.text_secondary}; padding: 2px 0px 6px 0px; }}"
        )
        self._sep.setStyleSheet(f"color: {c.separator};")
        self.tabs.setStyleSheet("")
        update_icon_btn(
            self.refresh_btn,
            SVG_REFRESH,
            icon_color=c.icon_color,
            bg=c.icon_btn_bg,
            hover=c.icon_btn_hover,
            pressed=c.icon_btn_pressed,
        )
        self.open_excel_btn.setIcon(make_icon(_SVG_EXCEL, c.success, 18))
        self.open_excel_btn.setStyleSheet(self._ghost_button_style(c, 34))

        for btn, svg in (
            (self.dict_add_btn, SVG_ADD),
            (self.dict_save_btn, SVG_SAVE),
            (self.dict_reload_btn, SVG_REFRESH),
            (self.dict_delete_btn, _SVG_DELETE),
            (self.alias_add_btn, SVG_ADD),
        ):
            update_icon_btn(
                btn,
                svg,
                icon_color=c.icon_color,
                bg=c.icon_btn_bg,
                hover=c.icon_btn_hover,
                pressed=c.icon_btn_pressed,
                button_size=32 if btn is self.alias_add_btn else 38,
                icon_size=16 if btn is self.alias_add_btn else 18,
            )

        self.dictionary_table.setStyleSheet(_table_style(c))
        self.dict_path_label.setStyleSheet(f"QLabel {{ color: {c.text_muted}; font-size: 11px; }}")
        self.usage_label.setStyleSheet(f"QLabel {{ color: {c.text_secondary}; font-weight: 700; }}")
        self.required_check.setStyleSheet(f"QCheckBox {{ color: {c.text_primary}; }}")
        for widget in (
            self.technical_edit,
            self.display_edit,
            self.group_edit,
            self.alias_input,
            self.aliases_edit,
            self.comment_edit,
            self.usage_text,
            self.field_type_combo,
        ):
            widget.setStyleSheet(_input_style(c))

        if self._headers:
            self._render_list()
        self._render_dictionary_table(self._current_entry_key())

    @staticmethod
    def _ghost_button_style(c: ThemeColors, size: int) -> str:
        return f"""
QToolButton {{
    background-color: transparent;
    border: none;
    border-radius: 9px;
    min-width: {size}px; min-height: {size}px;
    max-width: {size}px; max-height: {size}px;
}}
QToolButton:hover {{ background-color: {c.icon_btn_ghost_hover}; }}
QToolButton:pressed {{ background-color: {c.icon_btn_pressed}; }}
"""

    def _show_status(self, message: str, timeout_ms: int = 4000) -> None:
        mw = self.window()
        if hasattr(mw, "show_status"):
            mw.show_status(message, timeout_ms)

    def _reload(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(
                self,
                "Переменные",
                "Не указан путь к Excel-файлу (см. Настройки).",
            )
            return
        try:
            projects = ProjectRepository(self._settings.excel_path).load_projects()
            if projects:
                self._headers = [h for h in projects[0].headers if h.strip() and h != FILLDOC_ID_FIELD]
            else:
                self._headers = []
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Переменные", str(e))
            return

        self._render_list()

    def _open_excel_file(self) -> None:
        excel_path = (self._settings.excel_path or "").strip()
        if not excel_path:
            QMessageBox.warning(self, "Переменные", "Не указан путь к Excel-файлу (см. Настройки).")
            return

        path = Path(excel_path)
        if not path.exists() or not path.is_file():
            QMessageBox.warning(self, "Переменные", "Excel-файл недоступен или не существует.")
            return

        try:
            if sys.platform == "win32":
                os.startfile(str(path))  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])  # noqa: S603
            else:
                subprocess.Popen(["xdg-open", str(path)])  # noqa: S603
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Переменные", f"Не удалось открыть Excel-файл:\n{e}")

    def _clear_list(self) -> None:
        while self._list_layout.count():
            child = self._list_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def _show_empty(self, msg: str) -> None:
        self._clear_list()
        c = self._theme_colors
        lbl = QLabel(msg)
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl.setStyleSheet(f"QLabel {{ font-size: 13px; color: {c.text_muted}; padding: 24px; }}")
        self._list_layout.addWidget(lbl)
        self._list_layout.addStretch(1)

    def _render_list(self) -> None:
        self._clear_list()

        if not self._headers:
            self._show_empty("В таблице не найдено ни одного поля")
            return

        for header in self._headers:
            self._list_layout.addWidget(self._make_var_row(f"{{{header}}}"))
        self._list_layout.addStretch(1)

    def _make_var_row(self, var_text: str) -> QFrame:
        c = self._theme_colors
        frame = QFrame()
        frame.setStyleSheet(f"""
QFrame {{ background-color: transparent; border: none; border-radius: 3px; }}
QFrame:hover {{ background-color: {c.bg_hover}; }}
""")
        frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        lay = QHBoxLayout(frame)
        lay.setContentsMargins(4, 1, 4, 1)
        lay.setSpacing(4)

        copy_btn = _copy_btn(c)
        copy_btn.clicked.connect(lambda _checked, t=var_text, b=copy_btn: self._copy_var(t, b))
        lay.addWidget(copy_btn)

        lbl = QLabel(var_text)
        lbl.setStyleSheet(f"""
QLabel {{
    font-family: Consolas, 'Courier New', monospace;
    font-size: 12px;
    color: {c.text_accent};
    padding: 0px;
    background: transparent;
}}
""")
        lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        lay.addWidget(lbl, 1)
        return frame

    def _copy_var(self, text: str, btn: QToolButton) -> None:
        c = self._theme_colors
        QApplication.clipboard().setText(text)
        btn.setIcon(_make_icon(_SVG_CHECK, c.success, 14))
        btn.setToolTip("Скопировано!")
        QTimer.singleShot(
            1300,
            lambda: (
                btn.setIcon(_make_icon(_SVG_COPY, c.text_muted, 14)),
                btn.setToolTip("Копировать"),
            ),
        )

    def _reload_dictionary(self, show_errors: bool = True) -> None:
        try:
            self._entries = load_variable_entries()
            self._dict = load_variable_dictionary()
            self._reload_dictionary_usage(show_errors=False)
        except Exception as e:  # noqa: BLE001
            if show_errors:
                QMessageBox.critical(self, "Словарь", f"Не удалось загрузить словарь:\n{e}")
            self._entries = []
            self._dict = VariableDictionary()
            self._usage_by_key = {}
        self._render_dictionary_table()

    def _reload_dictionary_usage(self, show_errors: bool = True) -> None:
        self._usage_by_key = {}
        templates_dir = (self._settings.templates_dir or "").strip()
        if not templates_dir or not Path(templates_dir).is_dir():
            return

        try:
            cards = TemplateLibrary(templates_dir).scan()
            usage: dict[str, list[str]] = {}
            dictionary = VariableDictionary([entry for entry in self._entries if entry.technical_name.strip()])
            for card in cards:
                label = f"{card.category + ' / ' if card.category else ''}{card.name}"
                for raw in card.variables_unique:
                    entry = dictionary.resolve(raw)
                    if entry is None:
                        continue
                    usage.setdefault(_entry_key(entry), [])
                    if label not in usage[_entry_key(entry)]:
                        usage[_entry_key(entry)].append(label)
            self._usage_by_key = {key: sorted(values, key=str.lower) for key, values in usage.items()}
        except Exception as e:  # noqa: BLE001
            if show_errors:
                QMessageBox.warning(self, "Словарь", f"Не удалось обновить использование в шаблонах:\n{e}")

    def _render_dictionary_table(self, select_key: str | None = None) -> None:
        if select_key is None:
            select_key = self._current_entry_key()
        self._entries.sort(key=lambda e: (e.display_name.lower(), e.technical_name.lower()))

        self.dictionary_table.blockSignals(True)
        self.dictionary_table.setRowCount(len(self._entries))
        selected_row = -1
        for row, entry in enumerate(self._entries):
            values = [
                entry.display_name,
                entry.technical_name,
                entry.field_type,
                "Да" if entry.required else "",
            ]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setData(Qt.ItemDataRole.UserRole, _entry_key(entry))
                self.dictionary_table.setItem(row, col, item)
            if select_key and _entry_key(entry) == select_key:
                selected_row = row

        self.dictionary_table.blockSignals(False)
        if selected_row < 0 and self._entries:
            selected_row = 0
        if selected_row >= 0:
            self.dictionary_table.selectRow(selected_row)
            self._load_entry_into_form(selected_row)
        else:
            self._clear_form()

    def _current_entry_key(self) -> str | None:
        row = self.dictionary_table.currentRow()
        if 0 <= row < len(self._entries):
            return _entry_key(self._entries[row])
        return None

    def _on_dictionary_row_changed(self, current_row: int, _current_col: int, previous_row: int, _previous_col: int) -> None:
        if 0 <= previous_row < len(self._entries):
            self._save_form_to_row(previous_row, show_errors=False)
        if 0 <= current_row < len(self._entries):
            self._load_entry_into_form(current_row)
        else:
            self._clear_form()

    def _clear_form(self) -> None:
        self._updating_form = True
        for widget in (self.technical_edit, self.display_edit, self.group_edit, self.alias_input):
            widget.clear()
        self.aliases_edit.clear()
        self.comment_edit.clear()
        self.required_check.setChecked(False)
        self.field_type_combo.setCurrentText("text")
        self.usage_text.clear()
        self._updating_form = False

    def _load_entry_into_form(self, row: int) -> None:
        entry = self._entries[row]
        self._updating_form = True
        self.technical_edit.setText(entry.technical_name)
        self.display_edit.setText(entry.display_name)
        self.group_edit.setText(entry.group)
        self.field_type_combo.setCurrentText(entry.field_type if entry.field_type in FIELD_TYPES else "text")
        self.required_check.setChecked(entry.required)
        self.aliases_edit.setPlainText("\n".join(sorted(entry.variants, key=str.lower)))
        self.comment_edit.setPlainText(entry.comment)
        self.alias_input.clear()
        usage = self._usage_by_key.get(_entry_key(entry), [])
        self.usage_text.setPlainText("\n".join(usage) if usage else "Не используется")
        self._updating_form = False

    def _form_entry(self) -> VariableEntry:
        variants = {
            line.strip()
            for line in self.aliases_edit.toPlainText().splitlines()
            if line.strip()
        }
        return VariableEntry(
            technical_name=self.technical_edit.text().strip(),
            display_name=self.display_edit.text().strip() or self.technical_edit.text().strip(),
            variants=variants,
            field_type=self.field_type_combo.currentText().strip() or "text",
            group=self.group_edit.text().strip() or "project",
            required=self.required_check.isChecked(),
            comment=self.comment_edit.toPlainText().strip(),
        )

    def _on_form_changed(self, *_args) -> None:
        if self._updating_form:
            return
        row = self.dictionary_table.currentRow()
        if 0 <= row < len(self._entries):
            self._save_form_to_row(row, show_errors=False)

    def _save_form_to_row(self, row: int, *, show_errors: bool) -> bool:
        entry = self._form_entry()
        if not entry.technical_name:
            if show_errors:
                QMessageBox.warning(self, "Словарь", "Техническое имя не может быть пустым.")
            return False
        self._entries[row] = entry
        return True

    def _add_variable(self) -> None:
        base = "new_variable"
        used = {normalize_var_name(entry.technical_name) for entry in self._entries}
        technical_name = base
        idx = 2
        while normalize_var_name(technical_name) in used:
            technical_name = f"{base}_{idx}"
            idx += 1

        entry = VariableEntry(
            technical_name=technical_name,
            display_name="Новая переменная",
            variants=set(),
            field_type="text",
            group="project",
        )
        self._entries.append(entry)
        self._render_dictionary_table(_entry_key(entry))
        self.technical_edit.setFocus()
        self.technical_edit.selectAll()

    def _delete_variable(self) -> None:
        row = self.dictionary_table.currentRow()
        if not (0 <= row < len(self._entries)):
            return
        entry = self._entries[row]
        answer = QMessageBox.question(
            self,
            "Словарь",
            f"Удалить переменную «{entry.display_name}»?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        del self._entries[row]
        self._render_dictionary_table()

    def _add_alias(self) -> None:
        alias = self.alias_input.text().strip()
        if not alias:
            return
        aliases = [
            line.strip()
            for line in self.aliases_edit.toPlainText().splitlines()
            if line.strip()
        ]
        if alias not in aliases:
            aliases.append(alias)
            self.aliases_edit.setPlainText("\n".join(aliases))
        self.alias_input.clear()
        self._on_form_changed()

    def _save_dictionary(self) -> None:
        row = self.dictionary_table.currentRow()
        if 0 <= row < len(self._entries) and not self._save_form_to_row(row, show_errors=True):
            return

        seen: dict[str, str] = {}
        for entry in self._entries:
            key = normalize_var_name(entry.technical_name)
            if not key:
                QMessageBox.warning(self, "Словарь", "В словаре есть переменная без технического имени.")
                return
            if key in seen:
                QMessageBox.warning(
                    self,
                    "Словарь",
                    f"Повторяется техническое имя: {entry.technical_name}",
                )
                return
            seen[key] = entry.technical_name

        try:
            save_entries_to_file(self._entries, user_variables_path())
            self._dict = load_variable_dictionary()
            self._reload_dictionary_usage(show_errors=False)
            select_key = normalize_var_name(self._entries[row].technical_name) if 0 <= row < len(self._entries) else None
            self._render_dictionary_table(select_key)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Словарь", f"Не удалось сохранить словарь:\n{e}")
            return

        self.dictionary_changed.emit()
        self._show_status("Словарь переменных сохранён")

    def reload_all_variables(self) -> None:
        self._reload()
        self._reload_dictionary(show_errors=False)
