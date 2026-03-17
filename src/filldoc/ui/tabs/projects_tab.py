from __future__ import annotations

import json
import re
from pathlib import Path

from PySide6.QtCore import Qt, QPoint, QByteArray, QRect, QSize
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QColor, QFont, QIcon, QPixmap, QPainter, QTextOption
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStyle,
    QStyledItemDelegate,
    QTabWidget,
    QTextEdit,
    QToolButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project

PROJECT_NAME_FIELD = "Имя проекта"

_DROP_ACTIVE_STYLE = "QTableWidget { border: 2px dashed #4A90D9; border-radius: 4px; }"

_CARD_FIXED_FIELDS = [
    "Имя проекта",
    "ИНН кредитора",
    "ИНН должника",
    "Номер дела",
    "Номер листа и дата",
    "Номер ИП",
]

_CARD_FIELD_COL_MIN_W = 150
_CARD_FIELD_COL_MAX_W = 360
_CARD_FIELD_COL_DEFAULT_W = 190
_CARD_VALUE_COL_MIN_W = 220
_CARD_COL_GAP = 0
_CARD_COL_STEP = 12

# ── SVG-иконки ──────────────────────────────────────────────────────────────

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

_SVG_ADD = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10"/>
  <line x1="12" y1="8" x2="12" y2="16"/>
  <line x1="8"  y1="12" x2="16" y2="12"/>
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

_TEXTEDIT_STYLE = """
QTextEdit {
    background-color: #ffffff;
    border: 1px solid #dde2ea;
    border-radius: 4px;
    padding: 2px 6px;
    font-size: 12px;
    color: #1e2a38;
    selection-background-color: #c8dff7;
}
QTextEdit:focus {
    border-color: #5b9bd5;
    background-color: #fafcff;
}
"""

_FIXED_VALUE_STYLE = """
QLineEdit {
    background-color: #ffffff;
    border: 1px solid #dde2ea;
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 12px;
    color: #1e2a38;
}
QLineEdit:focus {
    border-color: #5b9bd5;
    background-color: #fafcff;
}
"""

_CARD_LABEL_CSS = (
    "color: #5b6a7a; font-size: 11px; font-weight: 600; "
    "padding: 0; margin: 0;"
)

_CUSTOM_NAME_EDIT_STYLE = """
QLineEdit {
    color: #5b6a7a;
    font-size: 11px;
    font-weight: 600;
    background: #ffffff;
    border: 1px solid #dde2ea;
    border-radius: 4px;
    padding: 1px 6px;
    min-height: 22px;
}
QLineEdit:focus {
    color: #4a7ab5;
    border-color: #5b9bd5;
    background-color: #fafcff;
}
QLineEdit:hover {
    border-color: #b9c9da;
}
"""

_CARD_COL_HEADER_CSS = (
    "color: #8fa0b3; font-size: 10px; font-weight: 700; "
    "letter-spacing: 0.5px; padding: 0 0 2px 0;"
)

_CARD_DIVIDER_BTN_STYLE = """
QToolButton {
    background: transparent;
    border: none;
    color: #9fb0c2;
    font-size: 10px;
    min-width:  14px;
    min-height: 18px;
    max-width:  14px;
    max-height: 18px;
    padding: 0;
}
QToolButton:hover {
    background: #e8edf3;
    color: #5b6a7a;
    border-radius: 4px;
}
QToolButton:pressed {
    background: #dbe3ec;
}
"""

_MINI_BTN_STYLE = """
QToolButton {
    background: transparent;
    border: none;
    color: #b8c6d4;
    font-size: 13px;
    min-width:  22px;
    min-height: 22px;
    max-width:  22px;
    max-height: 22px;
    border-radius: 5px;
}
QToolButton:hover {
    background: #e8edf3;
    color: #5b6a7a;
}
QToolButton:pressed {
    background: #d0d8e2;
}
"""

_MINI_DEL_BTN_STYLE = """
QToolButton {
    background: transparent;
    border: none;
    color: #d0a8a8;
    font-size: 13px;
    min-width:  22px;
    min-height: 22px;
    max-width:  22px;
    max-height: 22px;
    border-radius: 5px;
}
QToolButton:hover {
    background: #fde8e8;
    color: #b04040;
}
QToolButton:pressed {
    background: #f5d0d0;
}
"""

_ADD_FIELD_BTN_STYLE = """
QPushButton {
    background: transparent;
    color: #5b9bd5;
    border: 1.5px dashed #9dc3e6;
    border-radius: 6px;
    font-size: 11px;
    font-weight: 600;
    padding: 4px 12px;
}
QPushButton:hover {
    background: #eef5fc;
    border-color: #5b9bd5;
    color: #3a7ec0;
}
QPushButton:pressed {
    background: #d8ecf8;
}
"""


# ── Авто-генерация имени проекта ─────────────────────────────────────────────

_LEGAL_FORM_RE = re.compile(
    r"""
    ^                                           # начало строки
    (?:ООО|АО|ПАО|ЗАО|ОАО|НАО|ИП|АНО|НКО|    # аббревиатура ОПФ
       ФГУП|МУП|ГУП|ГК|ТСЖ|СНТ|ДНТ)
    [\s\u00A0]*                                 # пробел(ы) после
    """,
    re.VERBOSE | re.IGNORECASE,
)

_QUOTES_RE = re.compile(r'^[«"\'„\u201c\u201e](.+)[»"\'"\u201d\u201c]$')


def _strip_legal_form(text: str) -> str:
    """Убирает аббревиатуру ОПФ в начале и декоративные кавычки."""
    text = text.strip()
    text = _LEGAL_FORM_RE.sub("", text).strip()
    m = _QUOTES_RE.match(text)
    if m:
        text = m.group(1).strip()
    return text


def _auto_project_name(fields: dict[str, str]) -> str:
    """Формирует 'Кредитор — Должник' без указания ОПФ."""
    creditor = _strip_legal_form(fields.get("Кредитор", "").strip())
    debtor   = _strip_legal_form(fields.get("Должник",  "").strip())
    if creditor and debtor:
        return f"{creditor} — {debtor}"
    return creditor or debtor or ""


_PROJECT_LIST_STYLE = """
QListWidget {
    background-color: #ffffff;
    border: 1px solid #dde2ea;
    border-radius: 8px;
    outline: 0;
    padding: 4px 4px;
}
QListWidget::item {
    border: none;
    background: transparent;
    padding: 0;
    margin: 0;
}
QScrollBar:vertical {
    background: #f4f6f9;
    width: 6px;
    border-radius: 3px;
    margin: 4px 2px 4px 2px;
}
QScrollBar::handle:vertical {
    background: #c0cad6;
    border-radius: 3px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover {
    background: #8A9BB0;
}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical { background: none; }
"""

# ── Вспомогательные функции ──────────────────────────────────────────────────

def _make_icon(svg_src: str, color: str = "#ffffff", size: int = 18) -> QIcon:
    svg_bytes = QByteArray(svg_src.replace("currentColor", color).encode())
    renderer = QSvgRenderer(svg_bytes)
    px = QPixmap(size, size)
    px.fill(Qt.GlobalColor.transparent)
    painter = QPainter(px)
    renderer.render(painter)
    painter.end()
    return QIcon(px)


def _icon_btn(svg: str, tooltip: str, icon_color: str, bg: str, hover: str, pressed: str) -> QToolButton:
    btn = QToolButton()
    btn.setIcon(_make_icon(svg, icon_color))
    btn.setIconSize(QSize(18, 18))
    btn.setToolTip(tooltip)
    btn.setStyleSheet(_BTN_STYLE.format(bg=bg, hover=hover, pressed=pressed))
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def _mini_btn(text: str, tooltip: str, style: str = _MINI_BTN_STYLE) -> QToolButton:
    btn = QToolButton()
    btn.setText(text)
    btn.setToolTip(tooltip)
    btn.setStyleSheet(style)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


# ── Делегат списка проектов ───────────────────────────────────────────────────

class _ProjectItemDelegate(QStyledItemDelegate):
    """Рисует каждый элемент списка проектов как современную карточку."""

    _H = 28          # высота строки
    _RADIUS = 5      # скругление фона
    _ACCENT_W = 3    # ширина левой цветной полосы при выборе

    def paint(self, painter: QPainter, option, index) -> None:  # noqa: ANN001
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        is_selected = bool(option.state & QStyle.StateFlag.State_Selected)
        is_hovered  = bool(option.state & QStyle.StateFlag.State_MouseOver)

        # Archived items have a background role set
        bg_role = index.data(Qt.ItemDataRole.BackgroundRole)
        is_archived = bg_role is not None and isinstance(bg_role, QColor)

        rect = option.rect.adjusted(2, 2, -2, -2)

        # ── Фон ──────────────────────────────────────────────────────
        if is_selected:
            bg = QColor("#deeaff")
        elif is_hovered:
            bg = QColor("#f0f5fb")
        elif is_archived:
            bg = QColor("#f5f5f5")
        else:
            bg = QColor("#ffffff")

        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(bg)
        painter.drawRoundedRect(rect, self._RADIUS, self._RADIUS)

        # ── Левая акцентная полоса (выбранный элемент) ────────────────
        if is_selected:
            accent = QRect(rect.left() + 1, rect.top() + 5,
                           self._ACCENT_W, rect.height() - 10)
            painter.setBrush(QColor("#4A90D9"))
            painter.drawRoundedRect(accent, 1, 1)

        # ── Текст ──────────────────────────────────────────────────────
        if is_selected:
            text_color = QColor("#1a3a6b")
        elif is_archived:
            text_color = QColor("#8a8a8a")
        else:
            text_color = QColor("#1e2a38")

        font = QFont(option.font)
        font.setPointSizeF(9.0)
        font.setWeight(QFont.Weight.DemiBold if is_selected else QFont.Weight.Normal)
        painter.setFont(font)
        painter.setPen(text_color)

        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        text_rect = rect.adjusted(14, 0, -8, 0)
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            text,
        )

        painter.restore()

    def sizeHint(self, option, index) -> QSize:  # noqa: ANN001
        w = option.rect.width() if option.rect.width() > 0 else 200
        return QSize(w, self._H)


# ── Авто-изменяемый QTextEdit ────────────────────────────────────────────────

class _AutoResizeTextEdit(QTextEdit):
    """QTextEdit, автоматически подстраивающий высоту под содержимое."""

    _MAX_H = 400

    def __init__(self, placeholder: str = "", parent=None) -> None:
        super().__init__(parent)
        self.setPlaceholderText(placeholder)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.setWordWrapMode(QTextOption.WrapMode.WrapAtWordBoundaryOrAnywhere)
        self.setAcceptRichText(False)
        self.setStyleSheet(_TEXTEDIT_STYLE)
        self.setContentsMargins(0, 0, 0, 0)
        self.document().setDocumentMargin(0)
        self.document().documentLayout().documentSizeChanged.connect(
            lambda _: self._adjust_height()
        )

    def _adjust_height(self) -> None:
        vp_w = self.viewport().width()
        if vp_w <= 0:
            return
        doc = self.document()
        doc.setTextWidth(float(vp_w))
        doc_h = int(doc.documentLayout().documentSize().height())
        # CSS padding (2px top + 2px bottom) + border (1px + 1px)
        h = doc_h + 6
        h = min(h, self._MAX_H)
        if self.height() != h:
            self.setFixedHeight(h)

    def resizeEvent(self, event) -> None:  # noqa: ANN001
        super().resizeEvent(event)
        self._adjust_height()

    def showEvent(self, event) -> None:  # noqa: ANN001
        super().showEvent(event)
        self._adjust_height()


# ── Главный класс вкладки ────────────────────────────────────────────────────

class ProjectsTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings()
        self._projects: list[Project] = []
        self._archived_projects: list[Project] = []
        self._current: Project | None = None
        self._showing_archive: bool = False
        self._card_custom_keys: dict[str, list[str]] = {}
        self._current_tab_index: int = 0
        self._card_field_col_width: int = _CARD_FIELD_COL_DEFAULT_W
        self._card_content_widget: QWidget | None = None

        self.setAcceptDrops(True)

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        root.addLayout(top)

        self.load_btn = _icon_btn(_SVG_REFRESH, "Обновить проекты из Excel",
                                  "#ffffff", "#8A9BB0", "#6B7F96", "#556477")
        self.save_btn = _icon_btn(_SVG_SAVE, "Сохранить изменения",
                                  "#ffffff", "#8A9BB0", "#6B7F96", "#556477")
        self.add_btn = _icon_btn(_SVG_ADD, "Добавить проект",
                                 "#ffffff", "#8A9BB0", "#6B7F96", "#556477")
        top.addWidget(self.load_btn)
        top.addWidget(self.save_btn)
        top.addWidget(self.add_btn)
        top.addStretch(1)

        hint = QLabel("Перетащите .json-файл в окно для загрузки проекта")
        hint.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        hint.setStyleSheet("color: #888; font-style: italic; font-size: 11px;")
        top.addWidget(hint)

        # ── Сплиттер (левый список + правые вкладки) ──────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(6)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        self.list = QListWidget(self)
        self.list.setMinimumWidth(180)
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list.setDragEnabled(True)
        self.list.setAcceptDrops(True)
        self.list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.list.setMouseTracking(True)
        self.list.setStyleSheet(_PROJECT_LIST_STYLE)
        self.list.setItemDelegate(_ProjectItemDelegate(self.list))
        self.list.setSpacing(1)
        splitter.addWidget(self.list)

        # ── Вкладки ────────────────────────────────────────────────────────
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_card_tab(), "Карточка")

        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Поле", "Значение"])
        h0 = self.table.horizontalHeaderItem(0)
        h1 = self.table.horizontalHeaderItem(1)
        if h0:
            h0.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        if h1:
            h1.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked
        )
        self.tabs.addTab(self.table, "Реквизиты")

        splitter.addWidget(self.tabs)
        splitter.setSizes([280, 600])
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)

        # ── Сигналы ────────────────────────────────────────────────────────
        self.list.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked
        )
        self.list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.load_btn.clicked.connect(self._load_projects)
        self.save_btn.clicked.connect(self._save_all)
        self.add_btn.clicked.connect(self._add_project)
        self.list.customContextMenuRequested.connect(self._on_list_context_menu)
        self.list.currentRowChanged.connect(self._select_project)
        self.list.itemChanged.connect(self._on_list_item_edited)
        self.tabs.currentChanged.connect(self._on_tab_changed)

    # ── Построение вкладки «Карточка» ────────────────────────────────────────

    def _build_card_tab(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("QScrollArea, QScrollArea > QWidget > QWidget { background: #f4f6f9; }")

        content = QWidget()
        self._card_content_widget = content
        self._card_content_vbox = QVBoxLayout(content)
        self._card_content_vbox.setContentsMargins(12, 8, 12, 8)
        self._card_content_vbox.setSpacing(5)
        self._card_left_col_widgets: list[QWidget] = []

        # Заголовки колонок
        cols_header = QWidget()
        cols_header.setStyleSheet("background: transparent;")
        cols_header_h = QHBoxLayout(cols_header)
        cols_header_h.setContentsMargins(0, 0, 0, 0)
        cols_header_h.setSpacing(_CARD_COL_GAP)

        left_col_header = QLabel("Поле")
        left_col_header.setStyleSheet(_CARD_COL_HEADER_CSS)
        right_col_header = QLabel("Значение")
        right_col_header.setStyleSheet(_CARD_COL_HEADER_CSS)

        divider_controls = QWidget()
        divider_controls_h = QHBoxLayout(divider_controls)
        divider_controls_h.setContentsMargins(0, 0, 0, 0)
        divider_controls_h.setSpacing(0)
        divider_controls.setStyleSheet("background: transparent;")

        divider_left_btn = _mini_btn("◀", "Сдвинуть границу влево", _CARD_DIVIDER_BTN_STYLE)
        divider_right_btn = _mini_btn("▶", "Сдвинуть границу вправо", _CARD_DIVIDER_BTN_STYLE)
        divider_left_btn.clicked.connect(lambda checked=False: self._shift_card_divider(-_CARD_COL_STEP))
        divider_right_btn.clicked.connect(lambda checked=False: self._shift_card_divider(_CARD_COL_STEP))

        divider_controls_h.addWidget(divider_left_btn)
        divider_controls_h.addWidget(divider_right_btn)

        cols_header_h.addWidget(left_col_header)
        cols_header_h.addWidget(divider_controls)
        cols_header_h.addWidget(right_col_header, 1)
        self._card_content_vbox.addWidget(cols_header)
        self._register_card_left_widget(left_col_header)

        # Фиксированные поля
        self._card_fixed_edits: dict[str, QLineEdit] = {}
        for field_name in _CARD_FIXED_FIELDS:
            fw, edit = self._make_fixed_field_row(field_name)
            self._card_fixed_edits[field_name] = edit
            self._card_content_vbox.addWidget(fw)

        # Разделитель
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #dde2ea; margin: 2px 0;")
        self._card_content_vbox.addWidget(sep)

        # Заголовок секции доп. полей
        extras_header = QHBoxLayout()
        extras_label = QLabel("Дополнительные поля")
        extras_label.setStyleSheet("color: #6b7a8d; font-size: 10px; font-weight: 700;")
        extras_header.addWidget(extras_label)
        extras_header.addStretch()
        self._card_content_vbox.addLayout(extras_header)

        # Контейнер для доп. полей
        self._card_extras_container = QWidget()
        self._card_extras_container.setStyleSheet("background: transparent;")
        self._card_extras_vbox = QVBoxLayout(self._card_extras_container)
        self._card_extras_vbox.setContentsMargins(0, 0, 0, 0)
        self._card_extras_vbox.setSpacing(2)
        self._card_extras_vbox.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._card_content_vbox.addWidget(self._card_extras_container)

        # Кнопка добавления поля
        add_btn_row = QHBoxLayout()
        self._card_add_field_btn = QPushButton("+ Добавить поле")
        self._card_add_field_btn.setStyleSheet(_ADD_FIELD_BTN_STYLE)
        self._card_add_field_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._card_add_field_btn.clicked.connect(self._add_card_field)
        add_btn_row.addWidget(self._card_add_field_btn)
        add_btn_row.addStretch()
        self._card_content_vbox.addLayout(add_btn_row)

        self._card_content_vbox.addStretch()

        scroll.setWidget(content)
        outer.addWidget(scroll)

        self._card_custom_rows: list[tuple[QWidget, QLineEdit, _AutoResizeTextEdit]] = []
        self._set_card_field_col_width(self._card_field_col_width)
        return wrapper

    def _make_fixed_field_row(self, field_name: str) -> tuple[QWidget, QLineEdit]:
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        hbox = QHBoxLayout(row)
        hbox.setContentsMargins(0, 0, 0, 0)
        hbox.setSpacing(_CARD_COL_GAP)

        label = QLabel(field_name)
        label.setStyleSheet(_CARD_LABEL_CSS)
        self._register_card_left_widget(label)

        edit = QLineEdit()
        edit.setPlaceholderText("—")
        edit.setStyleSheet(_FIXED_VALUE_STYLE)
        edit.setFixedHeight(22)

        hbox.addWidget(label)
        hbox.addWidget(edit, 1)
        return row, edit

    def _make_custom_field_row(
        self, name: str = "", value: str = ""
    ) -> tuple[QWidget, QLineEdit, _AutoResizeTextEdit]:
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        row.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        hbox = QHBoxLayout(row)
        hbox.setContentsMargins(0, 0, 0, 0)
        hbox.setSpacing(_CARD_COL_GAP)
        hbox.setAlignment(Qt.AlignmentFlag.AlignTop)

        field_cell = QWidget()
        field_cell_h = QHBoxLayout(field_cell)
        field_cell_h.setSpacing(1)
        field_cell_h.setContentsMargins(0, 0, 0, 0)
        field_cell_h.setAlignment(Qt.AlignmentFlag.AlignTop)
        self._register_card_left_widget(field_cell)
        row.setProperty("_left_col_widget", field_cell)

        name_edit = QLineEdit(name)
        name_edit.setPlaceholderText("Название поля")
        name_edit.setStyleSheet(_CUSTOM_NAME_EDIT_STYLE)

        up_btn = _mini_btn("↑", "Переместить вверх")
        down_btn = _mini_btn("↓", "Переместить вниз")
        del_btn = _mini_btn("×", "Удалить поле", _MINI_DEL_BTN_STYLE)

        field_cell_h.addWidget(name_edit, 1)

        value_edit = _AutoResizeTextEdit(placeholder="—")
        if value:
            value_edit.setPlainText(value)

        controls_cell = QWidget()
        controls_h = QHBoxLayout(controls_cell)
        controls_h.setContentsMargins(1, 0, 0, 0)
        controls_h.setSpacing(1)
        controls_h.setAlignment(Qt.AlignmentFlag.AlignTop)
        controls_h.addWidget(up_btn)
        controls_h.addWidget(down_btn)
        controls_h.addWidget(del_btn)

        hbox.addWidget(field_cell)
        hbox.addWidget(value_edit, 1)
        hbox.addWidget(controls_cell)

        up_btn.clicked.connect(lambda checked=False, r=row: self._card_field_move_up(r))
        down_btn.clicked.connect(lambda checked=False, r=row: self._card_field_move_down(r))
        del_btn.clicked.connect(lambda checked=False, r=row: self._card_field_delete(r))

        return row, name_edit, value_edit

    # ── Карточка: рендер и чтение ─────────────────────────────────────────────

    def _render_card(self, project: Project) -> None:
        for field_name, edit in self._card_fixed_edits.items():
            edit.blockSignals(True)
            edit.setText(project.fields.get(field_name, ""))
            edit.blockSignals(False)

        # Очищаем старые доп. поля
        for row_widget, _, _ in self._card_custom_rows:
            self._card_extras_vbox.removeWidget(row_widget)
            left_col_widget = row_widget.property("_left_col_widget")
            if isinstance(left_col_widget, QWidget):
                self._unregister_card_left_widget(left_col_widget)
            row_widget.setParent(None)  # type: ignore[arg-type]
            row_widget.deleteLater()
        self._card_custom_rows.clear()

        # Создаём новые
        for key in self._card_custom_keys.get(project.project_id, []):
            rw, ne, ve = self._make_custom_field_row(name=key, value=project.fields.get(key, ""))
            self._card_custom_rows.append((rw, ne, ve))
            self._card_extras_vbox.addWidget(rw)

    def _read_card_into_project(self, project: Project) -> None:
        for field_name, edit in self._card_fixed_edits.items():
            project.fields[field_name] = edit.text().strip()

        custom_keys: list[str] = []
        for _, name_edit, value_edit in self._card_custom_rows:
            k = name_edit.text().strip()
            v = value_edit.toPlainText().strip()
            if k:
                project.fields[k] = v
                custom_keys.append(k)
        self._card_custom_keys[project.project_id] = custom_keys

        row = self.list.currentRow()
        item = self.list.item(row)
        if item is not None:
            self.list.itemChanged.disconnect(self._on_list_item_edited)
            item.setText(self._project_display_name(project))
            self.list.itemChanged.connect(self._on_list_item_edited)

    def _clear_card_display(self) -> None:
        for edit in self._card_fixed_edits.values():
            edit.blockSignals(True)
            edit.setText("")
            edit.blockSignals(False)
        for row_widget, _, _ in self._card_custom_rows:
            self._card_extras_vbox.removeWidget(row_widget)
            left_col_widget = row_widget.property("_left_col_widget")
            if isinstance(left_col_widget, QWidget):
                self._unregister_card_left_widget(left_col_widget)
            row_widget.setParent(None)  # type: ignore[arg-type]
            row_widget.deleteLater()
        self._card_custom_rows.clear()

    # ── Карточка: управление доп. полями ──────────────────────────────────────

    def _add_card_field(self) -> None:
        if self._current is None:
            return
        rw, ne, ve = self._make_custom_field_row()
        self._card_custom_rows.append((rw, ne, ve))
        self._card_extras_vbox.addWidget(rw)
        ne.setFocus()

    def _find_custom_row_index(self, row_widget: QWidget) -> int:
        for i, (w, _, _) in enumerate(self._card_custom_rows):
            if w is row_widget:
                return i
        return -1

    def _swap_custom_row_contents(self, i: int, j: int) -> None:
        _, ne_i, ve_i = self._card_custom_rows[i]
        _, ne_j, ve_j = self._card_custom_rows[j]
        n_i, v_i = ne_i.text(), ve_i.toPlainText()
        n_j, v_j = ne_j.text(), ve_j.toPlainText()
        ne_i.setText(n_j)
        ve_i.setPlainText(v_j)
        ne_j.setText(n_i)
        ve_j.setPlainText(v_i)

    def _card_field_move_up(self, row_widget: QWidget) -> None:
        idx = self._find_custom_row_index(row_widget)
        if idx <= 0:
            return
        self._swap_custom_row_contents(idx, idx - 1)

    def _card_field_move_down(self, row_widget: QWidget) -> None:
        idx = self._find_custom_row_index(row_widget)
        if idx < 0 or idx >= len(self._card_custom_rows) - 1:
            return
        self._swap_custom_row_contents(idx, idx + 1)

    def _card_field_delete(self, row_widget: QWidget) -> None:
        idx = self._find_custom_row_index(row_widget)
        if idx < 0:
            return
        self._card_custom_rows.pop(idx)
        self._card_extras_vbox.removeWidget(row_widget)
        left_col_widget = row_widget.property("_left_col_widget")
        if isinstance(left_col_widget, QWidget):
            self._unregister_card_left_widget(left_col_widget)
        row_widget.setParent(None)  # type: ignore[arg-type]
        row_widget.deleteLater()

    def _register_card_left_widget(self, widget: QWidget) -> None:
        self._card_left_col_widgets.append(widget)
        widget.setFixedWidth(self._card_field_col_width)

    def _unregister_card_left_widget(self, widget: QWidget) -> None:
        if widget in self._card_left_col_widgets:
            self._card_left_col_widgets.remove(widget)

    def _set_card_field_col_width(self, width: int) -> None:
        width = max(_CARD_FIELD_COL_MIN_W, min(_CARD_FIELD_COL_MAX_W, int(width)))
        if self._card_content_widget is not None:
            available = self._card_content_widget.width()
            max_by_window = available - _CARD_VALUE_COL_MIN_W - 28
            width = min(width, max(_CARD_FIELD_COL_MIN_W, max_by_window))
        self._card_field_col_width = width
        for widget in self._card_left_col_widgets:
            widget.setFixedWidth(width)

    def _shift_card_divider(self, delta: int) -> None:
        self._set_card_field_col_width(self._card_field_col_width + delta)

    def resizeEvent(self, event) -> None:  # noqa: ANN001
        super().resizeEvent(event)
        self._set_card_field_col_width(self._card_field_col_width)

    # ── Синхронизация вкладок ─────────────────────────────────────────────────

    def _on_tab_changed(self, new_index: int) -> None:
        if self._current is None:
            self._current_tab_index = new_index
            return
        old_index = self._current_tab_index
        self._current_tab_index = new_index
        if old_index == 0:
            self._read_card_into_project(self._current)
        else:
            self._read_table_into_project(self._current)
        if new_index == 0:
            self._render_card(self._current)
        else:
            self._render_project(self._current)

    def _sync_current_to_project(self) -> None:
        if self._current is None:
            return
        if self._current_tab_index == 0:
            self._read_card_into_project(self._current)
        else:
            self._read_table_into_project(self._current)

    # ── Drag & drop ───────────────────────────────────────────────────────────

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls() and any(
            u.toLocalFile().lower().endswith(".json") for u in event.mimeData().urls()
        ):
            self.table.setStyleSheet(_DROP_ACTIVE_STYLE)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: ANN001
        self.table.setStyleSheet("")
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self.table.setStyleSheet("")
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".json"):
                self._load_from_json(path)
                event.acceptProposedAction()
                return
        event.ignore()

    def _load_from_json(self, path: str) -> None:
        try:
            with open(path, encoding="utf-8") as f:
                data: dict = json.load(f)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "JSON", f"Не удалось прочитать файл:\n{e}")
            return

        if not isinstance(data, dict):
            QMessageBox.warning(self, "JSON", "Файл должен содержать JSON-объект (словарь полей).")
            return

        case_number = str(data.get("№ дела", "") or data.get("Номер дела", "")).strip()
        project_id = case_number or Path(path).stem
        fields = {str(k): str(v) for k, v in data.items()}
        project = Project(project_id=project_id, fields=fields, headers=list(fields.keys()))

        self._projects.append(project)
        self._add_project_to_list(project)
        self.list.setCurrentRow(len(self._projects) - 1)

    # ── Список проектов ───────────────────────────────────────────────────────

    def _add_project_to_list(self, project: Project, *, archived: bool = False) -> None:
        item = QListWidgetItem(self._project_display_name(project))
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        item.setData(Qt.ItemDataRole.UserRole, project)
        if archived:
            item.setForeground(QColor("#777777"))
            item.setBackground(QColor("#f2f2f2"))
        self.list.addItem(item)

    def _project_display_name(self, project: Project) -> str:
        name = (project.fields.get(PROJECT_NAME_FIELD) or "").strip()
        if name:
            return name
        auto = _auto_project_name(project.fields)
        return auto or project.project_id

    def _refresh_project_name(self, project: Project) -> None:
        """Пересчитывает 'Имя проекта' из Кредитор и Должник и сохраняет в поля."""
        auto = _auto_project_name(project.fields)
        if auto:
            project.fields[PROJECT_NAME_FIELD] = auto

    def _on_list_item_edited(self, item: QListWidgetItem) -> None:
        project = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(project, Project):
            project.fields[PROJECT_NAME_FIELD] = (item.text() or "").strip()

    def _on_list_context_menu(self, pos: QPoint) -> None:
        item = self.list.itemAt(pos)
        from PySide6.QtWidgets import QMenu
        menu = QMenu(self)
        if item is None:
            if self._showing_archive:
                show_current_action = menu.addAction("Показать текущие проекты")
            else:
                show_archive_action = menu.addAction("Показать архивные проекты")
            global_pos = self.list.mapToGlobal(pos)
            chosen_action = menu.exec(global_pos)
            if not chosen_action:
                return
            if not self._showing_archive and chosen_action == show_archive_action:
                self._show_archived_projects()
            elif self._showing_archive and chosen_action == show_current_action:
                self._show_current_projects()
            return

        if self._showing_archive:
            unarchive_action = menu.addAction("Убрать из архива")
        else:
            archive_action = menu.addAction("В архив")
            delete_action = menu.addAction("Удалить")

        global_pos = self.list.mapToGlobal(pos)
        chosen_action = menu.exec(global_pos)
        if not chosen_action:
            return

        row = self.list.row(item)
        if row < 0:
            return
        self.list.setCurrentRow(row)

        if not self._showing_archive and chosen_action == archive_action:
            self._archive_current()
        elif not self._showing_archive and chosen_action == delete_action:
            self._delete_current()
        elif self._showing_archive and chosen_action == unarchive_action:
            self._unarchive_current()

    def _show_current_projects(self) -> None:
        self._showing_archive = False
        self.list.clear()
        for p in self._projects:
            self._add_project_to_list(p, archived=False)
        if self._projects:
            self.list.setCurrentRow(0)
        else:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()

    def _show_archived_projects(self) -> None:
        self._showing_archive = True
        self.list.clear()
        for p in self._archived_projects:
            self._add_project_to_list(p, archived=True)
        if self._archived_projects:
            self.list.setCurrentRow(0)
        else:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()

    def _archive_current(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Архив", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        row = self.list.currentRow()
        if row < 0 or row >= self.list.count():
            QMessageBox.warning(self, "Архив", "Не выбран проект для архивации.")
            return
        item = self.list.item(row)
        if item is None:
            QMessageBox.warning(self, "Архив", "Не выбран проект для архивации.")
            return
        project = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(project, Project):
            QMessageBox.warning(self, "Архив", "Не удалось определить выбранный проект.")
            return

        title = self._project_display_name(project)
        answer = QMessageBox.question(
            self, "Архивировать проект", f"Переместить проект в архив:\n{title}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        try:
            store = ExcelProjectStore(self._settings.excel_path)
            store.move_project_to_archive(project)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Архив", f"Не удалось архивировать проект:\n{e}")
            return
        self._reload_from_excel(keep_mode="current")

    def _unarchive_current(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Архив", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        row = self.list.currentRow()
        if row < 0 or row >= self.list.count():
            QMessageBox.warning(self, "Архив", "Не выбран проект для восстановления.")
            return
        item = self.list.item(row)
        if item is None:
            QMessageBox.warning(self, "Архив", "Не выбран проект для восстановления.")
            return
        project = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(project, Project):
            QMessageBox.warning(self, "Архив", "Не удалось определить выбранный проект.")
            return

        title = self._project_display_name(project)
        answer = QMessageBox.question(
            self, "Вернуть из архива", f"Вернуть проект из архива в текущие:\n{title}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        try:
            store = ExcelProjectStore(self._settings.excel_path)
            store.restore_project_from_archive(project)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Архив", f"Не удалось вернуть проект из архива:\n{e}")
            return
        self._reload_from_excel(keep_mode="archive")

    def _reload_from_excel(self, *, keep_mode: str = "current") -> None:
        if not self._settings.excel_path:
            return
        try:
            store = ExcelProjectStore(self._settings.excel_path)
            self._projects = store.load_projects()
            self._archived_projects = store.load_projects_from_sheet("Архив")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))
            return
        if keep_mode == "archive":
            self._show_archived_projects()
        else:
            self._show_current_projects()

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    def _load_projects(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Проекты", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        try:
            store = ExcelProjectStore(self._settings.excel_path)
            try:
                store.repair_archive_headers()
            except Exception:
                pass

            self._projects = store.load_projects()
            try:
                self._archived_projects = store.load_projects_from_sheet("Архив")
            except Exception:
                self._archived_projects = []

            for p in self._projects:
                self._refresh_project_name(p)

            self.list.clear()
            for p in self._projects:
                self._add_project_to_list(p, archived=False)
            self._showing_archive = False
            if self._projects:
                self.list.setCurrentRow(0)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))

    def _select_project(self, row: int) -> None:
        if row < 0 or row >= self.list.count():
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()
            return
        item = self.list.item(row)
        if item is None:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()
            return
        project = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(project, Project):
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()
            return
        self._current = project
        self._refresh_project_name(project)
        self._render_project(self._current)
        self._render_card(self._current)

    def _render_project(self, project: Project) -> None:
        headers = getattr(project, "headers", None)
        items: list[tuple[str, str]] = []
        if headers:
            for h in headers:
                if not h or h == PROJECT_NAME_FIELD:
                    continue
                items.append((h, project.fields.get(h, "")))
        else:
            items = [(k, v) for k, v in project.fields.items() if k != PROJECT_NAME_FIELD]
        self.table.setRowCount(len(items))
        for i, (k, v) in enumerate(items):
            key_item = QTableWidgetItem(str(k))
            key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, key_item)
            value_item = QTableWidgetItem(str(v))
            if k in {"Кредитор", "Должник"} and str(v).strip() == "":
                value_item.setBackground(QColor("#ffd6e7"))
            self.table.setItem(i, 1, value_item)
        self.table.resizeColumnsToContents()

    def _read_table_into_project(self, project: Project) -> None:
        fields: dict[str, str] = {}
        for r in range(self.table.rowCount()):
            k_item = self.table.item(r, 0)
            v_item = self.table.item(r, 1)
            k = (k_item.text() if k_item else "").strip()
            v = (v_item.text() if v_item else "").strip()
            if k:
                fields[k] = v
        if PROJECT_NAME_FIELD in project.fields:
            fields[PROJECT_NAME_FIELD] = project.fields[PROJECT_NAME_FIELD]
        project.fields = fields

    def _add_project(self) -> None:
        next_index = len(self._projects) + 1
        project = Project(
            project_id=f"Новый проект {next_index}",
            fields={},
            headers=[h for h in (self._projects[0].headers or [])]
            if self._projects and self._projects[0].headers else None,
        )
        self._projects.append(project)
        self._add_project_to_list(project, archived=False)
        self.list.setCurrentRow(len(self._projects) - 1)

    def _delete_current(self) -> None:
        row = self.list.currentRow()
        if row < 0 or row >= self.list.count():
            QMessageBox.warning(self, "Проекты", "Не выбран проект для удаления.")
            return
        item = self.list.item(row)
        if item is None:
            QMessageBox.warning(self, "Проекты", "Не выбран проект для удаления.")
            return
        project = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(project, Project):
            QMessageBox.warning(self, "Проекты", "Не удалось определить выбранный проект.")
            return
        title = self._project_display_name(project)
        answer = QMessageBox.question(
            self, "Удаление проекта", f"Удалить проект:\n{title}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        if self._settings.excel_path and project.row_index is not None:
            try:
                store = ExcelProjectStore(self._settings.excel_path)
                store.delete_project(project)
            except Exception as e:  # noqa: BLE001
                QMessageBox.critical(self, "Проекты", f"Не удалось удалить проект из Excel:\n{e}")
                return

        if project in self._projects:
            self._projects.remove(project)
        self._card_custom_keys.pop(project.project_id, None)
        self.list.takeItem(row)

        if self._projects:
            self.list.setCurrentRow(min(row, len(self._projects) - 1))
        else:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()

    def _save_all(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Проекты", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        if not self._projects:
            QMessageBox.information(self, "Проекты", "Нет проектов для сохранения.")
            return

        try:
            if self._current is not None:
                self._sync_current_to_project()

            # Пересчитываем имена перед сохранением
            for p in self._projects:
                self._refresh_project_name(p)

            ordered_projects: list[Project] = []
            for i in range(self.list.count()):
                item = self.list.item(i)
                if item is None:
                    continue
                p = item.data(Qt.ItemDataRole.UserRole)
                if isinstance(p, Project) and p not in ordered_projects:
                    ordered_projects.append(p)
            for p in self._projects:
                if p not in ordered_projects:
                    ordered_projects.append(p)
            self._projects = ordered_projects

            store = ExcelProjectStore(self._settings.excel_path)
            store.save_all_projects(self._projects, self._archived_projects)
            QMessageBox.information(
                self, "Проекты",
                "Все изменения синхронизированы с Excel (с созданием резервной копии).",
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))
