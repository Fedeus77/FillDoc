from __future__ import annotations

import json
import os
import re
import shutil
import webbrowser
from pathlib import Path

from PySide6.QtCore import Qt, QPoint, QEvent, QRect, QSize, QTimer, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QColor, QFont, QFontMetrics, QIcon, QImage, QKeySequence, QPen, QPixmap, QPainter, QShortcut, QTextOption
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
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

try:
    import fitz  # PyMuPDF
    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project
from filldoc.ui.icons import make_icon as _icons_make_icon, SVG_REFRESH, SVG_SAVE, SVG_ADD, SVG_FOLDER_OPEN, SVG_UPLOAD, SVG_LINK, update_icon_btn
from filldoc.ui.theme import ThemeColors, ThemeManager

PROJECT_NAME_FIELD = "Имя проекта"
PROJECT_TYPE_FIELD = "Тип проекта"
PROJECT_COMMENT_FIELD = "Комментарий"

_DROP_ACTIVE_STYLE = "QTableWidget { border: 2px dashed #4A90D9; border-radius: 4px; }"

# ── Typography: project names ────────────────────────────────────────────────
# A modern, even, Windows-friendly sans stack with tighter width.
_PROJECT_NAME_FONT_FAMILIES: list[str] = [
    "Segoe UI Variable Text",
    "Segoe UI Variable Display",
    "Segoe UI",
    "Inter",
    "Bahnschrift",
]
_PROJECT_NAME_FONT_PT = 10.2 * 0.9  # -10%
_PROJECT_NAME_FONT_STRETCH = 95     # narrower width (100 = normal)
_PROJECT_NAME_LETTER_SPACING_PCT = 100.0


def _project_name_font(*, base: QFont | None = None, is_selected: bool = False) -> QFont:
    font = QFont(base) if base is not None else QFont()
    # Prefer a modern, even UI font if available; Qt will pick the first installed.
    if hasattr(font, "setFamilies"):
        font.setFamilies(_PROJECT_NAME_FONT_FAMILIES)
    else:
        font.setFamily(_PROJECT_NAME_FONT_FAMILIES[0])
    font.setPointSizeF(_PROJECT_NAME_FONT_PT)
    font.setStretch(_PROJECT_NAME_FONT_STRETCH)
    font.setWeight(QFont.Weight.DemiBold if is_selected else QFont.Weight.Medium)
    font.setLetterSpacing(QFont.SpacingType.PercentageSpacing, _PROJECT_NAME_LETTER_SPACING_PCT)
    return font

_CARD_FIXED_FIELDS = [
    PROJECT_TYPE_FIELD,
    "ИНН кредитора",
    "ИНН должника",
    "Номер осн. дела",
    "Номер банк. дела",
    "Номер листа",
    "Дата листа",
    "Номер ИП",
    "Дата ИП",
    "Сумма долга всего",
    "Сумма осн. долга",
    PROJECT_COMMENT_FIELD,
]

# Поля карточки, у которых есть отдельное поле для хранения ссылки.
_CARD_LINK_FIELDS: dict[str, str] = {
    "ИНН кредитора":   "Ссылка контур кредитор",
    "ИНН должника":    "Ссылка контур должник",
    "Номер осн. дела": "Ссылка осн. дело КАД",
    "Номер банк. дела": "Ссылка банк. Дело КАД",
    "Номер ИП":        "Ссылка ИП",
}

_CASE_NUMBER_FIELD_NEW = "Номер осн. дела"
_CASE_NUMBER_FIELD_OLD = "Номер дела"

_CARD_FIELD_COL_MIN_W = 120
_CARD_FIELD_COL_MAX_W = 360
_CARD_FIELD_COL_DEFAULT_W = 140
_CARD_VALUE_COL_MIN_W = 220
_CARD_COL_GAP = 0
_CARD_COL_STEP = 12

# ── SVG-иконки (алиасы из общего модуля) ─────────────────────────────────────

_SVG_REFRESH = SVG_REFRESH
_SVG_SAVE    = SVG_SAVE
_SVG_ADD     = SVG_ADD
_SVG_FOLDER  = SVG_FOLDER_OPEN
_SVG_UPLOAD  = SVG_UPLOAD
_SVG_LINK    = SVG_LINK

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

_CARD_TITLE_EDIT_STYLE = """
QLineEdit {
    background: transparent;
    border: none;
    border-bottom: 1.5px solid #dde2ea;
    border-radius: 0;
    padding: 2px 2px 4px 2px;
    font-size: 14px;
    font-weight: 700;
    color: #1e2a38;
}
QLineEdit:focus {
    border-bottom: 2px solid #5b9bd5;
}
QLineEdit:hover {
    border-bottom-color: #aabdd0;
}
"""

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

_LINK_BTN_STYLE = """
QToolButton {
    background: transparent;
    border: none;
    color: #4a7ab5;
    border-radius: 5px;
    min-width:  22px;
    min-height: 22px;
    max-width:  22px;
    max-height: 22px;
}
QToolButton:hover {
    background: #e8edf3;
}
QToolButton:pressed {
    background: #d0d8e2;
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
        return f"{creditor} - {debtor}"
    return creditor or debtor or ""


_PROJECT_LIST_STYLE = """
QListWidget {
    background-color: #fbfcfe;
    border: 1px solid #d8e0ea;
    border-radius: 12px;
    outline: 0;
    padding: 8px 7px;
}
QListWidget::item {
    border: none;
    background: transparent;
    padding: 0;
    margin: 0 0 4px 0;
}
QScrollBar:vertical {
    background: #eef2f7;
    width: 6px;
    border-radius: 3px;
    margin: 6px 2px 6px 2px;
}
QScrollBar::handle:vertical {
    background: #bcc7d4;
    border-radius: 3px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover {
    background: #8a9bb0;
}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical { background: none; }
"""

_REQUISITES_TABLE_STYLE = """
QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #f7faff;
    border: 1px solid #dbe3ec;
    border-radius: 10px;
    gridline-color: #e7edf4;
    outline: 0;
    color: #1e2a38;
    selection-background-color: #deebff;
    selection-color: #18345b;
}
QTableWidget::item {
    padding: 6px 8px;
    border: none;
}
QHeaderView::section {
    background-color: #eef3f8;
    color: #5b6a7a;
    border: none;
    border-bottom: 1px solid #dbe3ec;
    padding: 7px 8px;
    font-size: 11px;
    font-weight: 600;
}
QScrollBar:vertical {
    background: #eef2f7;
    width: 6px;
    border-radius: 3px;
    margin: 6px 2px;
}
QScrollBar::handle:vertical {
    background: #bcc7d4;
    border-radius: 3px;
    min-height: 24px;
}
QScrollBar::handle:vertical:hover {
    background: #8a9bb0;
}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical { height: 0px; }
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical { background: none; }
"""

# ── Вспомогательные функции ──────────────────────────────────────────────────

def _make_icon(svg_src: str, color: str = "#ffffff", size: int = 18) -> QIcon:
    return _icons_make_icon(svg_src, color, size)


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



def _link_btn() -> QToolButton:
    btn = QToolButton()
    btn.setIcon(_make_icon(_SVG_LINK, "#4a7ab5", 14))
    btn.setIconSize(QSize(14, 14))
    btn.setToolTip("Открыть ссылку")
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setStyleSheet(_LINK_BTN_STYLE)
    btn.setVisible(False)
    return btn


# ── Делегат списка проектов ───────────────────────────────────────────────────

class _ProjectItemDelegate(QStyledItemDelegate):
    """Рисует каждый элемент списка проектов как современную карточку."""

    _H = 40          # высота строки
    _RADIUS = 10     # скругление фона
    _ACCENT_W = 4    # ширина левой цветной полосы при выборе

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._c: ThemeColors = ThemeManager.instance().colors

    def set_colors(self, c: ThemeColors) -> None:
        self._c = c

    def paint(self, painter: QPainter, option, index) -> None:  # noqa: ANN001
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        is_selected = bool(option.state & QStyle.StateFlag.State_Selected)
        is_hovered  = bool(option.state & QStyle.StateFlag.State_MouseOver)

        # Archived items have a background role set
        bg_role = index.data(Qt.ItemDataRole.BackgroundRole)
        is_archived = bg_role is not None and isinstance(bg_role, QColor)

        rect = option.rect.adjusted(2, 2, -2, -4)
        c = self._c

        # ── Фон ──────────────────────────────────────────────────────
        if is_selected:
            bg = QColor(c.bg_card_selected)
            border = QColor(c.border_card_selected)
        elif is_hovered:
            bg = QColor(c.bg_card_hover)
            border = QColor(c.border_card)
        elif is_archived:
            bg = QColor(c.bg_card_archived)
            border = QColor(c.border_light)
        else:
            bg = QColor(c.bg_card)
            border = QColor(c.border_card)

        painter.setPen(border)
        painter.setBrush(bg)
        painter.drawRoundedRect(rect, self._RADIUS, self._RADIUS)

        # ── Левая акцентная полоса (выбранный элемент) ────────────────
        if is_selected:
            accent = QRect(rect.left() + 3, rect.top() + 7,
                           self._ACCENT_W, rect.height() - 14)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(c.accent_stripe))
            painter.drawRoundedRect(accent, 2, 2)

        # ── Текст ──────────────────────────────────────────────────────
        if is_selected:
            text_color = QColor(c.selection_text)
        elif is_archived:
            text_color = QColor(c.text_muted)
        else:
            text_color = QColor(c.text_primary)

        font = _project_name_font(base=option.font, is_selected=is_selected)
        painter.setFont(font)
        painter.setPen(text_color)

        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        text_rect = rect.adjusted(18, 0, -12, 0)
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            text,
        )

        # Тонкая зачёркивающая линия для архивных элементов
        if is_archived and text:
            fm = QFontMetrics(font)
            text_w = min(fm.horizontalAdvance(text), text_rect.width())
            mid_y = text_rect.top() + text_rect.height() // 2
            strike_color = QColor(c.text_muted) if not is_selected else QColor(c.text_secondary)
            painter.setPen(QPen(strike_color, 1))
            painter.drawLine(text_rect.left(), mid_y, text_rect.left() + text_w, mid_y)

        painter.restore()

    def sizeHint(self, option, index) -> QSize:  # noqa: ANN001
        w = option.rect.width() if option.rect.width() > 0 else 200
        return QSize(w, self._H)


class _NameOverlay(QWidget):
    """Плавающая карточка-продолжение: появляется поверх элемента списка и показывает
    полное имя проекта как визуальное продолжение усечённого текста."""

    # Отступы текста должны совпадать с _ProjectItemDelegate
    _PAD_L = 20
    _PAD_R = 16

    def __init__(self) -> None:
        super().__init__(
            None,
            Qt.WindowType.ToolTip | Qt.WindowType.FramelessWindowHint,
        )
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

        self._label = QLabel()
        self._label.setFont(_project_name_font(is_selected=False))

        lay = QHBoxLayout(self)
        lay.setContentsMargins(self._PAD_L, 0, self._PAD_R, 0)
        lay.addWidget(self._label)

    def show_for(self, text: str, item_global_rect: QRect, *, is_archived: bool) -> None:
        """Позиционирует и показывает оверлей поверх элемента, если текст не помещается."""
        c = ThemeManager.instance().colors
        if is_archived:
            bg = c.bg_card_archived
            border = c.border_light
            fg = c.text_muted
        else:
            bg = c.bg_card
            border = c.border_card_selected
            fg = c.text_primary

        self._label.setStyleSheet(f"color: {fg}; background: transparent;")
        self._label.setText(text)
        self.setStyleSheet(
            f"QWidget {{ background-color: {bg}; border: 1px solid {border};"
            f" border-radius: 10px; }}"
        )

        self._label.adjustSize()
        needed_w = self._label.width() + self._PAD_L + self._PAD_R + 4

        # Если текст и так помещается — не показываем
        if needed_w <= item_global_rect.width():
            self.hide()
            return

        x = item_global_rect.x() + 2
        y = item_global_rect.y() + 2
        h = item_global_rect.height() - 6

        # Не выходить за правый край экрана
        screen = self.screen()
        if screen:
            max_x = screen.geometry().right() - 8
            needed_w = min(needed_w, max_x - x)

        self.setGeometry(x, y, needed_w, h)
        self.show()
        self.raise_()


class _ProjectListWidget(QListWidget):
    """Список проектов с reorder и дропом JSON на конкретную строку."""

    json_dropped = Signal(str, int)

    # Параметры шрифта делегата (для измерения ширины текста)
    _FONT_FAMILIES = _PROJECT_NAME_FONT_FAMILIES
    _FONT_PT = _PROJECT_NAME_FONT_PT
    _TEXT_MARGINS = 34  # 20 left + 14 right (в делегате)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._overlay = _NameOverlay()
        self._pending_item: QListWidgetItem | None = None
        self._hover_timer = QTimer(self)
        self._hover_timer.setSingleShot(True)
        self._hover_timer.setInterval(100)
        self._hover_timer.timeout.connect(self._commit_overlay)

    def _text_fits(self, item: QListWidgetItem) -> bool:
        font = _project_name_font(is_selected=False)
        available = self.viewport().width() - self._TEXT_MARGINS
        return QFontMetrics(font).horizontalAdvance(item.text()) <= available

    def _commit_overlay(self) -> None:
        item = self._pending_item
        if item is None or self._text_fits(item):
            self._overlay.hide()
            return
        item_rect = self.visualItemRect(item)
        tl = self.viewport().mapToGlobal(item_rect.topLeft())
        bg = item.data(Qt.ItemDataRole.BackgroundRole)
        is_archived = bg is not None and isinstance(bg, QColor)
        self._overlay.show_for(
            item.text(), QRect(tl, item_rect.size()), is_archived=is_archived
        )

    def mouseMoveEvent(self, event) -> None:  # noqa: ANN001
        item = self.itemAt(event.position().toPoint())
        if item is not self._pending_item:
            self._pending_item = item
            self._overlay.hide()
            self._hover_timer.stop()
            if item is not None:
                self._hover_timer.start()
        super().mouseMoveEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: ANN001
        self._pending_item = None
        self._hover_timer.stop()
        self._overlay.hide()
        super().leaveEvent(event)

    def hideEvent(self, event) -> None:  # noqa: ANN001
        self._overlay.hide()
        super().hideEvent(event)

    def scrollContentsBy(self, dx: int, dy: int) -> None:
        self._overlay.hide()
        super().scrollContentsBy(dx, dy)

    @staticmethod
    def _json_paths(event) -> list[str]:  # noqa: ANN001
        if not event.mimeData().hasUrls():
            return []
        paths: list[str] = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".json"):
                paths.append(path)
        return paths

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if self._json_paths(event):
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # noqa: ANN001
        if self._json_paths(event):
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        json_paths = self._json_paths(event)
        if json_paths:
            item = self.itemAt(event.position().toPoint())
            row = self.row(item) if item is not None else -1
            self.json_dropped.emit(json_paths[0], row)
            event.acceptProposedAction()
            return
        super().dropEvent(event)


# ── Строка папки документов с hover-reveal ───────────────────────────────────

class _DocsFolderRow(QWidget):
    """Строка выбора папки — поле пути и кнопка скрыты до наведения курсора."""

    def __init__(
        self,
        path_edit: "QLineEdit",
        browse_btn: "QPushButton",
        open_btn: QToolButton,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setMouseTracking(True)

        self._path_edit = path_edit
        self._browse_btn = browse_btn
        self._open_btn = open_btn

        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)

        self._label = QLabel("Папка:")
        lay.addWidget(self._label)
        lay.addWidget(path_edit, 1)
        lay.addWidget(browse_btn)
        lay.addWidget(open_btn)

        self._set_revealed(False)
        self._apply_label_style(revealed=False)

    # ── Reveal / hide ─────────────────────────────────────────────────────────

    def _set_revealed(self, revealed: bool) -> None:
        self._path_edit.setVisible(revealed)
        self._browse_btn.setVisible(revealed)
        self._apply_label_style(revealed)

    def _apply_label_style(self, revealed: bool) -> None:
        c = ThemeManager.instance().colors
        if revealed:
            self._label.setStyleSheet(
                f"color: {c.text_secondary}; font-size: 11px; font-weight: 600; min-width: 42px;"
            )
        else:
            self._label.setStyleSheet(
                f"color: {c.text_muted}; font-size: 11px; font-weight: 600; min-width: 42px;"
            )

    def apply_theme(self, c: "ThemeColors") -> None:
        revealed = self._path_edit.isVisible()
        self._apply_label_style(revealed)

    # ── Hover events ──────────────────────────────────────────────────────────

    def enterEvent(self, event) -> None:  # noqa: ANN001
        self._set_revealed(True)
        super().enterEvent(event)

    def leaveEvent(self, event) -> None:  # noqa: ANN001
        # Не скрываем, пока поле в фокусе (пользователь печатает путь)
        if not self._path_edit.hasFocus():
            self._set_revealed(False)
        super().leaveEvent(event)

    def _on_path_focus_lost(self) -> None:
        """Вызывается когда поле пути теряет фокус."""
        # Если курсор вышел за пределы — скрываем
        if not self.rect().contains(self.mapFromGlobal(
            self.cursor().pos()
        )):
            self._set_revealed(False)


# ── Зона перетаскивания файлов ────────────────────────────────────────────────

class _DropZone(QFrame):
    """Зона для перетаскивания файлов в папку документов."""

    file_dropped = Signal(str)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(72)
        self.setMaximumHeight(90)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._label = QLabel("Перетащите файл сюда для добавления в папку")
        self._label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self._label)

        self._apply_normal_style()

    def _apply_normal_style(self) -> None:
        c = ThemeManager.instance().colors
        self.setStyleSheet(f"""
QFrame {{
    background-color: {c.bg_panel};
    border: 2px dashed {c.border_base};
    border-radius: 8px;
}}
""")
        self._label.setStyleSheet(
            f"color: {c.text_muted}; font-size: 11px; background: transparent; border: none;"
        )

    def _apply_hover_style(self) -> None:
        c = ThemeManager.instance().colors
        self.setStyleSheet(f"""
QFrame {{
    background-color: {c.bg_card_hover};
    border: 2px dashed {c.accent};
    border-radius: 8px;
}}
""")

    def apply_theme(self, _c: ThemeColors) -> None:
        self._apply_normal_style()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            self._apply_hover_style()
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: ANN001
        self._apply_normal_style()
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self._apply_normal_style()
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path and Path(path).is_file():
                self.file_dropped.emit(path)
                event.acceptProposedAction()
                return
        event.ignore()


# ── Поле пути с коротким отображением ───────────────────────────────────────

class _PathLineEdit(QLineEdit):
    """QLineEdit, показывающий укороченный путь (~ вместо домашней директории) когда не в фокусе."""

    @staticmethod
    def _shorten(path: str) -> str:
        if not path:
            return path
        home_norm = str(Path.home()).replace("\\", "/")
        path_norm = path.replace("\\", "/")
        if path_norm.lower().startswith((home_norm + "/").lower()):
            return "~/" + path_norm[len(home_norm) + 1:]
        if path_norm.lower() == home_norm.lower():
            return "~"
        return path

    @staticmethod
    def _expand(path: str) -> str:
        if path.startswith("~/") or path == "~":
            home = str(Path.home()).replace("\\", "/")
            return home + path[1:]
        return path

    def setText(self, text: str) -> None:
        self.setProperty("_full_path", text)
        if self.hasFocus():
            super().setText(text)
        else:
            self.blockSignals(True)
            super().setText(self._shorten(text))
            self.blockSignals(False)

    def text(self) -> str:
        full = self.property("_full_path")
        return full if full is not None else super().text()

    def focusInEvent(self, event) -> None:
        super().focusInEvent(event)
        full = self.property("_full_path") or ""
        if full and super().text() != full:
            self.blockSignals(True)
            super().setText(full)
            self.blockSignals(False)
            self.end(False)

    def focusOutEvent(self, event) -> None:
        raw = super().text()
        full = self._expand(raw)
        self.setProperty("_full_path", full)
        super().focusOutEvent(event)
        shortened = self._shorten(full)
        if super().text() != shortened:
            self.blockSignals(True)
            super().setText(shortened)
            self.blockSignals(False)


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
        c = ThemeManager.instance().colors
        self.setStyleSheet(f"""
QTextEdit {{
    background-color: {c.bg_input};
    border: 1px solid {c.border_input};
    border-radius: 4px;
    padding: 2px 6px;
    font-size: 12px;
    color: {c.text_primary};
    selection-background-color: {c.selection_bg};
}}
QTextEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}
""")
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
        self._current_tab_index: int = 0
        self._card_field_col_width: int = _CARD_FIELD_COL_DEFAULT_W
        self._card_content_widget: QWidget | None = None

        self._autosave_timer = QTimer(self)
        self._autosave_timer.setSingleShot(True)
        self._autosave_timer.setInterval(1500)
        self._autosave_timer.timeout.connect(self._autosave)

        self._requisites_resize_timer = QTimer(self)
        self._requisites_resize_timer.setSingleShot(True)
        self._requisites_resize_timer.setInterval(0)
        self._requisites_resize_timer.timeout.connect(self._update_requisites_layout)

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
        self._hint_label = hint
        top.addWidget(hint)

        # ── Сплиттер (левый список + правые вкладки) ──────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(6)
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

        # ── Левая панель: фильтр + список ─────────────────────────────────────
        left_panel = QWidget()
        left_panel.setMinimumWidth(180)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(4)

        self._filter_edit = QLineEdit()
        self._filter_edit.setPlaceholderText("Поиск проекта...")
        self._filter_edit.setClearButtonEnabled(True)
        self._filter_edit.setStyleSheet("""
QLineEdit {
    border: 1px solid #d8e0ea;
    border-radius: 8px;
    padding: 4px 10px;
    font-size: 12px;
    background: #f8fafc;
    color: #1e2a38;
}
QLineEdit:focus {
    border-color: #5b9bd5;
    background: #ffffff;
}
""")
        left_layout.addWidget(self._filter_edit)

        self.list = _ProjectListWidget(self)
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.list.setDragEnabled(True)
        self.list.setAcceptDrops(True)
        self.list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.list.setMouseTracking(True)
        self.list.setStyleSheet(_PROJECT_LIST_STYLE)
        self._delegate = _ProjectItemDelegate(self.list)
        self.list.setItemDelegate(self._delegate)
        self.list.setSpacing(4)
        left_layout.addWidget(self.list, 1)

        splitter.addWidget(left_panel)

        # ── Вкладки ────────────────────────────────────────────────────────
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_card_tab(), "Карточка")

        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Поле", "Значение"])
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(True)
        self.table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.table.setStyleSheet(_REQUISITES_TABLE_STYLE)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.viewport().installEventFilter(self)
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
        self.tabs.addTab(self._build_docs_tab(), "Документы")

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
        self.list.json_dropped.connect(self._on_project_list_json_dropped)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self.table.itemChanged.connect(self._schedule_autosave)
        self._filter_edit.textChanged.connect(self._apply_filter)

    # ── Тема ─────────────────────────────────────────────────────────────────

    def _docs_list_style(self) -> str:
        c = ThemeManager.instance().colors
        return f"""
QListWidget {{
    background-color: {c.bg_panel};
    border: 1px solid {c.border_base};
    border-radius: 6px;
    padding: 3px;
    font-size: 12px;
    outline: 0;
    color: {c.text_primary};
}}
QListWidget::item {{ padding: 4px 8px; border-radius: 3px; }}
QListWidget::item:selected {{
    background: {c.selection_bg};
    color: {c.selection_text};
}}
QListWidget::item:hover {{ background: {c.bg_hover}; }}
QScrollBar:vertical {{
    background: {c.bg_scrollbar}; width: 6px; border-radius: 3px; margin: 4px 2px;
}}
QScrollBar::handle:vertical {{
    background: {c.scrollbar_handle}; border-radius: 3px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {c.scrollbar_handle_hover}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
"""

    def _preview_scroll_style(self, c: "ThemeColors | None" = None) -> str:
        if c is None:
            c = ThemeManager.instance().colors
        return f"""
QScrollArea {{
    background: {c.bg_panel};
    border: 1px solid {c.border_base};
    border-radius: 6px;
}}
QScrollBar:vertical {{
    background: {c.bg_scrollbar}; width: 6px; border-radius: 3px; margin: 4px 2px;
}}
QScrollBar::handle:vertical {{
    background: {c.scrollbar_handle}; border-radius: 3px; min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {c.scrollbar_handle_hover}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
"""

    def _rename_btn_style(self, c: "ThemeColors | None" = None) -> str:
        if c is None:
            c = ThemeManager.instance().colors
        return f"""
QPushButton {{
    background: {c.accent};
    border: none;
    border-radius: 4px;
    padding: 2px 12px;
    font-size: 11px;
    font-weight: 600;
    color: {c.accent_text};
    min-height: 22px;
}}
QPushButton:hover {{ background: {c.accent_hover}; }}
QPushButton:pressed {{ background: {c.accent_pressed}; }}
QPushButton:disabled {{ background: {c.accent_disabled}; color: {c.text_muted}; }}
"""

    def _nav_btn_style(self, c: "ThemeColors | None" = None) -> str:
        if c is None:
            c = ThemeManager.instance().colors
        return f"""
QPushButton {{
    background: {c.bg_hover};
    border: 1px solid {c.border_base};
    border-radius: 4px;
    padding: 2px 10px;
    font-size: 10px;
    color: {c.text_secondary};
    min-height: 20px;
}}
QPushButton:hover {{ background: {c.bg_card_hover}; }}
QPushButton:pressed {{ background: {c.bg_alternate}; }}
QPushButton:disabled {{ color: {c.text_muted}; }}
"""

    def _table_style(self) -> str:
        c = ThemeManager.instance().colors
        return f"""
QTableWidget {{
    background-color: {c.bg_panel};
    alternate-background-color: {c.bg_alternate};
    border: 1px solid {c.border_base};
    border-radius: 10px;
    gridline-color: {c.border_light};
    outline: 0;
    color: {c.text_primary};
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
}}
QTableWidget::item {{ padding: 6px 8px; border: none; }}
QHeaderView::section {{
    background-color: {c.bg_header};
    color: {c.text_secondary};
    border: none;
    border-bottom: 1px solid {c.border_base};
    padding: 7px 8px;
    font-size: 11px;
    font-weight: 600;
}}
QScrollBar:vertical {{
    background: {c.bg_scrollbar};
    width: 6px;
    border-radius: 3px;
    margin: 6px 2px;
}}
QScrollBar::handle:vertical {{
    background: {c.scrollbar_handle};
    border-radius: 3px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {c.scrollbar_handle_hover}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}
"""

    def apply_theme(self, c: ThemeColors) -> None:
        """Применяет тему ко всем виджетам вкладки."""
        # Кнопки-иконки верхней панели
        update_icon_btn(self.load_btn, _SVG_REFRESH, icon_color=c.icon_color,
                        bg=c.icon_btn_bg, hover=c.icon_btn_hover, pressed=c.icon_btn_pressed)
        update_icon_btn(self.save_btn, _SVG_SAVE, icon_color=c.icon_color,
                        bg=c.icon_btn_bg, hover=c.icon_btn_hover, pressed=c.icon_btn_pressed)
        update_icon_btn(self.add_btn, _SVG_ADD, icon_color=c.icon_color,
                        bg=c.icon_btn_bg, hover=c.icon_btn_hover, pressed=c.icon_btn_pressed)

        # Подсказка
        self._hint_label.setStyleSheet(
            f"color: {c.text_muted}; font-style: italic; font-size: 11px;"
        )

        # Поле поиска
        self._filter_edit.setStyleSheet(f"""
QLineEdit {{
    border: 1px solid {c.border_base};
    border-radius: 8px;
    padding: 4px 10px;
    font-size: 12px;
    background: {c.bg_input};
    color: {c.text_primary};
}}
QLineEdit:focus {{
    border-color: {c.border_input_focus};
    background: {c.bg_input_focus};
}}
""")

        # Список проектов
        self.list.setStyleSheet(f"""
QListWidget {{
    background-color: {c.bg_panel};
    border: 1px solid {c.border_base};
    border-radius: 12px;
    outline: 0;
    padding: 8px 7px;
}}
QListWidget::item {{
    border: none;
    background: transparent;
    padding: 0;
    margin: 0 0 4px 0;
}}
QScrollBar:vertical {{
    background: {c.bg_scrollbar};
    width: 6px;
    border-radius: 3px;
    margin: 6px 2px 6px 2px;
}}
QScrollBar::handle:vertical {{
    background: {c.scrollbar_handle};
    border-radius: 3px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{ background: {c.scrollbar_handle_hover}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: none; }}
""")

        # Обновляем делегат и перерисовываем список
        self._delegate.set_colors(c)
        self.list.viewport().update()

        # Таблица реквизитов
        self.table.setStyleSheet(self._table_style())

        # Карточка: фон скролл-зоны и страницы
        if hasattr(self, "_card_scroll"):
            self._card_scroll.setStyleSheet(
                f"QScrollArea, QScrollArea > QWidget > QWidget {{ background: {c.bg_window}; }}"
            )
        if hasattr(self, "_card_page"):
            self._card_page.setStyleSheet(f"background: {c.bg_window};")

        # Разделители и заголовки в карточке
        if hasattr(self, "_card_title_sep"):
            self._card_title_sep.setStyleSheet(f"color: {c.separator}; margin: 4px 0 2px 0;")

        # Заголовок карточки
        if hasattr(self, "_card_title_edit"):
            self._card_title_edit.setStyleSheet(f"""
QLineEdit {{
    background: transparent;
    border: none;
    border-bottom: 1.5px solid {c.border_base};
    border-radius: 0;
    padding: 2px 2px 4px 2px;
    font-size: 14px;
    font-weight: 700;
    color: {c.text_primary};
}}
QLineEdit:focus {{ border-bottom: 2px solid {c.border_input_focus}; }}
QLineEdit:hover {{ border-bottom-color: {c.scrollbar_handle_hover}; }}
""")

        # Фиксированные поля карточки
        fixed_style = f"""
QLineEdit {{
    background-color: {c.bg_input};
    border: 1px solid {c.border_input};
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 12px;
    color: {c.text_primary};
}}
QLineEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}
"""
        if hasattr(self, "_card_fixed_edits"):
            for edit in self._card_fixed_edits.values():
                edit.setStyleSheet(fixed_style)

        textedit_style = f"""
QTextEdit {{
    background-color: {c.bg_input};
    border: 1px solid {c.border_input};
    border-radius: 4px;
    padding: 2px 6px;
    font-size: 12px;
    color: {c.text_primary};
    selection-background-color: {c.selection_bg};
}}
QTextEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}
"""
        if hasattr(self, "_card_fixed_edit_widgets") and PROJECT_COMMENT_FIELD in self._card_fixed_edit_widgets:
            comment_edit = self._card_fixed_edit_widgets[PROJECT_COMMENT_FIELD]
            if isinstance(comment_edit, _AutoResizeTextEdit):
                comment_edit.setStyleSheet(textedit_style)

        # Drop-зона документов
        if hasattr(self, "_docs_drop_zone"):
            self._docs_drop_zone.apply_theme(c)

        # Строка папки документов
        if hasattr(self, "_docs_folder_row"):
            self._docs_folder_row.apply_theme(c)
            self._docs_path_edit.setStyleSheet(f"""
QLineEdit {{
    background-color: {c.bg_input};
    border: 1px solid {c.border_input};
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 12px;
    color: {c.text_primary};
}}
QLineEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}
""")
            self._docs_folder_row._browse_btn.setStyleSheet(f"""
QPushButton {{
    background: {c.bg_hover};
    border: 1px solid {c.border_base};
    border-radius: 4px;
    padding: 2px 10px;
    font-size: 11px;
    color: {c.text_secondary};
    min-height: 22px;
}}
QPushButton:hover {{ background: {c.bg_card_hover}; }}
QPushButton:pressed {{ background: {c.bg_alternate}; }}
""")
            update_icon_btn(
                self._docs_open_folder_btn, _SVG_FOLDER,
                icon_color="#ffffff", bg=c.accent, hover=c.accent_hover, pressed=c.accent_pressed
            )

        # Список файлов в документах
        if hasattr(self, "_docs_list"):
            self._docs_list.setStyleSheet(self._docs_list_style())

        # Заголовки панели документов
        if hasattr(self, "_docs_files_label"):
            self._docs_files_label.setStyleSheet(
                f"color: {c.text_muted}; font-size: 10px; font-weight: 700; letter-spacing: 0.3px;"
            )
        if hasattr(self, "_docs_hint_label"):
            self._docs_hint_label.setStyleSheet(
                f"color: {c.text_muted}; font-size: 10px; font-style: italic;"
            )

        # Панель предпросмотра
        if hasattr(self, "_preview_scroll"):
            self._preview_scroll.setStyleSheet(self._preview_scroll_style(c))
        if hasattr(self, "_preview_label"):
            self._preview_label.setStyleSheet(
                f"background: transparent; color: {c.text_muted}; font-size: 12px; padding: 40px;"
            )
        if hasattr(self, "_preview_name_label"):
            self._preview_name_label.setStyleSheet(
                f"color: {c.text_secondary}; font-size: 11px; font-weight: 600;"
            )
        if hasattr(self, "_preview_rename_edit"):
            self._preview_rename_edit.setStyleSheet(f"""
QLineEdit {{
    background-color: {c.bg_input};
    border: 1px solid {c.border_input};
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 12px;
    color: {c.text_primary};
}}
QLineEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}
""")
        if hasattr(self, "_preview_rename_btn"):
            self._preview_rename_btn.setStyleSheet(self._rename_btn_style(c))
        if hasattr(self, "_preview_prev_btn"):
            self._preview_prev_btn.setStyleSheet(self._nav_btn_style(c))
        if hasattr(self, "_preview_next_btn"):
            self._preview_next_btn.setStyleSheet(self._nav_btn_style(c))
        if hasattr(self, "_preview_page_label"):
            self._preview_page_label.setStyleSheet(
                f"color: {c.text_secondary}; font-size: 10px; min-width: 100px;"
            )

    # ── Построение вкладки «Карточка» ────────────────────────────────────────

    def _build_card_tab(self) -> QWidget:
        wrapper = QWidget()
        outer = QVBoxLayout(wrapper)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        c0 = ThemeManager.instance().colors
        scroll.setStyleSheet(f"QScrollArea, QScrollArea > QWidget > QWidget {{ background: {c0.bg_window}; }}")
        self._card_scroll = scroll

        content = QWidget()
        content.setMaximumWidth(600)
        self._card_content_widget = content
        self._card_content_vbox = QVBoxLayout(content)
        self._card_content_vbox.setContentsMargins(12, 8, 12, 8)
        self._card_content_vbox.setSpacing(5)
        self._card_left_col_widgets: list[QWidget] = []
        self._card_fixed_edit_widgets: dict[str, QWidget] = {}

        # ── Поле «Имя проекта» — заголовок карточки ──────────────────────
        self._card_title_edit = QLineEdit()
        self._card_title_edit.setPlaceholderText("Имя проекта")
        self._card_title_edit.setStyleSheet(_CARD_TITLE_EDIT_STYLE)
        self._card_title_edit.setSizePolicy(
            QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed
        )
        self._card_title_edit.setMinimumWidth(120)

        def _sync_title_width() -> None:
            fm = self._card_title_edit.fontMetrics()
            txt = self._card_title_edit.text() or self._card_title_edit.placeholderText()
            w = fm.horizontalAdvance(txt) + 20
            self._card_title_edit.setMinimumWidth(min(w, 560))

        self._card_title_edit.textChanged.connect(lambda _: _sync_title_width())
        self._card_title_edit.textChanged.connect(self._on_card_title_changed)
        self._card_title_edit.editingFinished.connect(self._schedule_autosave)

        title_sep = QFrame()
        title_sep.setFrameShape(QFrame.Shape.HLine)
        c0 = ThemeManager.instance().colors
        title_sep.setStyleSheet(f"color: {c0.separator}; margin: 4px 0 2px 0;")
        self._card_title_sep = title_sep
        self._card_content_vbox.addWidget(self._card_title_edit)
        self._card_content_vbox.addWidget(title_sep)

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
        self._card_fixed_edits: dict[str, QWidget] = {}
        self._card_fixed_rows: dict[str, QWidget] = {}
        self._card_link_btns: list[tuple[QToolButton, str]] = []
        for field_name in _CARD_FIXED_FIELDS:
            fw, edit = self._make_fixed_field_row(
                field_name,
                link_field_name=_CARD_LINK_FIELDS.get(field_name),
            )
            self._card_fixed_edits[field_name] = edit
            self._card_fixed_edit_widgets[field_name] = edit
            self._card_fixed_rows[field_name] = fw
            self._card_content_vbox.addWidget(fw)

        self._card_fields_end_anchor = QWidget()
        self._card_fields_end_anchor.setVisible(False)
        self._card_content_vbox.addWidget(self._card_fields_end_anchor)

        self._card_content_vbox.addStretch()

        # Оборачиваем content в страничный виджет: content + stretch справа.
        # Это гарантирует, что контент не растягивается на всю ширину окна.
        page = QWidget()
        page.setStyleSheet(f"background: {c0.bg_window};")
        self._card_page = page
        page_hbox = QHBoxLayout(page)
        page_hbox.setContentsMargins(0, 0, 0, 0)
        page_hbox.setSpacing(0)
        page_hbox.addWidget(content)
        page_hbox.addStretch()
        scroll.setWidget(page)
        outer.addWidget(scroll)

        self._set_card_field_col_width(self._card_field_col_width)
        return wrapper

    def _make_fixed_field_row(
        self,
        field_name: str,
        *,
        label_text: str | None = None,
        link_field_name: str | None = None,
    ) -> tuple[QWidget, QWidget]:
        row = QWidget()
        row.setStyleSheet("background: transparent;")
        hbox = QHBoxLayout(row)
        hbox.setContentsMargins(0, 0, 0, 0)
        hbox.setSpacing(_CARD_COL_GAP)

        label = QLabel(label_text or field_name)
        label.setStyleSheet(_CARD_LABEL_CSS)
        self._register_card_left_widget(label)

        if field_name == PROJECT_COMMENT_FIELD:
            edit = _AutoResizeTextEdit(placeholder="—")
            edit.setMinimumHeight(22)
            edit.textChanged.connect(self._reorder_card_fixed_rows)
            edit.textChanged.connect(self._schedule_autosave)
        else:
            edit = QLineEdit()
            edit.setPlaceholderText("—")
            edit.setStyleSheet(_FIXED_VALUE_STYLE)
            edit.setFixedHeight(22)
            edit.textChanged.connect(lambda _text, self=self: self._reorder_card_fixed_rows())
            edit.editingFinished.connect(self._schedule_autosave)

        hbox.addWidget(label)
        hbox.addWidget(edit, 1)
        if link_field_name and isinstance(edit, QLineEdit):
            lb = _link_btn()
            hbox.addWidget(lb)
            self._wire_link_for_fixed_field(edit, lb, link_field_name)
            self._card_link_btns.append((lb, link_field_name))
        return row, edit

    def _wire_link_for_fixed_field(
        self,
        edit: QLineEdit,
        link_btn: QToolButton,
        link_field_name: str,
    ) -> None:
        """Ctrl+K на поле сохраняет URL в отдельное поле link_field_name.
        Иконка ссылки горит, когда это поле непустое."""

        def open_link() -> None:
            if not self._current:
                return
            url = self._current.fields.get(link_field_name, "").strip()
            if url:
                webbrowser.open(url)

        def insert_link() -> None:
            if not self._current:
                return
            existing = self._current.fields.get(link_field_name, "").strip()
            url_raw, ok = QInputDialog.getText(
                self, "Гиперссылка", "URL:", QLineEdit.EchoMode.Normal, existing
            )
            if not ok:
                return
            url = url_raw.strip()
            self._current.fields[link_field_name] = url
            link_btn.setVisible(bool(url))
            link_btn.setProperty("_url", url)
            self._schedule_autosave()

        link_btn.clicked.connect(open_link)

        sc = QShortcut(QKeySequence("Ctrl+K"), edit)
        sc.setContext(Qt.ShortcutContext.WidgetShortcut)
        sc.activated.connect(insert_link)

    # ── Карточка: рендер и чтение ─────────────────────────────────────────────

    def _reorder_card_fixed_rows(self) -> None:
        if not hasattr(self, "_card_fields_end_anchor"):
            return
        pinned = [PROJECT_TYPE_FIELD] if PROJECT_TYPE_FIELD in self._card_fixed_rows else []
        ordered = [
            field_name
            for field_name in _CARD_FIXED_FIELDS
            if field_name not in pinned and self._card_field_value(field_name)
        ]
        ordered += [
            field_name
            for field_name in _CARD_FIXED_FIELDS
            if field_name not in pinned and not self._card_field_value(field_name)
        ]
        ordered = pinned + ordered
        for row_widget in self._card_fixed_rows.values():
            self._card_content_vbox.removeWidget(row_widget)
        insert_at = self._card_content_vbox.indexOf(self._card_fields_end_anchor)
        for offset, field_name in enumerate(ordered):
            row_widget = self._card_fixed_rows.get(field_name)
            if row_widget is not None:
                self._card_content_vbox.insertWidget(insert_at + offset, row_widget)

    def _card_field_value(self, field_name: str) -> str:
        widget = self._card_fixed_edits.get(field_name)
        if isinstance(widget, QLineEdit):
            return widget.text().strip()
        if isinstance(widget, _AutoResizeTextEdit):
            return widget.toPlainText().strip()
        return ""

    def _render_card(self, project: Project) -> None:
        self._card_title_edit.blockSignals(True)
        self._card_title_edit.setText(project.fields.get(PROJECT_NAME_FIELD, ""))
        self._card_title_edit.blockSignals(False)

        for field_name, edit in self._card_fixed_edits.items():
            edit.blockSignals(True)
            if field_name == _CASE_NUMBER_FIELD_NEW:
                val = project.fields.get(_CASE_NUMBER_FIELD_NEW, "").strip()
                if not val:
                    val = project.fields.get(_CASE_NUMBER_FIELD_OLD, "").strip()
                    if val:
                        # Мягкая миграция: чтобы реквизиты/таблица тоже видели новое имя
                        project.fields[_CASE_NUMBER_FIELD_NEW] = val
                if isinstance(edit, QLineEdit):
                    edit.setText(val)
            else:
                value = project.fields.get(field_name, "")
                if isinstance(edit, QLineEdit):
                    edit.setText(value)
                elif isinstance(edit, _AutoResizeTextEdit):
                    edit.setPlainText(value)
            edit.blockSignals(False)
        self._reorder_card_fixed_rows()

        # Обновляем состояние иконок ссылок для текущего проекта.
        for link_btn, link_field_name in self._card_link_btns:
            url = project.fields.get(link_field_name, "").strip()
            link_btn.setVisible(bool(url))
            link_btn.setProperty("_url", url)

    def _read_card_into_project(self, project: Project) -> None:
        project.fields[PROJECT_NAME_FIELD] = self._card_title_edit.text().strip()

        for field_name, edit in self._card_fixed_edits.items():
            if isinstance(edit, QLineEdit):
                val = edit.text().strip()
            elif isinstance(edit, _AutoResizeTextEdit):
                val = edit.toPlainText().strip()
            else:
                val = ""
            if field_name == _CASE_NUMBER_FIELD_NEW:
                project.fields[_CASE_NUMBER_FIELD_NEW] = val
                # Для совместимости: если в Excel/шапках есть старое поле — держим в синхроне
                headers = getattr(project, "headers", None) or []
                if _CASE_NUMBER_FIELD_OLD in headers or _CASE_NUMBER_FIELD_OLD in project.fields:
                    project.fields[_CASE_NUMBER_FIELD_OLD] = val
            else:
                project.fields[field_name] = val

        row = self.list.currentRow()
        item = self.list.item(row)
        if item is not None:
            self._refresh_list_item_text(project)

    def _clear_card_display(self) -> None:
        self._card_title_edit.blockSignals(True)
        self._card_title_edit.setText("")
        self._card_title_edit.blockSignals(False)
        for edit in self._card_fixed_edits.values():
            edit.blockSignals(True)
            if isinstance(edit, QLineEdit):
                edit.setText("")
            elif isinstance(edit, _AutoResizeTextEdit):
                edit.setPlainText("")
            edit.blockSignals(False)
        for link_btn, _ in self._card_link_btns:
            link_btn.setVisible(False)

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

    def eventFilter(self, obj, event) -> bool:  # noqa: ANN001
        if (
            hasattr(self, "_preview_scroll")
            and obj is self._preview_scroll.viewport()
            and event.type() == QEvent.Type.Resize
        ):
            self._scale_preview()
        if obj is self.table.viewport() and event.type() == QEvent.Type.Resize:
            # При изменении ширины — пересчитать высоты строк для переноса текста.
            self._requisites_resize_timer.start()
        return super().eventFilter(obj, event)

    # ── Синхронизация вкладок ─────────────────────────────────────────────────

    def _on_tab_changed(self, new_index: int) -> None:
        if self._current is None:
            self._current_tab_index = new_index
            return
        old_index = self._current_tab_index
        self._current_tab_index = new_index
        if old_index == 0:
            self._read_card_into_project(self._current)
        elif old_index == 1:
            self._read_table_into_project(self._current)
        if new_index == 0:
            self._render_card(self._current)
        elif new_index == 1:
            self._render_project(self._current)
        elif new_index == 2:
            self._load_project_docs_path()
            self._refresh_docs_list()

    def _sync_current_to_project(self) -> None:
        if self._current is None:
            return
        if self._current_tab_index == 0:
            self._read_card_into_project(self._current)
        elif self._current_tab_index == 1:
            self._read_table_into_project(self._current)

    # ── Drag & drop ───────────────────────────────────────────────────────────

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls() and any(
            u.toLocalFile().lower().endswith(".json") for u in event.mimeData().urls()
        ):
            c = ThemeManager.instance().colors
            self.table.setStyleSheet(f"QTableWidget {{ border: 2px dashed {c.accent}; border-radius: 4px; }}")
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event) -> None:  # noqa: ANN001
        self.table.setStyleSheet(self._table_style())
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self.table.setStyleSheet(self._table_style())
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".json"):
                self._load_from_json(path)
                event.acceptProposedAction()
                return
        event.ignore()

    def _read_json_fields(self, path: str) -> dict[str, str] | None:
        try:
            with open(path, encoding="utf-8") as f:
                data: dict = json.load(f)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "JSON", f"Не удалось прочитать файл:\n{e}")
            return None

        if not isinstance(data, dict):
            QMessageBox.warning(self, "JSON", "Файл должен содержать JSON-объект (словарь полей).")
            return None

        return {str(k): "" if v is None else str(v) for k, v in data.items()}

    def _load_from_json(self, path: str) -> None:
        fields = self._read_json_fields(path)
        if fields is None:
            return

        case_number = str(fields.get("№ дела", "") or fields.get("Номер дела", "")).strip()
        project_id = case_number or Path(path).stem
        project = Project(project_id=project_id, fields=fields, headers=list(fields.keys()))

        self._projects.append(project)
        self._add_project_to_list(project)
        self.list.setCurrentRow(len(self._projects) - 1)

    def _on_project_list_json_dropped(self, path: str, row: int) -> None:
        if row < 0 or row >= self.list.count():
            self._load_from_json(path)
            return

        item = self.list.item(row)
        if item is None:
            self._load_from_json(path)
            return

        project = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(project, Project):
            self._load_from_json(path)
            return

        self.list.setCurrentRow(row)
        self._merge_project_from_json(project, path)

    def _merge_project_from_json(self, project: Project, path: str) -> None:
        fields = self._read_json_fields(path)
        if fields is None:
            return

        if project is self._current:
            self._sync_current_to_project()

        added_count = 0
        replaced_count = 0
        kept_count = 0

        for key, new_value_raw in fields.items():
            new_value = new_value_raw.strip()
            old_value = str(project.fields.get(key, "") or "").strip()

            if key not in project.fields or old_value == "":
                project.fields[key] = new_value_raw
                added_count += 1
                continue

            if new_value == "" or old_value == new_value:
                kept_count += 1
                continue

            answer = QMessageBox.question(
                self,
                "Конфликт значения",
                (
                    f"В проекте уже есть значение для поля:\n\n"
                    f"{key}\n\n"
                    f"Текущее значение:\n{old_value}\n\n"
                    f"Новое значение из JSON:\n{new_value}\n\n"
                    f"Заменить текущее значение новым?"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer == QMessageBox.StandardButton.Yes:
                project.fields[key] = new_value_raw
                replaced_count += 1
            else:
                kept_count += 1

        headers = list(project.headers or [])
        for key in fields:
            if key and key not in headers:
                headers.append(key)
        if headers:
            project.headers = headers

        self._refresh_project_name(project)
        self._refresh_list_item_text(project)
        if project is self._current:
            self._render_card(project)
            self._render_project(project)

        self._schedule_autosave()
        self._show_status(
            f"JSON загружен: +{added_count} новых, ~{replaced_count} обновлено, {kept_count} без изменений"
        )

    # ── Список проектов ───────────────────────────────────────────────────────

    def _add_project_to_list(self, project: Project, *, archived: bool = False) -> None:
        name = self._project_display_name(project)
        item = QListWidgetItem(name)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        item.setData(Qt.ItemDataRole.UserRole, project)
        if archived:
            item.setForeground(QColor("#777777"))
            item.setBackground(QColor("#f2f2f2"))
        self.list.addItem(item)

    def _apply_filter(self, text: str) -> None:
        """При пустом запросе восстанавливает нормальный вид; при непустом — показывает
        совпадающие проекты из обоих списков (архивные — зачёркнутым стилем)."""
        query = text.strip().lower()
        if not query:
            if self._showing_archive:
                self._show_archived_projects()
            else:
                self._show_current_projects()
            return

        def _matches(p: Project) -> bool:
            label = self._project_display_name(p).lower()
            extra = " ".join(str(v) for v in p.fields.values()).lower()
            return query in label or query in extra

        self.list.clear()
        for p in self._projects:
            if _matches(p):
                self._add_project_to_list(p, archived=False)
        for p in self._archived_projects:
            if _matches(p):
                self._add_project_to_list(p, archived=True)

        if self.list.count() > 0:
            self.list.setCurrentRow(0)
        else:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()

    def _project_display_name(self, project: Project) -> str:
        name = (project.fields.get(PROJECT_NAME_FIELD) or "").strip()
        if name:
            return name
        auto = _auto_project_name(project.fields)
        return auto or project.project_id

    def _refresh_project_name(self, project: Project) -> None:
        """Заполняет 'Имя проекта' из Кредитор и Должник, только если поле ещё не задано вручную."""
        existing = (project.fields.get(PROJECT_NAME_FIELD) or "").strip()
        if existing:
            return
        auto = _auto_project_name(project.fields)
        if auto:
            project.fields[PROJECT_NAME_FIELD] = auto

    def _refresh_list_item_text(self, project: Project) -> None:
        for i in range(self.list.count()):
            item = self.list.item(i)
            if item is None or item.data(Qt.ItemDataRole.UserRole) is not project:
                continue
            name = self._project_display_name(project)
            self.list.itemChanged.disconnect(self._on_list_item_edited)
            item.setText(name)
            self.list.itemChanged.connect(self._on_list_item_edited)
            break

    def _on_list_item_edited(self, item: QListWidgetItem) -> None:
        project = item.data(Qt.ItemDataRole.UserRole)
        if isinstance(project, Project):
            project.fields[PROJECT_NAME_FIELD] = (item.text() or "").strip()
            self._schedule_autosave()

    def _on_card_title_changed(self, text: str) -> None:
        """Мгновенно обновляет имя проекта и элемент списка при вводе в поле заголовка карточки."""
        if self._current is None:
            return
        self._current.fields[PROJECT_NAME_FIELD] = text.strip()
        self._refresh_list_item_text(self._current)
        self._schedule_autosave()

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

        # Определяем статус элемента по его реальной принадлежности к архиву,
        # а не по глобальному флагу (важно при поиске в смешанном режиме).
        item_project = item.data(Qt.ItemDataRole.UserRole)
        item_is_archived = (
            isinstance(item_project, Project)
            and item_project in self._archived_projects
        )

        if item_is_archived:
            unarchive_action = menu.addAction("Убрать из архива")
            delete_archived_action = menu.addAction("Удалить")
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

        if item_is_archived and chosen_action == unarchive_action:
            self._unarchive_current()
        elif item_is_archived and chosen_action == delete_archived_action:
            self._delete_archived_current()
        elif not item_is_archived and chosen_action == archive_action:
            self._archive_current()
        elif not item_is_archived and chosen_action == delete_action:
            self._delete_current()

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

    def _delete_archived_current(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Архив", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        row = self.list.currentRow()
        if row < 0 or row >= self.list.count():
            QMessageBox.warning(self, "Удаление", "Не выбран проект для удаления.")
            return
        item = self.list.item(row)
        if item is None:
            QMessageBox.warning(self, "Удаление", "Не выбран проект для удаления.")
            return
        project = item.data(Qt.ItemDataRole.UserRole)
        if not isinstance(project, Project):
            QMessageBox.warning(self, "Удаление", "Не удалось определить выбранный проект.")
            return

        title = self._project_display_name(project)
        answer = QMessageBox.question(
            self, "Удаление архивного проекта",
            f"Удалить проект из архива:\n{title}?\n\nЭто действие необратимо.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        try:
            store = ExcelProjectStore(self._settings.excel_path)
            store.delete_project_from_archive(project)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Удаление", f"Не удалось удалить проект из архива:\n{e}")
            return

        if project in self._archived_projects:
            self._archived_projects.remove(project)
        self.list.takeItem(row)

        if self._archived_projects:
            self.list.setCurrentRow(min(row, len(self._archived_projects) - 1))
        else:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()

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
        # Если был активен поиск — восстанавливаем его
        filter_text = self._filter_edit.text()
        if filter_text.strip():
            self._apply_filter(filter_text)

    # ── Статус-бар ────────────────────────────────────────────────────────────

    def _show_status(self, message: str, timeout_ms: int = 4000) -> None:
        mw = self.window()
        if hasattr(mw, "show_status"):
            mw.show_status(message, timeout_ms)

    # ── Настройки ─────────────────────────────────────────────────────────────

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s
        if hasattr(self, "_docs_path_edit"):
            self._load_project_docs_path()
            if self._current_tab_index == 2:
                self._refresh_docs_list()

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

            self._filter_edit.blockSignals(True)
            self._filter_edit.clear()
            self._filter_edit.blockSignals(False)
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
        if hasattr(self, "_docs_path_edit"):
            self._load_project_docs_path()
            if self._current_tab_index == 2:
                self._refresh_docs_list()

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
        self.table.blockSignals(True)
        self.table.setRowCount(len(items))
        for i, (k, v) in enumerate(items):
            key_item = QTableWidgetItem(str(k))
            key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, key_item)
            value_item = QTableWidgetItem(str(v))
            if k in {"Кредитор", "Должник"} and str(v).strip() == "":
                value_item.setBackground(QColor("#ffd6e7"))
            self.table.setItem(i, 1, value_item)
        self.table.blockSignals(False)
        self.table.resizeColumnToContents(0)
        self._update_requisites_layout()

    def _update_requisites_layout(self) -> None:
        # Важное: высота строк зависит от ширины колонки "Значение".
        try:
            self.table.resizeRowsToContents()
        except Exception:  # noqa: BLE001
            pass

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
        self.list.takeItem(row)

        if self._projects:
            self.list.setCurrentRow(min(row, len(self._projects) - 1))
        else:
            self._current = None
            self.table.setRowCount(0)
            self._clear_card_display()

    def _schedule_autosave(self) -> None:
        self._autosave_timer.start()

    def _autosave(self) -> None:
        if not self._settings.excel_path or not self._projects:
            return
        self._save_all(silent=True)

    def _save_all(self, *, silent: bool = False) -> None:
        if not self._settings.excel_path:
            if not silent:
                QMessageBox.warning(self, "Проекты", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        if not self._projects:
            if not silent:
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
            if not silent:
                self._show_status("Все изменения синхронизированы с Excel")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))

    # ── Вкладка «Документы» ───────────────────────────────────────────────────

    def _build_docs_tab(self) -> QWidget:
        widget = QWidget()
        vbox = QVBoxLayout(widget)
        vbox.setContentsMargins(10, 8, 10, 8)
        vbox.setSpacing(6)

        # Строка выбора папки (скрыта до наведения курсора)
        c0 = ThemeManager.instance().colors
        self._docs_path_edit = _PathLineEdit()
        self._docs_path_edit.setPlaceholderText("Путь к папке документов для этого проекта…")
        self._docs_path_edit.setStyleSheet(f"""
QLineEdit {{
    background-color: {c0.bg_input};
    border: 1px solid {c0.border_input};
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 12px;
    color: {c0.text_primary};
}}
QLineEdit:focus {{
    border-color: {c0.border_input_focus};
    background-color: {c0.bg_input_focus};
}}
""")

        browse_docs_btn = QPushButton("Выбрать…")
        browse_docs_btn.setStyleSheet(f"""
QPushButton {{
    background: {c0.bg_hover};
    border: 1px solid {c0.border_base};
    border-radius: 4px;
    padding: 2px 10px;
    font-size: 11px;
    color: {c0.text_secondary};
    min-height: 22px;
}}
QPushButton:hover {{ background: {c0.bg_card_hover}; }}
QPushButton:pressed {{ background: {c0.bg_alternate}; }}
""")
        browse_docs_btn.setCursor(Qt.CursorShape.PointingHandCursor)

        open_folder_btn = _icon_btn(
            _SVG_FOLDER, "Открыть папку в проводнике",
            "#ffffff", c0.accent, c0.accent_hover, c0.accent_pressed,
        )
        self._docs_open_folder_btn = open_folder_btn

        self._docs_folder_row = _DocsFolderRow(
            self._docs_path_edit, browse_docs_btn, open_folder_btn
        )
        vbox.addWidget(self._docs_folder_row)

        # Зона перетаскивания
        self._docs_drop_zone = _DropZone()
        self._docs_drop_zone.file_dropped.connect(self._on_file_dropped)
        vbox.addWidget(self._docs_drop_zone)

        # Сплиттер: список файлов | предпросмотр
        docs_splitter = QSplitter(Qt.Orientation.Horizontal)
        docs_splitter.setHandleWidth(5)
        docs_splitter.setChildrenCollapsible(False)

        # ── Левая панель: заголовок + список файлов ──
        left_panel = QWidget()
        left_vbox = QVBoxLayout(left_panel)
        left_vbox.setContentsMargins(0, 0, 0, 0)
        left_vbox.setSpacing(4)

        files_header = QHBoxLayout()
        files_label = QLabel("Файлы в папке")
        files_label.setStyleSheet(
            f"color: {c0.text_muted}; font-size: 10px; font-weight: 700; letter-spacing: 0.3px;"
        )
        self._docs_files_label = files_label
        hint_label = QLabel("двойной клик — переименовать")
        hint_label.setStyleSheet(f"color: {c0.text_muted}; font-size: 10px; font-style: italic;")
        self._docs_hint_label = hint_label
        refresh_docs_btn = _mini_btn("↻", "Обновить список файлов")
        files_header.addWidget(files_label)
        files_header.addWidget(hint_label)
        files_header.addStretch()
        files_header.addWidget(refresh_docs_btn)
        left_vbox.addLayout(files_header)

        self._docs_list = QListWidget()
        self._docs_list.setStyleSheet(self._docs_list_style())
        self._docs_list.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.SelectedClicked
        )
        left_vbox.addWidget(self._docs_list, 1)
        docs_splitter.addWidget(left_panel)

        # ── Правая панель: предпросмотр ──
        preview_panel = self._build_preview_panel()
        docs_splitter.addWidget(preview_panel)

        docs_splitter.setSizes([280, 520])
        docs_splitter.setStretchFactor(0, 0)
        docs_splitter.setStretchFactor(1, 1)

        vbox.addWidget(docs_splitter, 1)

        # Подключаем сигналы
        self._docs_folder_row._browse_btn.clicked.connect(self._browse_docs_dir)
        self._docs_open_folder_btn.clicked.connect(self._open_docs_folder)
        self._docs_path_edit.editingFinished.connect(self._on_docs_path_changed)
        self._docs_path_edit.editingFinished.connect(
            self._docs_folder_row._on_path_focus_lost
        )
        self._docs_list.itemChanged.connect(self._on_doc_renamed)
        self._docs_list.currentItemChanged.connect(self._on_doc_selected)
        refresh_docs_btn.clicked.connect(self._refresh_docs_list)

        return widget

    def _refresh_docs_list(self) -> None:
        """Обновляет список файлов из текущей папки документов."""
        docs_path = self._docs_path_edit.text().strip()
        self._docs_list.blockSignals(True)
        self._docs_list.clear()
        if docs_path:
            p = Path(docs_path)
            if p.is_dir():
                try:
                    files = sorted(
                        (f for f in p.iterdir() if f.is_file()),
                        key=lambda x: x.name.lower(),
                    )
                    for f in files:
                        item = QListWidgetItem(f.name)
                        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                        item.setData(Qt.ItemDataRole.UserRole, str(f))
                        self._docs_list.addItem(item)
                except Exception as e:  # noqa: BLE001
                    QMessageBox.warning(
                        self, "Документы",
                        f"Не удалось прочитать содержимое папки:\n{e}",
                    )
        self._docs_list.blockSignals(False)
        if hasattr(self, "_preview_label"):
            self._clear_preview()

    def _on_file_dropped(self, src_path: str) -> None:
        """Копирует перетащенный файл в папку документов."""
        docs_dir = self._docs_path_edit.text().strip()
        if not docs_dir:
            QMessageBox.warning(
                self, "Документы",
                "Укажите папку для сохранения документов.",
            )
            return
        dst_dir = Path(docs_dir)
        if not dst_dir.is_dir():
            QMessageBox.warning(
                self, "Документы",
                f"Папка не найдена:\n{docs_dir}",
            )
            return
        src = Path(src_path)
        dst = dst_dir / src.name
        if dst.exists():
            counter = 1
            while True:
                candidate = dst_dir / f"{src.stem} ({counter}){src.suffix}"
                if not candidate.exists():
                    dst = candidate
                    break
                counter += 1
        try:
            shutil.copy2(str(src), str(dst))
            self._refresh_docs_list()
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Документы", f"Не удалось скопировать файл:\n{e}")

    def _on_doc_renamed(self, item: QListWidgetItem) -> None:
        """Переименовывает файл на диске при изменении имени в списке."""
        old_path_str = item.data(Qt.ItemDataRole.UserRole)
        if not old_path_str:
            return
        old_path = Path(old_path_str)
        new_name = item.text().strip()
        if not new_name or new_name == old_path.name:
            return
        new_path = old_path.parent / new_name
        self._close_preview_pdf()
        try:
            old_path.rename(new_path)
            item.setData(Qt.ItemDataRole.UserRole, str(new_path))
            self._preview_current_path = str(new_path)
            self._preview_rename_edit.setText(new_name)
            self._show_preview(str(new_path))
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Переименование", f"Не удалось переименовать файл:\n{e}")
            self._docs_list.blockSignals(True)
            item.setText(old_path.name)
            self._docs_list.blockSignals(False)
            if old_path.is_file():
                self._show_preview(old_path_str)

    # ── Панель предпросмотра ─────────────────────────────────────────────────

    def _build_preview_panel(self) -> QWidget:
        panel = QWidget()
        vbox = QVBoxLayout(panel)
        vbox.setContentsMargins(4, 0, 0, 0)
        vbox.setSpacing(4)
        c0 = ThemeManager.instance().colors

        # Строка переименования
        rename_row = QHBoxLayout()
        rename_row.setSpacing(4)
        name_label = QLabel("Имя:")
        name_label.setStyleSheet(
            f"color: {c0.text_secondary}; font-size: 11px; font-weight: 600;"
        )
        self._preview_name_label = name_label
        self._preview_rename_edit = QLineEdit()
        self._preview_rename_edit.setPlaceholderText("Выберите файл…")
        self._preview_rename_edit.setStyleSheet(f"""
QLineEdit {{
    background-color: {c0.bg_input};
    border: 1px solid {c0.border_input};
    border-radius: 4px;
    padding: 1px 6px;
    font-size: 12px;
    color: {c0.text_primary};
}}
QLineEdit:focus {{
    border-color: {c0.border_input_focus};
    background-color: {c0.bg_input_focus};
}}
""")

        self._preview_rename_btn = QPushButton("Переименовать")
        self._preview_rename_btn.setStyleSheet(self._rename_btn_style(c0))
        self._preview_rename_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._preview_rename_btn.setEnabled(False)

        rename_row.addWidget(name_label)
        rename_row.addWidget(self._preview_rename_edit, 1)
        rename_row.addWidget(self._preview_rename_btn)
        vbox.addLayout(rename_row)

        # Область предпросмотра
        self._preview_scroll = QScrollArea()
        self._preview_scroll.setWidgetResizable(True)
        self._preview_scroll.setFrameShape(QFrame.Shape.StyledPanel)
        self._preview_scroll.setStyleSheet(self._preview_scroll_style(c0))

        self._preview_label = QLabel("Выберите файл\nдля предпросмотра")
        self._preview_label.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter
        )
        self._preview_label.setStyleSheet(
            f"background: transparent; color: {c0.text_muted}; font-size: 12px; padding: 40px;"
        )
        self._preview_label.setWordWrap(True)
        self._preview_scroll.setWidget(self._preview_label)
        vbox.addWidget(self._preview_scroll, 1)

        # Навигация по страницам (для многостраничных PDF)
        self._preview_nav = QWidget()
        nav_h = QHBoxLayout(self._preview_nav)
        nav_h.setContentsMargins(0, 2, 0, 0)
        nav_h.setSpacing(6)

        self._preview_prev_btn = QPushButton("← Пред.")
        self._preview_next_btn = QPushButton("След. →")
        _nav_btn_css = self._nav_btn_style(c0)
        self._preview_prev_btn.setStyleSheet(_nav_btn_css)
        self._preview_next_btn.setStyleSheet(_nav_btn_css)
        self._preview_prev_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._preview_next_btn.setCursor(Qt.CursorShape.PointingHandCursor)

        self._preview_page_label = QLabel("")
        self._preview_page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._preview_page_label.setStyleSheet(
            f"color: {c0.text_secondary}; font-size: 10px; min-width: 100px;"
        )

        nav_h.addStretch()
        nav_h.addWidget(self._preview_prev_btn)
        nav_h.addWidget(self._preview_page_label)
        nav_h.addWidget(self._preview_next_btn)
        nav_h.addStretch()
        self._preview_nav.hide()
        vbox.addWidget(self._preview_nav)

        # Состояние предпросмотра
        self._preview_original_pixmap: QPixmap | None = None
        self._preview_pdf_doc = None
        self._preview_pdf_page: int = 0
        self._preview_pdf_total: int = 0
        self._preview_current_path: str = ""

        # Сигналы панели предпросмотра
        self._preview_rename_btn.clicked.connect(self._rename_from_preview)
        self._preview_rename_edit.returnPressed.connect(self._rename_from_preview)
        self._preview_prev_btn.clicked.connect(self._preview_prev_page)
        self._preview_next_btn.clicked.connect(self._preview_next_page)

        # Масштабирование предпросмотра при изменении размера
        self._preview_scroll.viewport().installEventFilter(self)

        return panel

    # ── Предпросмотр: выбор и отображение ────────────────────────────────────

    def _on_doc_selected(
        self,
        current: QListWidgetItem | None,
        _previous: QListWidgetItem | None,
    ) -> None:
        if current is None:
            self._clear_preview()
            return
        file_path = current.data(Qt.ItemDataRole.UserRole)
        if not file_path or not Path(file_path).is_file():
            self._clear_preview()
            return
        self._preview_rename_edit.setText(Path(file_path).name)
        self._preview_rename_btn.setEnabled(True)
        self._preview_current_path = file_path
        self._show_preview(file_path)

    def _show_preview(self, file_path: str) -> None:
        self._close_preview_pdf()
        suffix = Path(file_path).suffix.lower()

        if suffix == ".pdf":
            self._show_pdf_preview(file_path)
        elif suffix in {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp",
            ".ico", ".webp", ".tiff", ".tif",
        }:
            self._show_image_preview(file_path)
        else:
            self._preview_original_pixmap = None
            self._preview_nav.hide()
            self._preview_label.setPixmap(QPixmap())
            self._preview_label.setStyleSheet(
                "background: transparent; color: #8a9aaa; font-size: 12px; padding: 40px;"
            )
            self._preview_label.setText(
                f"Предпросмотр недоступен\nдля файлов {suffix}"
                if suffix
                else "Предпросмотр недоступен"
            )

    def _show_pdf_preview(self, file_path: str) -> None:
        if not _HAS_FITZ:
            self._preview_original_pixmap = None
            self._preview_nav.hide()
            self._preview_label.setPixmap(QPixmap())
            self._preview_label.setStyleSheet(
                "background: transparent; color: #8a9aaa; font-size: 12px; padding: 40px;"
            )
            self._preview_label.setText(
                "Для предпросмотра PDF установите PyMuPDF:\npip install PyMuPDF"
            )
            return

        try:
            doc = fitz.open(file_path)  # type: ignore[union-attr]
            self._preview_pdf_doc = doc
            self._preview_pdf_page = 0
            self._preview_pdf_total = len(doc)
            self._render_pdf_page()
            if self._preview_pdf_total > 1:
                self._preview_nav.show()
                self._update_page_nav()
            else:
                self._preview_nav.hide()
        except Exception as e:  # noqa: BLE001
            self._preview_original_pixmap = None
            self._preview_nav.hide()
            self._preview_label.setPixmap(QPixmap())
            self._preview_label.setStyleSheet(
                "background: transparent; color: #c44; font-size: 11px; padding: 40px;"
            )
            self._preview_label.setText(f"Ошибка открытия PDF:\n{e}")

    def _render_pdf_page(self) -> None:
        if self._preview_pdf_doc is None:
            return
        page = self._preview_pdf_doc[self._preview_pdf_page]
        mat = fitz.Matrix(2.0, 2.0)  # type: ignore[union-attr]
        pix = page.get_pixmap(matrix=mat, alpha=False)

        img = QImage(pix.samples, pix.width, pix.height, pix.stride,
                     QImage.Format.Format_RGB888)
        pixmap = QPixmap.fromImage(img)

        self._preview_original_pixmap = pixmap
        self._scale_preview()
        self._update_page_nav()

    def _show_image_preview(self, file_path: str) -> None:
        self._preview_nav.hide()
        pixmap = QPixmap(file_path)
        if pixmap.isNull():
            self._preview_original_pixmap = None
            self._preview_label.setPixmap(QPixmap())
            self._preview_label.setStyleSheet(
                "background: transparent; color: #c44; font-size: 11px; padding: 40px;"
            )
            self._preview_label.setText("Не удалось загрузить изображение")
            return
        self._preview_original_pixmap = pixmap
        self._scale_preview()

    # ── Предпросмотр: масштабирование и навигация ────────────────────────────

    def _scale_preview(self) -> None:
        if self._preview_original_pixmap is None or self._preview_original_pixmap.isNull():
            return
        vp_w = self._preview_scroll.viewport().width() - 16
        if vp_w <= 0:
            vp_w = 400
        if self._preview_original_pixmap.width() > vp_w:
            scaled = self._preview_original_pixmap.scaledToWidth(
                vp_w, Qt.TransformationMode.SmoothTransformation,
            )
        else:
            scaled = self._preview_original_pixmap
        self._preview_label.setStyleSheet("background: transparent; padding: 4px;")
        self._preview_label.setText("")
        self._preview_label.setPixmap(scaled)
        self._preview_label.setMinimumHeight(scaled.height())

    def _update_page_nav(self) -> None:
        self._preview_page_label.setText(
            f"Страница {self._preview_pdf_page + 1} из {self._preview_pdf_total}"
        )
        self._preview_prev_btn.setEnabled(self._preview_pdf_page > 0)
        self._preview_next_btn.setEnabled(
            self._preview_pdf_page < self._preview_pdf_total - 1
        )

    def _preview_prev_page(self) -> None:
        if self._preview_pdf_page > 0:
            self._preview_pdf_page -= 1
            self._render_pdf_page()
            self._preview_scroll.verticalScrollBar().setValue(0)

    def _preview_next_page(self) -> None:
        if self._preview_pdf_page < self._preview_pdf_total - 1:
            self._preview_pdf_page += 1
            self._render_pdf_page()
            self._preview_scroll.verticalScrollBar().setValue(0)

    def _clear_preview(self) -> None:
        self._close_preview_pdf()
        self._preview_original_pixmap = None
        self._preview_current_path = ""
        self._preview_rename_edit.setText("")
        self._preview_rename_btn.setEnabled(False)
        self._preview_nav.hide()
        self._preview_label.setPixmap(QPixmap())
        self._preview_label.setStyleSheet(
            "background: transparent; color: #8a9aaa; font-size: 12px; padding: 40px;"
        )
        self._preview_label.setText("Выберите файл\nдля предпросмотра")

    def _close_preview_pdf(self) -> None:
        if self._preview_pdf_doc is not None:
            try:
                self._preview_pdf_doc.close()
            except Exception:  # noqa: BLE001
                pass
            self._preview_pdf_doc = None
        self._preview_pdf_page = 0
        self._preview_pdf_total = 0

    # ── Предпросмотр: переименование ─────────────────────────────────────────

    def _rename_from_preview(self) -> None:
        item = self._docs_list.currentItem()
        if item is None:
            return
        old_path_str = item.data(Qt.ItemDataRole.UserRole)
        if not old_path_str:
            return
        old_path = Path(old_path_str)
        new_name = self._preview_rename_edit.text().strip()
        if not new_name or new_name == old_path.name:
            return
        new_path = old_path.parent / new_name
        if new_path.exists():
            QMessageBox.warning(
                self, "Переименование",
                f"Файл с именем «{new_name}» уже существует.",
            )
            return

        self._close_preview_pdf()
        try:
            old_path.rename(new_path)
            self._docs_list.blockSignals(True)
            item.setText(new_name)
            item.setData(Qt.ItemDataRole.UserRole, str(new_path))
            self._docs_list.blockSignals(False)
            self._preview_current_path = str(new_path)
            self._show_preview(str(new_path))
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(
                self, "Переименование",
                f"Не удалось переименовать файл:\n{e}",
            )
            self._preview_rename_edit.setText(old_path.name)
            if old_path.is_file():
                self._show_preview(old_path_str)

    def _browse_docs_dir(self) -> None:
        """Открывает диалог выбора папки для документов."""
        current = self._docs_path_edit.text().strip()
        path = QFileDialog.getExistingDirectory(
            self, "Выбор папки с документами",
            current if current else str(Path.home()),
        )
        if path:
            self._docs_path_edit.setText(path)
            self._on_docs_path_changed()

    def _open_docs_folder(self) -> None:
        """Открывает папку документов в проводнике."""
        path = self._docs_path_edit.text().strip()
        if not path:
            self._show_status("Укажите путь к папке с документами")
            return
        p = Path(path)
        if not p.is_dir():
            QMessageBox.warning(self, "Открыть папку", f"Папка не найдена:\n{path}")
            return
        try:
            os.startfile(str(p))  # type: ignore[attr-defined]
        except Exception as e:  # noqa: BLE001
            QMessageBox.warning(self, "Открыть папку", f"Не удалось открыть папку:\n{e}")

    def _load_project_docs_path(self) -> None:
        """Подставляет в поле пути папку документов текущего проекта."""
        if self._current is not None:
            path = self._settings.docs_dir
            for key in self._project_docs_keys(self._current):
                saved = self._settings.project_docs_dirs.get(key, "").strip()
                if saved:
                    path = saved
                    break
        else:
            path = ""
        self._docs_path_edit.blockSignals(True)
        self._docs_path_edit.setText(path)
        self._docs_path_edit.blockSignals(False)

    def _on_docs_path_changed(self) -> None:
        """Сохраняет путь к папке документов для текущего проекта и обновляет список."""
        path = self._docs_path_edit.text().strip()
        if self._current is not None:
            for key in self._project_docs_keys(self._current):
                self._settings.project_docs_dirs[key] = path
        try:
            self._settings.save()
        except Exception:  # noqa: BLE001
            pass
        self._refresh_docs_list()

    def _project_docs_keys(self, project: Project) -> list[str]:
        """Возвращает набор стабильных ключей для привязки папки документов."""
        keys: list[str] = []

        pid = (project.project_id or "").strip()
        if pid:
            # legacy key format
            keys.append(pid)
            keys.append(f"id:{pid}")

        row_index = getattr(project, "row_index", None)
        if isinstance(row_index, int) and row_index > 1:
            keys.append(f"row:{row_index}")

        for case_key in ("Номер осн. дела", "Номер дела", "№ дела", "№дела"):
            case_num = str(project.fields.get(case_key, "")).strip()
            if case_num:
                keys.append(f"case:{case_num}")

        # Удаляем дубликаты, сохраняя порядок.
        return list(dict.fromkeys(keys))
