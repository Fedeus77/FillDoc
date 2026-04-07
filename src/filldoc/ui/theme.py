"""Система тем FillDoc — TickTick/Cursor-вдохновлённая тёмная и светлая темы."""
from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ThemeColors:
    # ── Фоны ─────────────────────────────────────────────────────────────────
    bg_window: str          # фон главного окна
    bg_panel: str           # фон панелей/боковых колонок
    bg_card: str            # фон карточки/элемента списка
    bg_card_hover: str      # ховер карточки
    bg_card_selected: str   # выбранная карточка
    bg_card_archived: str   # заархивированная карточка
    bg_input: str           # фон поля ввода
    bg_input_focus: str     # фон поля ввода в фокусе
    bg_header: str          # фон заголовка таблицы/секции
    bg_alternate: str       # альтернативная строка таблицы
    bg_hover: str           # общий ховер-фон
    bg_tab: str             # фон вкладки
    bg_tab_selected: str    # фон активной вкладки
    bg_scrollbar: str       # трек скроллбара
    bg_tooltip: str         # фон тултипа

    # ── Текст ─────────────────────────────────────────────────────────────────
    text_primary: str       # основной текст
    text_secondary: str     # вторичный текст
    text_muted: str         # приглушённый текст
    text_accent: str        # акцентный/ссылочный текст
    text_label: str         # метка формы
    text_placeholder: str   # плейсхолдер

    # ── Границы ───────────────────────────────────────────────────────────────
    border_base: str        # основная граница
    border_light: str       # лёгкая граница
    border_input: str       # граница поля ввода
    border_input_focus: str # граница поля в фокусе
    border_card: str        # граница карточки
    border_card_selected: str  # граница выбранной карточки

    # ── Акцент ────────────────────────────────────────────────────────────────
    accent: str             # основной акцент (кнопки, выделения)
    accent_hover: str       # ховер акцента
    accent_pressed: str     # нажатый акцент
    accent_disabled: str    # выключенный акцент
    accent_text: str        # текст на акцентной кнопке
    accent_stripe: str      # левая полоса выбранного элемента

    # ── Скроллбар ─────────────────────────────────────────────────────────────
    scrollbar_handle: str
    scrollbar_handle_hover: str

    # ── Кнопки-иконки ─────────────────────────────────────────────────────────
    icon_btn_bg: str
    icon_btn_hover: str
    icon_btn_pressed: str
    icon_btn_ghost_hover: str
    icon_color: str         # цвет SVG-иконок

    # ── Опасные действия ──────────────────────────────────────────────────────
    danger: str
    danger_hover_bg: str
    danger_text: str

    # ── Статусы ───────────────────────────────────────────────────────────────
    success: str
    warning: str

    # ── Выделение ─────────────────────────────────────────────────────────────
    selection_bg: str
    selection_text: str

    # ── Разделитель ───────────────────────────────────────────────────────────
    separator: str

    # ── Имя темы ──────────────────────────────────────────────────────────────
    name: str               # "dark" | "light"


# ── Тёмная тема (TickTick/Cursor inspired) ────────────────────────────────────
DARK = ThemeColors(
    name="dark",

    bg_window="#1a1a1c",
    bg_panel="#1e1e21",
    bg_card="#252528",
    bg_card_hover="#2c2c30",
    bg_card_selected="#163356",
    bg_card_archived="#222224",
    bg_input="#2a2a2d",
    bg_input_focus="#2e2e32",
    bg_header="#222225",
    bg_alternate="#262629",
    bg_hover="#2c2c30",
    bg_tab="#1e1e21",
    bg_tab_selected="#2a2a2d",
    bg_scrollbar="#1e1e21",
    bg_tooltip="#3a3a40",

    text_primary="#dcdcdc",
    text_secondary="#9a9aab",
    text_muted="#6a6a7a",
    text_accent="#4ea6ff",
    text_label="#8a8a9a",
    text_placeholder="#555560",

    border_base="#38383e",
    border_light="#2e2e34",
    border_input="#3e3e46",
    border_input_focus="#4ea6ff",
    border_card="#30303a",
    border_card_selected="#2a5d9f",

    accent="#4ea6ff",
    accent_hover="#3a96f0",
    accent_pressed="#2a80d8",
    accent_disabled="#2a4060",
    accent_text="#ffffff",
    accent_stripe="#4ea6ff",

    scrollbar_handle="#42424a",
    scrollbar_handle_hover="#5a5a66",

    icon_btn_bg="#3a3a42",
    icon_btn_hover="#4a4a54",
    icon_btn_pressed="#2e2e36",
    icon_btn_ghost_hover="#2c2c34",
    icon_color="#c0c0cc",

    danger="#ff6b6b",
    danger_hover_bg="#3d1f1f",
    danger_text="#ff8888",

    success="#4ec9b0",
    warning="#e0a84c",

    selection_bg="#163356",
    selection_text="#dcdcdc",

    separator="#2e2e36",
)


# ── Светлая тема ──────────────────────────────────────────────────────────────
LIGHT = ThemeColors(
    name="light",

    bg_window="#f0f2f6",
    bg_panel="#ffffff",
    bg_card="#ffffff",
    bg_card_hover="#f3f7fc",
    bg_card_selected="#e5efff",
    bg_card_archived="#f4f4f6",
    bg_input="#ffffff",
    bg_input_focus="#fafcff",
    bg_header="#eef3f8",
    bg_alternate="#f7faff",
    bg_hover="#f0f4f8",
    bg_tab="#f0f2f6",
    bg_tab_selected="#ffffff",
    bg_scrollbar="#eef2f7",
    bg_tooltip="#ffffff",

    text_primary="#1e2a38",
    text_secondary="#5b6a7a",
    text_muted="#9aa5b4",
    text_accent="#4a90d9",
    text_label="#5b6a7a",
    text_placeholder="#aab8c4",

    border_base="#dde2ea",
    border_light="#e8edf3",
    border_input="#dde2ea",
    border_input_focus="#5b9bd5",
    border_card="#e3e9f1",
    border_card_selected="#b7cbec",

    accent="#4a90d9",
    accent_hover="#357abd",
    accent_pressed="#2a6099",
    accent_disabled="#9bb8d4",
    accent_text="#ffffff",
    accent_stripe="#4a90d9",

    scrollbar_handle="#bcc7d4",
    scrollbar_handle_hover="#8a9bb0",

    icon_btn_bg="#8a9bb0",
    icon_btn_hover="#6b7f96",
    icon_btn_pressed="#556477",
    icon_btn_ghost_hover="#e8edf3",
    icon_color="#ffffff",

    danger="#e05454",
    danger_hover_bg="#fde8e8",
    danger_text="#d0a8a8",

    success="#2d8a4e",
    warning="#b58700",

    selection_bg="#e5efff",
    selection_text="#18345b",

    separator="#dde2ea",
)


def _scrollbar_qss(c: ThemeColors) -> str:
    return f"""
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
QScrollBar::handle:vertical:hover {{
    background: {c.scrollbar_handle_hover};
}}
QScrollBar::add-line:vertical,
QScrollBar::sub-line:vertical {{ height: 0px; }}
QScrollBar::add-page:vertical,
QScrollBar::sub-page:vertical {{ background: none; }}
QScrollBar:horizontal {{
    background: {c.bg_scrollbar};
    height: 6px;
    border-radius: 3px;
    margin: 2px 6px 2px 6px;
}}
QScrollBar::handle:horizontal {{
    background: {c.scrollbar_handle};
    border-radius: 3px;
    min-width: 24px;
}}
QScrollBar::handle:horizontal:hover {{
    background: {c.scrollbar_handle_hover};
}}
QScrollBar::add-line:horizontal,
QScrollBar::sub-line:horizontal {{ width: 0px; }}
QScrollBar::add-page:horizontal,
QScrollBar::sub-page:horizontal {{ background: none; }}
"""


def build_global_stylesheet(c: ThemeColors) -> str:
    """Строит глобальный QSS для всего приложения."""
    scrollbars = _scrollbar_qss(c)
    return f"""
/* ── Окно ── */
QMainWindow, QDialog {{
    background-color: {c.bg_window};
    color: {c.text_primary};
}}
QWidget {{
    background-color: transparent;
    color: {c.text_primary};
    font-family: "Segoe UI", system-ui, sans-serif;
    font-size: 13px;
}}

/* ── Вкладки ── */
QTabWidget::pane {{
    border: 1px solid {c.border_base};
    background-color: {c.bg_panel};
    border-radius: 8px;
    margin-top: -1px;
}}
QTabBar::tab {{
    background-color: {c.bg_tab};
    color: {c.text_secondary};
    border: 1px solid transparent;
    border-bottom: none;
    padding: 8px 20px;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    margin-right: 2px;
    font-size: 13px;
}}
QTabBar::tab:selected {{
    background-color: {c.bg_tab_selected};
    color: {c.text_primary};
    border-color: {c.border_base};
    font-weight: 600;
}}
QTabBar::tab:hover:!selected {{
    background-color: {c.bg_hover};
    color: {c.text_primary};
}}

/* ── Строка статуса ── */
QStatusBar {{
    background-color: {c.bg_panel};
    color: {c.text_muted};
    border-top: 1px solid {c.border_base};
    font-size: 12px;
}}

/* ── Кнопки ── */
QPushButton {{
    background-color: {c.bg_hover};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 6px;
    padding: 7px 16px;
    font-size: 13px;
}}
QPushButton:hover {{
    background-color: {c.bg_card_hover};
    border-color: {c.border_input_focus};
}}
QPushButton:pressed {{
    background-color: {c.bg_hover};
}}
QPushButton:disabled {{
    color: {c.text_muted};
    border-color: {c.border_light};
}}

/* ── Поля ввода ── */
QLineEdit {{
    background-color: {c.bg_input};
    color: {c.text_primary};
    border: 1px solid {c.border_input};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 13px;
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
}}
QLineEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}
QLineEdit:hover {{
    border-color: {c.scrollbar_handle_hover};
}}
QLineEdit::placeholder {{
    color: {c.text_placeholder};
}}

/* ── Текстовые поля ── */
QTextEdit, QPlainTextEdit {{
    background-color: {c.bg_input};
    color: {c.text_primary};
    border: 1px solid {c.border_input};
    border-radius: 6px;
    padding: 4px 8px;
    font-size: 13px;
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
}}
QTextEdit:focus, QPlainTextEdit:focus {{
    border-color: {c.border_input_focus};
    background-color: {c.bg_input_focus};
}}

/* ── Метки ── */
QLabel {{
    color: {c.text_primary};
    background: transparent;
}}

/* ── Комбобокс ── */
QComboBox {{
    background-color: {c.bg_input};
    color: {c.text_primary};
    border: 1px solid {c.border_input};
    border-radius: 6px;
    padding: 6px 32px 6px 10px;
    font-size: 13px;
}}
QComboBox:focus {{
    border-color: {c.border_input_focus};
}}
QComboBox:hover {{
    border-color: {c.scrollbar_handle_hover};
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox::down-arrow {{
    width: 10px;
    height: 10px;
}}
QComboBox QAbstractItemView {{
    background-color: {c.bg_panel};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 6px;
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    padding: 6px 10px;
    min-height: 28px;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {c.bg_hover};
}}

/* ── Таблицы ── */
QTableWidget {{
    background-color: {c.bg_panel};
    alternate-background-color: {c.bg_alternate};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 8px;
    gridline-color: {c.border_light};
    outline: 0;
    selection-background-color: {c.selection_bg};
    selection-color: {c.selection_text};
    font-size: 13px;
}}
QTableWidget::item {{
    padding: 6px 8px;
    border: none;
}}
QTableWidget::item:selected {{
    background-color: {c.selection_bg};
    color: {c.selection_text};
}}
QHeaderView::section {{
    background-color: {c.bg_header};
    color: {c.text_secondary};
    border: none;
    border-bottom: 1px solid {c.border_base};
    padding: 7px 8px;
    font-size: 12px;
    font-weight: 600;
}}
QHeaderView::section:horizontal:first {{
    border-top-left-radius: 8px;
}}
QHeaderView::section:horizontal:last {{
    border-top-right-radius: 8px;
}}

/* ── Списки ── */
QListWidget {{
    background-color: {c.bg_panel};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 8px;
    outline: 0;
    padding: 4px;
}}
QListWidget::item {{
    border-radius: 6px;
    padding: 6px 10px;
}}
QListWidget::item:selected {{
    background-color: {c.selection_bg};
    color: {c.selection_text};
}}
QListWidget::item:hover {{
    background-color: {c.bg_hover};
}}

/* ── Деревья ── */
QTreeWidget {{
    background-color: {c.bg_panel};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 8px;
    outline: 0;
    font-size: 13px;
}}
QTreeWidget::item {{
    padding: 4px 6px;
    border-radius: 4px;
}}
QTreeWidget::item:selected {{
    background-color: {c.selection_bg};
    color: {c.selection_text};
}}
QTreeWidget::item:hover {{
    background-color: {c.bg_hover};
}}
QTreeWidget::branch {{
    background-color: transparent;
}}

/* ── Сплиттер ── */
QSplitter::handle {{
    background-color: {c.border_base};
}}
QSplitter::handle:horizontal {{
    width: 1px;
}}
QSplitter::handle:vertical {{
    height: 1px;
}}

/* ── Тулбар кнопки ── */
QToolButton {{
    background-color: transparent;
    color: {c.text_secondary};
    border: none;
    border-radius: 6px;
    padding: 4px;
}}
QToolButton:hover {{
    background-color: {c.bg_hover};
    color: {c.text_primary};
}}
QToolButton:pressed {{
    background-color: {c.icon_btn_pressed};
}}

/* ── Диалоги ── */
QMessageBox {{
    background-color: {c.bg_panel};
    color: {c.text_primary};
}}
QMessageBox QPushButton {{
    min-width: 80px;
}}

/* ── Форма ── */
QFormLayout QLabel {{
    color: {c.text_label};
    font-size: 13px;
}}

/* ── Тултипы ── */
QToolTip {{
    background-color: {c.bg_tooltip};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 4px;
    padding: 4px 8px;
    font-size: 12px;
}}

/* ── Скроллбары ── */
{scrollbars}
"""


# ── ThemeManager ──────────────────────────────────────────────────────────────

class ThemeManager:
    """Синглтон-менеджер темы приложения."""

    _instance: "ThemeManager | None" = None
    _current: ThemeColors = DARK

    @classmethod
    def instance(cls) -> "ThemeManager":
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    @property
    def colors(self) -> ThemeColors:
        return self._current

    @property
    def is_dark(self) -> bool:
        return self._current.name == "dark"

    def set_theme(self, name: str) -> None:
        """Переключить тему: 'dark' или 'light'."""
        self._current = DARK if name == "dark" else LIGHT

    def get_by_name(self, name: str) -> ThemeColors:
        return DARK if name == "dark" else LIGHT
