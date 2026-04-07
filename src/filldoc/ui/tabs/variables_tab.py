from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

from PySide6.QtCore import Qt, QByteArray, QSize, QTimer
from PySide6.QtGui import QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.ui.theme import ThemeColors, ThemeManager

# ── SVG иконки ───────────────────────────────────────────────────────────────

_SVG_REFRESH = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M1 4v6h6"/>
  <path d="M23 20v-6h-6"/>
  <path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4-4.64 4.36A9 9 0 0 1 3.51 15"/>
</svg>"""

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

# ── Стили ────────────────────────────────────────────────────────────────────

_REFRESH_BTN_STYLE = """
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

_COPY_BTN_STYLE = """
QToolButton {
    background-color: transparent;
    border: none;
    border-radius: 4px;
    padding: 2px;
}
QToolButton:hover {
    background-color: #e8edf5;
}
QToolButton:pressed {
    background-color: #d0daea;
}
"""

_BOTTOM_ICON_BTN_STYLE = """
QToolButton {{
    background-color: transparent;
    border: none;
    border-radius: 9px;
    min-width:  34px;
    min-height: 34px;
    max-width:  34px;
    max-height: 34px;
}}
QToolButton:hover {{
    background-color: #e8edf5;
}}
QToolButton:pressed {{
    background-color: #d0daea;
}}
"""

_VAR_ROW_STYLE = """
QFrame {
    background-color: transparent;
    border: none;
    border-radius: 3px;
}
QFrame:hover {
    background-color: #f4f6f9;
}
"""

_VAR_LABEL_STYLE = """
QLabel {
    font-family: Consolas, 'Courier New', monospace;
    font-size: 12px;
    color: #2a4a7a;
    padding: 0px;
    background: transparent;
}
"""

_DESC_STYLE = """
QLabel {
    font-size: 12px;
    color: #6b7f96;
    padding: 2px 0px 6px 0px;
}
"""

_EMPTY_STYLE = """
QLabel {
    font-size: 13px;
    color: #9aa5b4;
    padding: 24px;
}
"""

# ── Вспомогательные функции ───────────────────────────────────────────────────

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


def _refresh_btn() -> QToolButton:
    btn = QToolButton()
    btn.setIcon(_make_icon(_SVG_REFRESH, "#ffffff", 18))
    btn.setIconSize(QSize(18, 18))
    btn.setToolTip("Обновить список переменных из Excel")
    btn.setStyleSheet(
        _REFRESH_BTN_STYLE.format(bg="#8A9BB0", hover="#6B7F96", pressed="#556477")
    )
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def _open_excel_btn() -> QToolButton:
    btn = QToolButton()
    btn.setIcon(_make_icon(_SVG_EXCEL, "#2d8a4e", 18))
    btn.setIconSize(QSize(18, 18))
    btn.setToolTip("Открыть Excel-файл из настроек")
    btn.setStyleSheet(_BOTTOM_ICON_BTN_STYLE)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


# ── Основной класс ────────────────────────────────────────────────────────────

class VariablesTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self._settings = AppSettings()
        self._headers: list[str] = []

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(4)

        # ── Верхняя панель ────────────────────────────────────────────
        top = QHBoxLayout()
        top.setSpacing(6)

        desc = QLabel("Переменные из таблицы Excel — используйте в шаблонах:")
        desc.setStyleSheet(_DESC_STYLE)
        self._desc_label = desc
        top.addWidget(desc, 1)

        self.open_excel_btn = _open_excel_btn()
        top.addWidget(self.open_excel_btn)

        self.refresh_btn = _refresh_btn()
        top.addWidget(self.refresh_btn)
        root.addLayout(top)

        # ── Разделитель ───────────────────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        sep.setStyleSheet("color: #dde2ea;")
        self._sep = sep
        root.addWidget(sep)

        # ── Список переменных ─────────────────────────────────────────
        self.scroll = QScrollArea(self)
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.Shape.NoFrame)

        self._list_container = QWidget()
        self._list_layout = QVBoxLayout(self._list_container)
        self._list_layout.setContentsMargins(4, 2, 4, 2)
        self._list_layout.setSpacing(0)

        self.scroll.setWidget(self._list_container)
        root.addWidget(self.scroll, 1)

        # ── Начальное состояние ───────────────────────────────────────
        self._show_empty("Нажмите  ↻  для загрузки переменных из Excel")

        self.refresh_btn.clicked.connect(self._reload)
        self.open_excel_btn.clicked.connect(self._open_excel_file)

    # ── Настройки ─────────────────────────────────────────────────────────────

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    def apply_theme(self, c: ThemeColors) -> None:
        """Применяет тему ко всем виджетам вкладки."""
        # Описание
        self._desc_label.setStyleSheet(
            f"QLabel {{ font-size: 12px; color: {c.text_secondary}; padding: 2px 0px 6px 0px; }}"
        )

        # Разделитель
        self._sep.setStyleSheet(f"color: {c.separator};")

        # Кнопки
        self.refresh_btn.setIcon(_make_icon(_SVG_REFRESH, c.icon_color, 18))
        self.refresh_btn.setStyleSheet(
            _REFRESH_BTN_STYLE.format(bg=c.icon_btn_bg, hover=c.icon_btn_hover, pressed=c.icon_btn_pressed)
        )
        self.open_excel_btn.setIcon(_make_icon(_SVG_EXCEL, c.success, 18))
        self.open_excel_btn.setStyleSheet(f"""
QToolButton {{
    background-color: transparent;
    border: none;
    border-radius: 9px;
    min-width: 34px; min-height: 34px;
    max-width: 34px; max-height: 34px;
}}
QToolButton:hover {{ background-color: {c.icon_btn_ghost_hover}; }}
QToolButton:pressed {{ background-color: {c.icon_btn_pressed}; }}
""")

        # Перерисовываем список переменных (с обновлёнными цветами)
        if self._headers:
            self._render_list()

    # ── Загрузка данных ───────────────────────────────────────────────────────

    def _reload(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(
                self, "Переменные",
                "Не указан путь к Excel-файлу (см. Настройки).",
            )
            return
        try:
            store = ExcelProjectStore(self._settings.excel_path)
            projects = store.load_projects()
            if projects:
                self._headers = [h for h in projects[0].headers if h.strip()]
            else:
                self._headers = []
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Переменные", str(e))
            return

        self._render_list()

    def _open_excel_file(self) -> None:
        excel_path = (self._settings.excel_path or "").strip()
        if not excel_path:
            QMessageBox.warning(
                self,
                "Переменные",
                "Не указан путь к Excel-файлу (см. Настройки).",
            )
            return

        path = Path(excel_path)
        if not path.exists() or not path.is_file():
            QMessageBox.warning(
                self,
                "Переменные",
                "Excel-файл недоступен или не существует.",
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
            QMessageBox.critical(self, "Переменные", f"Не удалось открыть Excel-файл:\n{e}")

    # ── Рендер списка ─────────────────────────────────────────────────────────

    def _clear_list(self) -> None:
        while self._list_layout.count():
            child = self._list_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def _show_empty(self, msg: str) -> None:
        self._clear_list()
        c = ThemeManager.instance().colors
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
            row = self._make_var_row(f"{{{header}}}")
            self._list_layout.addWidget(row)

        self._list_layout.addStretch(1)

    def _make_var_row(self, var_text: str) -> QFrame:
        c = ThemeManager.instance().colors
        frame = QFrame()
        frame.setStyleSheet(f"""
QFrame {{ background-color: transparent; border: none; border-radius: 3px; }}
QFrame:hover {{ background-color: {c.bg_hover}; }}
""")
        frame.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        lay = QHBoxLayout(frame)
        lay.setContentsMargins(4, 1, 4, 1)
        lay.setSpacing(4)

        copy_btn = QToolButton()
        copy_btn.setIcon(_make_icon(_SVG_COPY, c.text_muted, 13))
        copy_btn.setIconSize(QSize(13, 13))
        copy_btn.setFixedSize(20, 20)
        copy_btn.setToolTip("Копировать")
        copy_btn.setStyleSheet(f"""
QToolButton {{ background-color: transparent; border: none; border-radius: 4px; padding: 2px; }}
QToolButton:hover {{ background-color: {c.icon_btn_ghost_hover}; }}
QToolButton:pressed {{ background-color: {c.icon_btn_pressed}; }}
""")
        copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_btn.clicked.connect(
            lambda _checked, t=var_text, b=copy_btn: self._copy_var(t, b)
        )
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

    # ── Копирование ───────────────────────────────────────────────────────────

    def _copy_var(self, text: str, btn: QToolButton) -> None:
        c = ThemeManager.instance().colors
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
