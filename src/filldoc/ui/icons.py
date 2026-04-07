"""Общие SVG-иконки и фабрики иконок/кнопок для всего UI FillDoc."""
from __future__ import annotations

from PySide6.QtCore import Qt, QByteArray, QSize
from PySide6.QtGui import QIcon, QPainter, QPixmap
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QToolButton

# ── SVG-иконки ────────────────────────────────────────────────────────────────

SVG_REFRESH = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M1 4v6h6"/>
  <path d="M23 20v-6h-6"/>
  <path d="M20.49 9A9 9 0 0 0 5.64 5.64L1 10m22 4-4.64 4.36A9 9 0 0 1 3.51 15"/>
</svg>"""

SVG_SAVE = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
  <polyline points="17 21 17 13 7 13 7 21"/>
  <polyline points="7 3 7 8 15 8"/>
</svg>"""

SVG_FOLDER = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.0"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/>
</svg>"""

SVG_FOLDER_OPEN = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
</svg>"""

SVG_ADD = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10"/>
  <line x1="12" y1="8" x2="12" y2="16"/>
  <line x1="8"  y1="12" x2="16" y2="12"/>
</svg>"""

SVG_UPLOAD = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <polyline points="16 16 12 12 8 16"/>
  <line x1="12" y1="12" x2="12" y2="21"/>
  <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/>
</svg>"""

SVG_LINK = """
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"
     fill="none" stroke="currentColor" stroke-width="2.2"
     stroke-linecap="round" stroke-linejoin="round">
  <path d="M10 13a5 5 0 0 1 0-7l1-1a5 5 0 0 1 7 7l-1 1"/>
  <path d="M14 11a5 5 0 0 1 0 7l-1 1a5 5 0 0 1-7-7l1-1"/>
</svg>"""

# ── Стиль кнопки-иконки ───────────────────────────────────────────────────────

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

# ── Фабрики ───────────────────────────────────────────────────────────────────

def make_icon(svg_src: str, color: str = "#ffffff", size: int = 18) -> QIcon:
    colored = svg_src.replace("currentColor", color)
    data = QByteArray(colored.encode())
    renderer = QSvgRenderer(data)
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap)
    renderer.render(painter)
    painter.end()
    return QIcon(pixmap)


def icon_btn(
    svg: str,
    tooltip: str,
    icon_color: str = "#ffffff",
    bg: str = "#8A9BB0",
    hover: str = "#6B7F96",
    pressed: str = "#556477",
    icon_size: int = 18,
) -> QToolButton:
    btn = QToolButton()
    btn.setIcon(make_icon(svg, icon_color, icon_size))
    btn.setIconSize(QSize(icon_size, icon_size))
    btn.setToolTip(tooltip)
    btn.setStyleSheet(_BTN_STYLE.format(bg=bg, hover=hover, pressed=pressed))
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def update_icon_btn(
    btn: QToolButton,
    svg: str,
    icon_color: str = "#ffffff",
    bg: str = "#8A9BB0",
    hover: str = "#6B7F96",
    pressed: str = "#556477",
    icon_size: int = 18,
) -> None:
    """Обновляет иконку и стиль существующей кнопки (для смены темы)."""
    btn.setIcon(make_icon(svg, icon_color, icon_size))
    btn.setStyleSheet(_BTN_STYLE.format(bg=bg, hover=hover, pressed=pressed))
