from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QEvent, Qt, Signal
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from filldoc.ui.theme import ThemeColors, ThemeManager

try:
    import fitz  # PyMuPDF

    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False


_IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".ico",
    ".webp",
    ".tiff",
    ".tif",
}


class DocumentPreviewWidget(QWidget):
    rename_requested = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)

        vbox = QVBoxLayout(self)
        vbox.setContentsMargins(4, 0, 0, 0)
        vbox.setSpacing(4)
        c0 = ThemeManager.instance().colors

        rename_row = QHBoxLayout()
        rename_row.setSpacing(4)
        self._name_label = QLabel("Имя:")

        self._rename_edit = QLineEdit()
        self._rename_edit.setPlaceholderText("Выберите файл…")

        self._rename_btn = QPushButton("Переименовать")
        self._rename_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._rename_btn.setEnabled(False)

        rename_row.addWidget(self._name_label)
        rename_row.addWidget(self._rename_edit, 1)
        rename_row.addWidget(self._rename_btn)
        vbox.addLayout(rename_row)

        self._preview_scroll = QScrollArea()
        self._preview_scroll.setWidgetResizable(True)
        self._preview_scroll.setFrameShape(QFrame.Shape.StyledPanel)

        self._preview_label = QLabel("Выберите файл\nдля предпросмотра")
        self._preview_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self._preview_label.setWordWrap(True)
        self._preview_scroll.setWidget(self._preview_label)
        vbox.addWidget(self._preview_scroll, 1)

        self._preview_nav = QWidget()
        nav_h = QHBoxLayout(self._preview_nav)
        nav_h.setContentsMargins(0, 2, 0, 0)
        nav_h.setSpacing(6)

        self._preview_prev_btn = QPushButton("← Пред.")
        self._preview_next_btn = QPushButton("След. →")
        self._preview_prev_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._preview_next_btn.setCursor(Qt.CursorShape.PointingHandCursor)

        self._preview_page_label = QLabel("")
        self._preview_page_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        nav_h.addStretch()
        nav_h.addWidget(self._preview_prev_btn)
        nav_h.addWidget(self._preview_page_label)
        nav_h.addWidget(self._preview_next_btn)
        nav_h.addStretch()
        self._preview_nav.hide()
        vbox.addWidget(self._preview_nav)

        self._preview_original_pixmap: QPixmap | None = None
        self._preview_pdf_doc = None
        self._preview_pdf_page: int = 0
        self._preview_pdf_total: int = 0
        self.current_path: str = ""

        self._rename_btn.clicked.connect(self.rename_requested)
        self._rename_edit.returnPressed.connect(self.rename_requested)
        self._preview_prev_btn.clicked.connect(self._preview_prev_page)
        self._preview_next_btn.clicked.connect(self._preview_next_page)
        self._preview_scroll.viewport().installEventFilter(self)

        self.apply_theme(c0)

    def apply_theme(self, c: ThemeColors | None = None) -> None:
        if c is None:
            c = ThemeManager.instance().colors

        self._preview_scroll.setStyleSheet(self._preview_scroll_style(c))
        self._preview_label.setStyleSheet(
            f"background: transparent; color: {c.text_muted}; font-size: 12px; padding: 40px;"
        )
        self._name_label.setStyleSheet(
            f"color: {c.text_secondary}; font-size: 11px; font-weight: 600;"
        )
        self._rename_edit.setStyleSheet(f"""
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
        nav_btn_css = self._nav_btn_style(c)
        self._rename_btn.setStyleSheet(self._rename_btn_style(c))
        self._preview_prev_btn.setStyleSheet(nav_btn_css)
        self._preview_next_btn.setStyleSheet(nav_btn_css)
        self._preview_page_label.setStyleSheet(
            f"color: {c.text_secondary}; font-size: 10px; min-width: 100px;"
        )

    def set_file(self, file_path: str) -> None:
        self._rename_edit.setText(Path(file_path).name)
        self._rename_btn.setEnabled(True)
        self.current_path = file_path
        self.show_preview(file_path)

    def rename_text(self) -> str:
        return self._rename_edit.text()

    def set_rename_text(self, text: str) -> None:
        self._rename_edit.setText(text)

    def update_after_rename(self, file_path: str) -> None:
        self.current_path = file_path
        self._rename_edit.setText(Path(file_path).name)
        self.show_preview(file_path)

    def show_preview(self, file_path: str) -> None:
        self.close_pdf()
        suffix = Path(file_path).suffix.lower()

        if suffix == ".pdf":
            self._show_pdf_preview(file_path)
        elif suffix in _IMAGE_SUFFIXES:
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

    def clear(self) -> None:
        self.close_pdf()
        self._preview_original_pixmap = None
        self.current_path = ""
        self._rename_edit.setText("")
        self._rename_btn.setEnabled(False)
        self._preview_nav.hide()
        self._preview_label.setPixmap(QPixmap())
        self._preview_label.setStyleSheet(
            "background: transparent; color: #8a9aaa; font-size: 12px; padding: 40px;"
        )
        self._preview_label.setText("Выберите файл\nдля предпросмотра")

    def close_pdf(self) -> None:
        if self._preview_pdf_doc is not None:
            try:
                self._preview_pdf_doc.close()
            except Exception:  # noqa: BLE001
                pass
            self._preview_pdf_doc = None
        self._preview_pdf_page = 0
        self._preview_pdf_total = 0

    def eventFilter(self, obj, event) -> bool:  # noqa: ANN001
        if obj is self._preview_scroll.viewport() and event.type() == QEvent.Type.Resize:
            self._scale_preview()
        return super().eventFilter(obj, event)

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

        img = QImage(
            pix.samples,
            pix.width,
            pix.height,
            pix.stride,
            QImage.Format.Format_RGB888,
        )
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

    def _scale_preview(self) -> None:
        if self._preview_original_pixmap is None or self._preview_original_pixmap.isNull():
            return
        vp_w = self._preview_scroll.viewport().width() - 16
        if vp_w <= 0:
            vp_w = 400
        if self._preview_original_pixmap.width() > vp_w:
            scaled = self._preview_original_pixmap.scaledToWidth(
                vp_w,
                Qt.TransformationMode.SmoothTransformation,
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

    @staticmethod
    def _preview_scroll_style(c: ThemeColors) -> str:
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

    @staticmethod
    def _rename_btn_style(c: ThemeColors) -> str:
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

    @staticmethod
    def _nav_btn_style(c: ThemeColors) -> str:
        return f"""
QPushButton {{
    background: {c.bg_hover};
    border: 1px solid {c.border_base};
    border-radius: 4px;
    padding: 2px 10px;
    font-size: 11px;
    color: {c.text_primary};
    min-height: 22px;
}}
QPushButton:hover {{ background: {c.bg_input_focus}; border-color: {c.border_input_focus}; }}
QPushButton:pressed {{ background: {c.selection_bg}; }}
QPushButton:disabled {{ color: {c.text_muted}; background: {c.bg_input}; }}
"""

