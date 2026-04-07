from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Property, QEasingCurve, QPropertyAnimation, QRect, Signal, Qt
from PySide6.QtGui import QColor, QPainter, QPainterPath
from PySide6.QtWidgets import (
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from filldoc.core.settings import AppSettings
from filldoc.ui.theme import ThemeColors, ThemeManager


# ── Переключатель тёмной/светлой темы ────────────────────────────────────────

class _ThemeToggle(QWidget):
    """Анимированный тогл переключения темы (TickTick-style)."""

    toggled = Signal(bool)  # True = dark

    _TRACK_W = 52
    _TRACK_H = 28
    _THUMB_D = 22

    def __init__(self, is_dark: bool = True, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setFixedSize(self._TRACK_W, self._TRACK_H)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        self._dark = is_dark
        self._anim_progress: float = 1.0 if is_dark else 0.0

        self._anim = QPropertyAnimation(self, b"anim_progress", self)
        self._anim.setDuration(200)
        self._anim.setEasingCurve(QEasingCurve.Type.OutCubic)

    # ── Property для анимации ─────────────────────────────────────────────────

    def _get_anim_progress(self) -> float:
        return self._anim_progress

    def _set_anim_progress(self, v: float) -> None:
        self._anim_progress = v
        self.update()

    anim_progress = Property(float, _get_anim_progress, _set_anim_progress)

    # ── State ─────────────────────────────────────────────────────────────────

    @property
    def is_dark(self) -> bool:
        return self._dark

    def set_dark(self, dark: bool) -> None:
        if self._dark == dark:
            return
        self._dark = dark
        target = 1.0 if dark else 0.0
        self._anim.stop()
        self._anim.setStartValue(self._anim_progress)
        self._anim.setEndValue(target)
        self._anim.start()

    def mousePressEvent(self, _event) -> None:  # noqa: ANN001
        new_dark = not self._dark
        self.set_dark(new_dark)
        self.toggled.emit(new_dark)

    # ── Отрисовка ─────────────────────────────────────────────────────────────

    def paintEvent(self, _event) -> None:  # noqa: ANN001
        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        w, h = self._TRACK_W, self._TRACK_H
        r = h / 2

        # Интерполяция цвета трека
        t = self._anim_progress
        dark_track = QColor("#4ea6ff")
        light_track = QColor("#c0c8d4")
        track_color = QColor(
            int(light_track.red()   + (dark_track.red()   - light_track.red())   * t),
            int(light_track.green() + (dark_track.green() - light_track.green()) * t),
            int(light_track.blue()  + (dark_track.blue()  - light_track.blue())  * t),
        )

        # Трек
        path = QPainterPath()
        path.addRoundedRect(0, 0, w, h, r, r)
        p.fillPath(path, track_color)

        # Иконки луна/солнце внутри трека
        p.setPen(QColor(255, 255, 255, 200))
        p.setFont(self.font())
        # Солнце (левая сторона — светлая тема)
        sun_alpha = int((1.0 - t) * 200)
        p.setPen(QColor(255, 255, 255, sun_alpha))
        p.drawText(QRect(4, 0, 20, h), Qt.AlignmentFlag.AlignCenter, "☀")
        # Луна (правая сторона — тёмная тема)
        moon_alpha = int(t * 200)
        p.setPen(QColor(255, 255, 255, moon_alpha))
        p.drawText(QRect(w - 24, 0, 20, h), Qt.AlignmentFlag.AlignCenter, "🌙")

        # Ползунок
        pad = (h - self._THUMB_D) // 2
        travel = w - self._THUMB_D - 2 * pad
        thumb_x = int(pad + travel * t)
        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(QColor("#ffffff"))
        p.drawEllipse(thumb_x, pad, self._THUMB_D, self._THUMB_D)

        p.end()


# ── Главная вкладка настроек ──────────────────────────────────────────────────

class SettingsTab(QWidget):
    settings_changed = Signal()
    theme_changed = Signal(str)   # "dark" | "light"

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings.load()

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 24, 24, 24)
        root.setSpacing(0)

        # ── Заголовок ─────────────────────────────────────────────────────────
        title = QLabel("Настройки")
        title.setObjectName("settings_title")
        root.addWidget(title)
        root.addSpacing(20)

        # ── Секция: внешний вид ───────────────────────────────────────────────
        appearance_label = QLabel("Внешний вид")
        appearance_label.setObjectName("settings_section_label")
        root.addWidget(appearance_label)
        root.addSpacing(8)

        appearance_card = QWidget()
        appearance_card.setObjectName("settings_card")
        appearance_lay = QVBoxLayout(appearance_card)
        appearance_lay.setContentsMargins(16, 12, 16, 12)
        appearance_lay.setSpacing(0)

        theme_row = QHBoxLayout()
        theme_row.setSpacing(12)
        theme_label = QLabel("Тёмная тема")
        theme_label.setObjectName("settings_row_label")

        self._toggle = _ThemeToggle(is_dark=(self._settings.theme == "dark"))
        self._toggle.toggled.connect(self._on_theme_toggled)

        self._theme_status = QLabel("Тёмная" if self._settings.theme == "dark" else "Светлая")
        self._theme_status.setObjectName("settings_row_hint")

        theme_row.addWidget(theme_label)
        theme_row.addStretch(1)
        theme_row.addWidget(self._theme_status)
        theme_row.addWidget(self._toggle)
        appearance_lay.addLayout(theme_row)
        root.addWidget(appearance_card)
        root.addSpacing(20)

        # ── Секция: пути ──────────────────────────────────────────────────────
        paths_label = QLabel("Пути к файлам и папкам")
        paths_label.setObjectName("settings_section_label")
        root.addWidget(paths_label)
        root.addSpacing(8)

        paths_card = QWidget()
        paths_card.setObjectName("settings_card")
        paths_lay = QVBoxLayout(paths_card)
        paths_lay.setContentsMargins(16, 12, 16, 12)
        paths_lay.setSpacing(8)

        form = QFormLayout()
        form.setSpacing(10)
        form.setContentsMargins(0, 0, 0, 0)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

        self.excel_edit = QLineEdit(self._settings.excel_path)
        self.templates_edit = QLineEdit(self._settings.templates_dir)
        self.output_edit = QLineEdit(self._settings.output_dir)
        self.docs_dir_edit = QLineEdit(self._settings.docs_dir)

        form.addRow("Excel-файл проектов:", self._path_row(self.excel_edit, is_file=True))
        form.addRow("Папка шаблонов:", self._path_row(self.templates_edit, is_file=False))
        form.addRow("Папка выгрузки:", self._path_row(self.output_edit, is_file=False))
        form.addRow("Папка документов:", self._path_row(self.docs_dir_edit, is_file=False))
        paths_lay.addLayout(form)
        root.addWidget(paths_card)
        root.addSpacing(16)

        # ── Кнопки ────────────────────────────────────────────────────────────
        btns = QHBoxLayout()
        btns.setSpacing(10)

        self.save_btn = QPushButton("Сохранить настройки")
        self.save_btn.setObjectName("settings_primary_btn")
        self.check_btn = QPushButton("Проверить пути")
        self.check_btn.setObjectName("settings_secondary_btn")
        btns.addWidget(self.save_btn)
        btns.addWidget(self.check_btn)
        btns.addStretch(1)
        root.addLayout(btns)

        root.addStretch(1)

        self.save_btn.clicked.connect(self._save)
        self.check_btn.clicked.connect(self._check)

    def _path_row(self, edit: QLineEdit, is_file: bool) -> QWidget:
        w = QWidget(self)
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(6)
        lay.addWidget(edit, 1)
        b = QPushButton("Выбрать…", w)
        b.setObjectName("settings_browse_btn")
        b.setFixedWidth(90)
        lay.addWidget(b)

        def pick():
            if is_file:
                path, _ = QFileDialog.getOpenFileName(self, "Выбор Excel-файла", str(Path.home()), "Excel (*.xlsx *.xlsm)")
            else:
                path = QFileDialog.getExistingDirectory(self, "Выбор папки", str(Path.home()))
            if path:
                edit.setText(path)

        b.clicked.connect(pick)
        return w

    def _on_theme_toggled(self, is_dark: bool) -> None:
        name = "dark" if is_dark else "light"
        self._settings.theme = name
        self._theme_status.setText("Тёмная" if is_dark else "Светлая")
        self.theme_changed.emit(name)

    def get_settings(self) -> AppSettings:
        return AppSettings(
            excel_path=self.excel_edit.text().strip(),
            templates_dir=self.templates_edit.text().strip(),
            output_dir=self.output_edit.text().strip(),
            docs_dir=self.docs_dir_edit.text().strip(),
            project_docs_dirs=dict(self._settings.project_docs_dirs),
            theme=self._settings.theme,
        )

    def _save(self) -> None:
        self._settings = self.get_settings()
        try:
            on_disk = AppSettings.load()
            self._settings.project_docs_dirs.update(on_disk.project_docs_dirs)
        except Exception:  # noqa: BLE001
            pass
        try:
            self._settings.save()
            self.settings_changed.emit()
            QMessageBox.information(self, "FillDoc", "Настройки сохранены.")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "FillDoc", str(e))

    def _check(self) -> None:
        s = self.get_settings()
        errs = s.validate_paths()
        if errs:
            QMessageBox.warning(self, "Проверка путей", "\n".join(errs))
        else:
            QMessageBox.information(self, "Проверка путей", "Все пути доступны.")

    def apply_theme(self, c: ThemeColors) -> None:
        """Применяет тему к вкладке настроек."""
        # Обновляем переключатель темы
        is_dark = (c.name == "dark")
        self._toggle.set_dark(is_dark)
        self._theme_status.setText("Тёмная" if is_dark else "Светлая")

        self.setStyleSheet(f"""
QWidget#settings_card {{
    background-color: {c.bg_panel};
    border: 1px solid {c.border_base};
    border-radius: 10px;
}}
QLabel#settings_title {{
    color: {c.text_primary};
    font-size: 20px;
    font-weight: 700;
    background: transparent;
}}
QLabel#settings_section_label {{
    color: {c.text_muted};
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    background: transparent;
}}
QLabel#settings_row_label {{
    color: {c.text_primary};
    font-size: 13px;
    background: transparent;
}}
QLabel#settings_row_hint {{
    color: {c.text_muted};
    font-size: 12px;
    background: transparent;
}}
QPushButton#settings_primary_btn {{
    background-color: {c.accent};
    color: {c.accent_text};
    border: none;
    border-radius: 6px;
    padding: 8px 20px;
    font-size: 13px;
    font-weight: 600;
}}
QPushButton#settings_primary_btn:hover {{
    background-color: {c.accent_hover};
}}
QPushButton#settings_primary_btn:pressed {{
    background-color: {c.accent_pressed};
}}
QPushButton#settings_secondary_btn {{
    background-color: {c.bg_hover};
    color: {c.text_primary};
    border: 1px solid {c.border_base};
    border-radius: 6px;
    padding: 8px 20px;
    font-size: 13px;
}}
QPushButton#settings_secondary_btn:hover {{
    background-color: {c.bg_card_hover};
    border-color: {c.border_input_focus};
}}
QPushButton#settings_browse_btn {{
    background-color: {c.bg_hover};
    color: {c.text_secondary};
    border: 1px solid {c.border_base};
    border-radius: 5px;
    padding: 5px 10px;
    font-size: 12px;
}}
QPushButton#settings_browse_btn:hover {{
    background-color: {c.bg_card_hover};
    color: {c.text_primary};
}}
""")
