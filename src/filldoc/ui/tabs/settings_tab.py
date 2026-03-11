from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QFormLayout,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QHBoxLayout,
)

from filldoc.core.settings import AppSettings


class SettingsTab(QWidget):
    settings_changed = Signal()

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings.load()

        root = QVBoxLayout(self)
        form = QFormLayout()
        root.addLayout(form)

        self.excel_edit = QLineEdit(self._settings.excel_path)
        self.templates_edit = QLineEdit(self._settings.templates_dir)
        self.output_edit = QLineEdit(self._settings.output_dir)

        form.addRow("Excel-файл проектов:", self._path_row(self.excel_edit, is_file=True))
        form.addRow("Папка шаблонов:", self._path_row(self.templates_edit, is_file=False))
        form.addRow("Папка выгрузки:", self._path_row(self.output_edit, is_file=False))

        btns = QHBoxLayout()
        root.addLayout(btns)

        self.save_btn = QPushButton("Сохранить настройки")
        self.check_btn = QPushButton("Проверить доступность путей")
        btns.addWidget(self.save_btn)
        btns.addWidget(self.check_btn)
        btns.addStretch(1)

        self.save_btn.clicked.connect(self._save)
        self.check_btn.clicked.connect(self._check)

        root.addStretch(1)

    def _path_row(self, edit: QLineEdit, is_file: bool) -> QWidget:
        w = QWidget(self)
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(edit, 1)
        b = QPushButton("Выбрать…", w)
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

    def get_settings(self) -> AppSettings:
        return AppSettings(
            excel_path=self.excel_edit.text().strip(),
            templates_dir=self.templates_edit.text().strip(),
            output_dir=self.output_edit.text().strip(),
        )

    def _save(self) -> None:
        self._settings = self.get_settings()
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

