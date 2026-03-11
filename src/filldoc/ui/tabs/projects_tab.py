from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QListWidget,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QAbstractItemView,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project


class ProjectsTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings()
        self._projects: list[Project] = []
        self._current: Project | None = None

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        root.addLayout(top)

        self.load_btn = QPushButton("Обновить проекты из Excel")
        self.save_btn = QPushButton("Сохранить проект в Excel")
        top.addWidget(self.load_btn)
        top.addWidget(self.save_btn)
        top.addStretch(1)

        content = QHBoxLayout()
        root.addLayout(content, 1)

        self.list = QListWidget(self)
        self.list.setMinimumWidth(280)
        content.addWidget(self.list, 0)

        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Поле", "Значение"])
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked)
        content.addWidget(self.table, 1)

        self.load_btn.clicked.connect(self._load_projects)
        self.save_btn.clicked.connect(self._save_current)
        self.list.currentRowChanged.connect(self._select_project)

    def _project_display_name(self, project: Project) -> str:
        creditor = project.fields.get("Кредитор", "").strip()
        debtor = project.fields.get("Должник", "").strip()
        if creditor and debtor:
            return f"{creditor} — {debtor}"
        if creditor:
            return creditor
        if debtor:
            return debtor
        return project.project_id

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    def _load_projects(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Проекты", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        try:
            store = ExcelProjectStore(self._settings.excel_path)
            self._projects = store.load_projects()
            self.list.clear()
            for p in self._projects:
                self.list.addItem(self._project_display_name(p))
            if self._projects:
                self.list.setCurrentRow(0)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))

    def _select_project(self, row: int) -> None:
        if row < 0 or row >= len(self._projects):
            self._current = None
            self.table.setRowCount(0)
            return
        self._current = self._projects[row]
        self._render_project(self._current)

    def _render_project(self, project: Project) -> None:
        items = sorted(project.fields.items(), key=lambda kv: kv[0].lower())
        self.table.setRowCount(len(items))
        for i, (k, v) in enumerate(items):
            self.table.setItem(i, 0, QTableWidgetItem(str(k)))
            self.table.setItem(i, 1, QTableWidgetItem(str(v)))
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
        project.fields = fields

    def _save_current(self) -> None:
        if not self._current:
            QMessageBox.warning(self, "Проекты", "Не выбран проект.")
            return
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Проекты", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        try:
            self._read_table_into_project(self._current)
            store = ExcelProjectStore(self._settings.excel_path)
            store.save_project_fields(self._current)
            QMessageBox.information(self, "Проекты", "Изменения сохранены в Excel (с созданием резервной копии).")
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))

