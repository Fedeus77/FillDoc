from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QColor
from PySide6.QtWidgets import (
    QAbstractItemView,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project

_DROP_ACTIVE_STYLE = "QTableWidget { border: 2px dashed #4A90D9; border-radius: 4px; }"


class ProjectsTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings()
        self._projects: list[Project] = []
        self._current: Project | None = None

        self.setAcceptDrops(True)

        root = QVBoxLayout(self)
        top = QHBoxLayout()
        root.addLayout(top)

        self.load_btn = QPushButton("Обновить проекты из Excel")
        self.save_btn = QPushButton("Сохранить изменения")
        self.add_btn = QPushButton("Добавить проект")
        self.delete_btn = QPushButton("Удалить проект")
        top.addWidget(self.load_btn)
        top.addWidget(self.save_btn)
        top.addWidget(self.add_btn)
        top.addWidget(self.delete_btn)
        top.addStretch(1)

        hint = QLabel("Перетащите .json-файл в окно для загрузки проекта")
        hint.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        hint.setStyleSheet("color: #888; font-style: italic; font-size: 11px;")
        top.addWidget(hint)

        content = QHBoxLayout()
        root.addLayout(content, 1)

        self.list = QListWidget(self)
        self.list.setMinimumWidth(280)
        content.addWidget(self.list, 0)

        self.table = QTableWidget(self)
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Поле", "Значение"])
        # Заголовки выравниваем по левому краю
        h0 = self.table.horizontalHeaderItem(0)
        h1 = self.table.horizontalHeaderItem(1)
        if h0 is not None:
            h0.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        if h1 is not None:
            h1.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        # Редактировать можно только значения (вторая колонка)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked)
        content.addWidget(self.table, 1)

        self.load_btn.clicked.connect(self._load_projects)
        self.save_btn.clicked.connect(self._save_all)
        self.add_btn.clicked.connect(self._add_project)
        self.delete_btn.clicked.connect(self._delete_current)
        self.list.currentRowChanged.connect(self._select_project)

    # ------------------------------------------------------------------ drag & drop

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

        project_id = str(data.get("№ дела", "")).strip() or Path(path).stem
        fields = {str(k): str(v) for k, v in data.items()}
        project = Project(
            project_id=project_id,
            fields=fields,
            headers=list(fields.keys()),
        )

        self._projects.append(project)
        self.list.addItem(self._project_display_name(project))
        self.list.setCurrentRow(len(self._projects) - 1)

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
        # Отображаем поля в том же порядке, что и столбцы Excel:
        # сначала идем по заголовкам (headers), затем — по оставшимся полям, если они есть.
        headers = getattr(project, "headers", None)
        items: list[tuple[str, str]] = []
        if headers:
            # Для нового проекта и для проектов с неполным набором полей
            # проходим по всем заголовкам и подставляем пустое значение,
            # если в проекте такого поля ещё нет.
            for h in headers:
                if not h:
                    continue
                items.append((h, project.fields.get(h, "")))
        else:
            items = list(project.fields.items())
        self.table.setRowCount(len(items))
        for i, (k, v) in enumerate(items):
            key_item = QTableWidgetItem(str(k))
            # Запрещаем редактирование названий полей
            key_item.setFlags(key_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(i, 0, key_item)
            value_item = QTableWidgetItem(str(v))
            # Подсветка ключевых полей, пока они пустые
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
        project.fields = fields

    def _add_project(self) -> None:
        # Создаём пустой проект, который появится только в UI,
        # пока пользователь не сохранит его в Excel.
        next_index = len(self._projects) + 1
        project = Project(
            project_id=f"Новый проект {next_index}",
            fields={},
            headers=[h for h in (self._projects[0].headers or [])] if self._projects and self._projects[0].headers else None,
        )
        self._projects.append(project)
        self.list.addItem(self._project_display_name(project))
        self.list.setCurrentRow(len(self._projects) - 1)

    def _delete_current(self) -> None:
        row = self.list.currentRow()
        if row < 0 or row >= len(self._projects):
            QMessageBox.warning(self, "Проекты", "Не выбран проект для удаления.")
            return

        project = self._projects[row]
        title = self._project_display_name(project)
        answer = QMessageBox.question(
            self,
            "Удаление проекта",
            f"Удалить проект:\n{title}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.Cancel,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return

        # Если есть путь к Excel и проект связан со строкой в Excel — удаляем и там.
        if self._settings.excel_path and project.row_index is not None:
            try:
                store = ExcelProjectStore(self._settings.excel_path)
                store.delete_project(project)
            except Exception as e:  # noqa: BLE001
                QMessageBox.critical(self, "Проекты", f"Не удалось удалить проект из Excel:\n{e}")
                return

        # Удаляем из локального списка и UI
        self._projects.pop(row)
        self.list.takeItem(row)

        if self._projects:
            # Выбираем соседний элемент
            new_row = min(row, len(self._projects) - 1)
            self.list.setCurrentRow(new_row)
        else:
            self._current = None
            self.table.setRowCount(0)

    def _save_all(self) -> None:
        if not self._settings.excel_path:
            QMessageBox.warning(self, "Проекты", "Не указан путь к Excel-файлу проектов (см. Настройки).")
            return
        if not self._projects:
            QMessageBox.information(self, "Проекты", "Нет проектов для сохранения.")
            return

        try:
            # Перед сохранением забираем текущие правки из таблицы
            if self._current is not None:
                self._read_table_into_project(self._current)

            store = ExcelProjectStore(self._settings.excel_path)
            store.save_all_projects(self._projects)
            QMessageBox.information(
                self,
                "Проекты",
                "Все изменения синхронизированы с Excel (с созданием резервной копии).",
            )
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Проекты", str(e))

