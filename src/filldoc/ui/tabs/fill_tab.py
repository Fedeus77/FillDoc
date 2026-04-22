from __future__ import annotations

from PySide6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QMessageBox,
    QLabel,
    QComboBox,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QTableWidget,
    QTableWidgetItem,
    QSplitter,
)

from filldoc.core.settings import AppSettings
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project
from filldoc.fill.missing_fields import compute_missing_fields
from filldoc.generator.docx_generator import generate_docx_from_template
from filldoc.templates.scanner import TemplateLibrary
from filldoc.variables.dictionary import default_dictionary


class FillTab(QWidget):
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._settings = AppSettings()
        self._projects: list[Project] = []
        self._templates = []
        self._dict = default_dictionary()

        root = QVBoxLayout(self)

        top = QHBoxLayout()
        root.addLayout(top)
        top.addWidget(QLabel("Проект:"))
        self.project_combo = QComboBox(self)
        top.addWidget(self.project_combo, 1)
        self.reload_btn = QPushButton("Обновить данные (проекты/шаблоны)")
        top.addWidget(self.reload_btn)

        split = QSplitter(self)
        root.addWidget(split, 1)

        left = QWidget(self)
        left_layout = QVBoxLayout(left)
        left_layout.addWidget(QLabel("Шаблоны (можно выбрать несколько):"))
        self.templates_list = QListWidget(self)
        self.templates_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        left_layout.addWidget(self.templates_list, 1)
        self.analyze_btn = QPushButton("Сформировать список полей")
        left_layout.addWidget(self.analyze_btn)
        split.addWidget(left)

        right = QWidget(self)
        r = QVBoxLayout(right)
        r.addWidget(QLabel("Поля: заполненные / требующие заполнения"))
        self.missing_table = QTableWidget(self)
        self.missing_table.setColumnCount(2)
        self.missing_table.setHorizontalHeaderLabels(["Поле", "Значение (ввести)"])
        r.addWidget(self.missing_table, 1)

        self.generate_btn = QPushButton("Заполнить шаблоны (сгенерировать .docx)")
        self.generate_btn.setEnabled(False)
        r.addWidget(self.generate_btn)
        split.addWidget(right)
        split.setSizes([420, 680])

        self.reload_btn.clicked.connect(self._reload_all)
        self.analyze_btn.clicked.connect(self._analyze)
        self.generate_btn.clicked.connect(self._generate)

    def set_settings(self, s: AppSettings) -> None:
        self._settings = s

    def _reload_all(self) -> None:
        errs = self._settings.validate_paths()
        if errs:
            QMessageBox.warning(self, "Заполнение", "\n".join(errs))
            return

        try:
            store = ExcelProjectStore(self._settings.excel_path)
            self._projects = store.load_projects()
            self.project_combo.clear()
            for p in self._projects:
                self.project_combo.addItem(p.project_id)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Заполнение", f"Не удалось загрузить проекты: {e}")
            return

        try:
            lib = TemplateLibrary(self._settings.templates_dir)
            self._templates = lib.scan()
            self.templates_list.clear()
            for t in self._templates:
                it = QListWidgetItem(f"{t.category + ' / ' if t.category else ''}{t.name}")
                it.setData(0, t.path)
                self.templates_list.addItem(it)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Заполнение", f"Не удалось загрузить шаблоны: {e}")
            return

        self.missing_table.setRowCount(0)
        self.generate_btn.setEnabled(False)

    def _current_project(self) -> Project | None:
        pid = self.project_combo.currentText().strip()
        for p in self._projects:
            if p.project_id == pid:
                return p
        return None

    def _selected_template_cards(self):
        paths = []
        for it in self.templates_list.selectedItems():
            p = it.data(0)
            if p:
                paths.append(p)
        cards = [t for t in self._templates if t.path in set(paths)]
        return cards

    def _analyze(self) -> None:
        project = self._current_project()
        if not project:
            QMessageBox.warning(self, "Заполнение", "Не выбран проект.")
            return
        cards = self._selected_template_cards()
        if not cards:
            QMessageBox.warning(self, "Заполнение", "Не выбран ни один шаблон.")
            return

        merged_vars = []
        seen = set()
        for c in cards:
            for v in c.variables_unique:
                if v in seen:
                    continue
                seen.add(v)
                merged_vars.append(v)

        missing, filled = compute_missing_fields(merged_vars, project.fields, self._dict)

        self.missing_table.setRowCount(len(missing))
        for i, mf in enumerate(missing):
            self.missing_table.setItem(i, 0, QTableWidgetItem(mf.display_name))
            self.missing_table.setItem(i, 1, QTableWidgetItem(""))
        self.missing_table.resizeColumnsToContents()

        if missing:
            self.generate_btn.setEnabled(True)  # MVP: допускаем генерацию даже если есть незаполненные (заменятся на пустые)
            QMessageBox.information(
                self,
                "Заполнение",
                f"Найдено переменных: {len(merged_vars)}.\nЗаполнено: {len(filled)}.\nТребует заполнения: {len(missing)}.",
            )
        else:
            self.generate_btn.setEnabled(True)
            QMessageBox.information(self, "Заполнение", "Все поля уже заполнены. Можно формировать документы.")

    def _collect_missing_values(self) -> dict[str, str]:
        mapping: dict[str, str] = {}
        for r in range(self.missing_table.rowCount()):
            k_item = self.missing_table.item(r, 0)
            v_item = self.missing_table.item(r, 1)
            k = (k_item.text() if k_item else "").strip()
            v = (v_item.text() if v_item else "").strip()
            if k:
                mapping[k] = v
        return mapping

    def _generate(self) -> None:
        project = self._current_project()
        if not project:
            QMessageBox.warning(self, "Заполнение", "Не выбран проект.")
            return
        cards = self._selected_template_cards()
        if not cards:
            QMessageBox.warning(self, "Заполнение", "Не выбран ни один шаблон.")
            return

        # Сохраняем введенные значения в проект (и при желании можно потом сохранить в Excel во вкладке “Проекты”)
        extra = self._collect_missing_values()
        for k, v in extra.items():
            if v.strip():
                project.fields[k] = v.strip()

        # Mapping для подстановки: ключи берем как есть (display name), а для шаблонов используем raw переменные.
        # MVP-упрощение: подстановка по exact "{raw}" из шаблона.
        # Поэтому готовим mapping по всем возможным вариантам:
        mapping: dict[str, str] = {}

        # 1) Из проектных полей добавим как "{ключ}".
        for k, v in project.fields.items():
            mapping[str(k)] = str(v or "")

        # 2) Для переменных из шаблонов: если словарь разрешил — добавим "{raw}" -> значение по display_name.
        for c in cards:
            for raw in c.variables_unique:
                entry = self._dict.resolve(raw)
                if entry:
                    val = project.fields.get(entry.display_name) or project.fields.get(entry.technical_name) or ""
                    mapping[raw] = str(val)
                else:
                    # без словаря: попробуем по совпадению ключей
                    mapping[raw] = str(project.fields.get(raw, ""))

        out_files = []
        try:
            for c in cards:
                customer = (project.fields.get("Заказчик") or "").strip()
                debtor = (project.fields.get("Должник") or "").strip()
                # Простое правило имени файла для MVP:
                # {Имя шаблона} - {Заказчик} - {Должник}
                parts = [c.name]
                if customer:
                    parts.append(customer)
                if debtor:
                    parts.append(debtor)
                out_name = " - ".join(parts)
                out = generate_docx_from_template(c.path, self._settings.output_dir, out_name, mapping)
                out_files.append(out)
        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Генерация", str(e))
            return

        QMessageBox.information(self, "Генерация", "Сформировано документов:\n" + "\n".join(out_files))

