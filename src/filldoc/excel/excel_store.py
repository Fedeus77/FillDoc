from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path

from openpyxl import load_workbook

from filldoc.core.errors import ExcelError
from .models import Project


class ExcelProjectStore:
    """
    MVP-договоренности (пока без “гибкой конфигурации”):
    - читаем первую вкладку
    - первая строка = заголовки столбцов
    - проект = строка, где есть хоть одно значение
    - идентификатор проекта: колонка '№ дела' если есть и не пусто, иначе 'row:<номер>'
    """

    def __init__(self, excel_path: str) -> None:
        self.excel_path = excel_path

    def load_projects(self) -> list[Project]:
        path = Path(self.excel_path)
        if not path.exists():
            raise ExcelError("Excel-файл недоступен или не существует.")
        try:
            wb = load_workbook(filename=path, read_only=False, data_only=True)
            ws = wb.worksheets[0]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return []
            headers = [str(h).strip() if h is not None else "" for h in rows[0]]
            projects: list[Project] = []
            for idx, row in enumerate(rows[1:], start=2):
                if row is None:
                    continue
                values = ["" if v is None else str(v) for v in row]
                if all(v.strip() == "" for v in values):
                    continue
                fields = {headers[i]: values[i] for i in range(min(len(headers), len(values))) if headers[i]}
                pid = (fields.get("№ дела") or "").strip()
                if not pid:
                    pid = f"row:{idx}"
                # Сохраняем также headers и номер строки (idx),
                # чтобы потом можно было сохранять изменения без зависимости от "№ дела".
                projects.append(Project(project_id=pid, fields=fields, headers=headers, row_index=idx))
            return projects
        except ExcelError:
            raise
        except Exception as e:  # noqa: BLE001
            raise ExcelError(f"Не удалось загрузить проекты из Excel: {e}") from e

    def create_backup(self) -> Path:
        src = Path(self.excel_path)
        if not src.exists():
            raise ExcelError("Не удалось создать резервную копию: Excel-файл не найден.")
        backup_dir = src.parent / "_filldoc_backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = backup_dir / f"{src.stem}__backup__{ts}{src.suffix}"
        try:
            shutil.copy2(src, dst)
            return dst
        except Exception as e:  # noqa: BLE001
            raise ExcelError(f"Не удалось создать резервную копию Excel-файла: {e}") from e

    def save_project_fields(self, project: Project) -> None:
        """
        Запись обратно в Excel:
        - используем сохранённый номер строки (row_index), чтобы найти нужную строку
        - создаем бэкап перед записью
        """
        path = Path(self.excel_path)
        if not path.exists():
            raise ExcelError("Excel-файл недоступен или не существует.")

        self.create_backup()

        try:
            wb = load_workbook(filename=path, read_only=False, data_only=False)
            ws = wb.worksheets[0]
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
            # Определяем строку, в которую нужно писать.
            # Для проектов, загруженных из Excel, row_index всегда задан.
            target_row = project.row_index
            if target_row is None or target_row < 2 or target_row > ws.max_row:
                raise ExcelError("Не удалось определить строку проекта в Excel для сохранения.")

            header_to_col = {headers[i]: i + 1 for i in range(len(headers)) if headers[i]}
            for k, v in project.fields.items():
                if k not in header_to_col:
                    continue
                ws.cell(row=target_row, column=header_to_col[k]).value = v
            wb.save(path)
        except ExcelError:
            raise
        except Exception as e:  # noqa: BLE001
            raise ExcelError(f"Не удалось сохранить изменения в Excel: {e}") from e

    def add_project(self, project: Project) -> None:
        """Добавляет новую строку в конец Excel-листа и обновляет project.row_index."""
        path = Path(self.excel_path)
        if not path.exists():
            raise ExcelError("Excel-файл недоступен или не существует.")

        self.create_backup()

        try:
            wb = load_workbook(filename=path, read_only=False, data_only=False)
            ws = wb.worksheets[0]
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

            header_to_idx = {h: i for i, h in enumerate(headers) if h}
            new_row: list[str] = [""] * len(headers)
            for k, v in project.fields.items():
                if k in header_to_idx:
                    new_row[header_to_idx[k]] = v

            ws.append(new_row)
            project.row_index = ws.max_row
            project.headers = [h for h in headers if h]
            wb.save(path)
        except ExcelError:
            raise
        except Exception as e:  # noqa: BLE001
            raise ExcelError(f"Не удалось добавить проект в Excel: {e}") from e

    def save_all_projects(self, projects: list[Project]) -> None:
        """
        Полная синхронизация списка проектов с Excel:
        - пересобираем все строки (кроме заголовка) по текущему списку projects
        - обновляем row_index у проектов под новые строки
        """
        path = Path(self.excel_path)
        if not path.exists():
            raise ExcelError("Excel-файл недоступен или не существует.")

        self.create_backup()

        try:
            wb = load_workbook(filename=path, read_only=False, data_only=False)
            ws = wb.worksheets[0]
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]

            # Очищаем все строки, кроме заголовка
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)

            header_to_idx = {h: i for i, h in enumerate(headers) if h}

            # Пересобираем таблицу по текущему списку проектов
            for idx, project in enumerate(projects, start=2):
                row_values: list[str] = [""] * len(headers)
                for k, v in project.fields.items():
                    if k in header_to_idx:
                        row_values[header_to_idx[k]] = v
                ws.append(row_values)
                project.row_index = idx
                project.headers = [h for h in headers if h]

            wb.save(path)
        except ExcelError:
            raise
        except Exception as e:  # noqa: BLE001
            raise ExcelError(f"Не удалось синхронизировать проекты с Excel: {e}") from e

    def delete_project(self, project: Project) -> None:
        """Удаляет строку проекта из Excel по row_index."""
        path = Path(self.excel_path)
        if not path.exists():
            raise ExcelError("Excel-файл недоступен или не существует.")

        if project.row_index is None:
            raise ExcelError("Невозможно удалить проект: не задан номер строки в Excel.")

        self.create_backup()

        try:
            wb = load_workbook(filename=path, read_only=False, data_only=False)
            ws = wb.worksheets[0]
            target_row = project.row_index
            if target_row < 2 or target_row > ws.max_row:
                raise ExcelError("Не удалось определить строку проекта в Excel для удаления.")
            ws.delete_rows(target_row, 1)
            wb.save(path)
        except ExcelError:
            raise
        except Exception as e:  # noqa: BLE001
            raise ExcelError(f"Не удалось удалить проект из Excel: {e}") from e

