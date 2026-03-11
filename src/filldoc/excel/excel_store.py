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
                projects.append(Project(project_id=pid, fields=fields))
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
        MVP: запись обратно в Excel реализуем минимально:
        - ищем строку по совпадению в колонке '№ дела'
        - если колонки не существует или проект row:* — не сохраняем (явная ошибка)
        - создаем бэкап перед записью
        """
        if project.project_id.startswith("row:"):
            raise ExcelError("Нельзя сохранить проект без '№ дела': добавьте колонку '№ дела' и заполните значение.")

        path = Path(self.excel_path)
        if not path.exists():
            raise ExcelError("Excel-файл недоступен или не существует.")

        self.create_backup()

        try:
            wb = load_workbook(filename=path, read_only=False, data_only=False)
            ws = wb.worksheets[0]
            headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
            if "№ дела" not in headers:
                raise ExcelError("В Excel не найден обязательный столбец '№ дела'.")

            col_idx = headers.index("№ дела") + 1
            target_row = None
            for r in range(2, ws.max_row + 1):
                v = ws.cell(row=r, column=col_idx).value
                if ("" if v is None else str(v)).strip() == project.project_id:
                    target_row = r
                    break
            if target_row is None:
                raise ExcelError(f"Не найден проект с '№ дела' = {project.project_id}.")

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

