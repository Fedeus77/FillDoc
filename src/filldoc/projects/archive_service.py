from __future__ import annotations

from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project


class ArchiveService:
    """Thin service around Excel archive operations."""

    def __init__(self, excel_path: str) -> None:
        self._store = ExcelProjectStore(excel_path)

    def move_to_archive(self, project: Project) -> None:
        self._store.move_project_to_archive(project)

    def restore_from_archive(self, project: Project) -> None:
        self._store.restore_project_from_archive(project)

    def delete_from_archive(self, project: Project) -> None:
        self._store.delete_project_from_archive(project)

    def load_archive(self, sheet_name: str = "Архив") -> list[Project]:
        return self._store.load_projects_from_sheet(sheet_name)

    def repair_headers(self, sheet_name: str = "Архив") -> bool:
        return self._store.repair_archive_headers(sheet_name)

