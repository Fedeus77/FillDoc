from __future__ import annotations

from filldoc.excel.models import Project
from filldoc.projects.repository import ProjectRepository


class ArchiveService:
    """Thin service around Excel archive operations."""

    def __init__(self, excel_path: str) -> None:
        self._repository = ProjectRepository(excel_path)

    def move_to_archive(self, project: Project) -> None:
        self._repository.move_to_archive(project)

    def restore_from_archive(self, project: Project) -> None:
        self._repository.restore_from_archive(project)

    def delete_from_archive(self, project: Project) -> None:
        self._repository.delete_from_archive(project)

    def load_archive(self, sheet_name: str = "Архив") -> list[Project]:
        return self._repository.load_archive(sheet_name)

    def repair_headers(self, sheet_name: str = "Архив") -> bool:
        return self._repository.repair_archive_headers(sheet_name)
