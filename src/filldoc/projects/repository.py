from __future__ import annotations

from dataclasses import dataclass

from filldoc.core.errors import ExcelError
from filldoc.excel.excel_store import ExcelProjectStore
from filldoc.excel.models import Project


ARCHIVE_SHEET_NAME = "Архив"


@dataclass(frozen=True)
class ProjectConflict:
    project: Project
    sheet_name: str | None
    reason: str


class ProjectConflictError(ExcelError):
    def __init__(self, conflicts: list[ProjectConflict]) -> None:
        self.conflicts = conflicts
        super().__init__("Excel-файл изменился снаружи после загрузки проектов.")


class ProjectRepository:
    """Application-facing project storage API.

    UI code should use this class instead of talking to ExcelProjectStore
    directly. ExcelProjectStore remains responsible for workbook mechanics.
    """

    def __init__(self, excel_path: str) -> None:
        self._store = ExcelProjectStore(excel_path)

    def load_projects(self) -> list[Project]:
        return self._store.load_projects()

    def load_archive(self, sheet_name: str = ARCHIVE_SHEET_NAME) -> list[Project]:
        return self._store.load_projects_from_sheet(sheet_name)

    def repair_archive_headers(self, sheet_name: str = ARCHIVE_SHEET_NAME) -> bool:
        return self._store.repair_archive_headers(sheet_name)

    def save_project_fields(self, project: Project, *, force: bool = False) -> None:
        conflicts = self.find_conflicts([project]) if not force else []
        if conflicts:
            raise ProjectConflictError(conflicts)
        self._store.save_project_fields(project)

    def save_all_projects(
        self,
        active_projects: list[Project],
        archived_projects: list[Project] | None = None,
        *,
        force: bool = False,
    ) -> None:
        conflicts = [] if force else self.find_conflicts(active_projects, archived_projects)
        if conflicts:
            raise ProjectConflictError(conflicts)
        self._store.save_all_projects(active_projects, archived_projects)

    def move_to_archive(self, project: Project) -> None:
        self._store.move_project_to_archive(project)

    def restore_from_archive(self, project: Project) -> None:
        self._store.restore_project_from_archive(project)

    def delete_project(self, project: Project) -> None:
        self._store.delete_project(project)

    def delete_from_archive(self, project: Project) -> None:
        self._store.delete_project_from_archive(project)

    def find_conflicts(
        self,
        active_projects: list[Project],
        archived_projects: list[Project] | None = None,
    ) -> list[ProjectConflict]:
        conflicts: list[ProjectConflict] = []
        conflicts.extend(self._find_project_conflicts(active_projects, sheet_name=None))
        if archived_projects is not None:
            conflicts.extend(self._find_project_conflicts(archived_projects, sheet_name=ARCHIVE_SHEET_NAME))
        return conflicts

    def _find_project_conflicts(
        self,
        projects: list[Project],
        *,
        sheet_name: str | None,
    ) -> list[ProjectConflict]:
        conflicts: list[ProjectConflict] = []
        for project in projects:
            if project.loaded_snapshot is None:
                continue
            current_snapshot = self._store.current_project_snapshot(project, sheet_name=sheet_name)
            if current_snapshot is None:
                conflicts.append(ProjectConflict(project=project, sheet_name=sheet_name, reason="missing"))
            elif current_snapshot != project.loaded_snapshot:
                conflicts.append(ProjectConflict(project=project, sheet_name=sheet_name, reason="changed"))
        return conflicts
