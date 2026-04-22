from __future__ import annotations

from filldoc.excel.models import Project
from filldoc.projects.docs_paths import (
    project_docs_keys,
    remember_project_docs_path,
    resolve_project_docs_path,
)


def test_project_docs_keys_keep_legacy_and_stable_keys() -> None:
    project = Project(
        project_id="case-1",
        fields={"Номер дела": "A-10", "№ дела": "A-10"},
        row_index=5,
    )

    assert project_docs_keys(project) == [
        "case-1",
        "id:case-1",
        "row:5",
        "case:A-10",
    ]


def test_resolve_project_docs_path_prefers_saved_project_path() -> None:
    project = Project(project_id="case-1", fields={}, row_index=2)

    assert resolve_project_docs_path(
        project,
        default_docs_dir="C:/default",
        project_docs_dirs={"row:2": "C:/project-docs"},
    ) == "C:/project-docs"


def test_remember_project_docs_path_writes_all_project_keys() -> None:
    project = Project(project_id="case-1", fields={"Номер осн. дела": "A-10"}, row_index=2)
    mapping: dict[str, str] = {}

    remember_project_docs_path(project, "C:/docs", mapping)

    assert mapping == {
        "case-1": "C:/docs",
        "id:case-1": "C:/docs",
        "row:2": "C:/docs",
        "case:A-10": "C:/docs",
    }

