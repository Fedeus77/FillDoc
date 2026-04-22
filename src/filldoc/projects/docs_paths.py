from __future__ import annotations

from collections.abc import Mapping, MutableMapping

from filldoc.excel.models import Project


CASE_NUMBER_KEYS: tuple[str, ...] = (
    "Номер осн. дела",
    "Номер дела",
    "№ дела",
    "№дела",
)


def project_docs_keys(project: Project) -> list[str]:
    """Return stable keys used to bind a documents folder to a project."""
    keys: list[str] = []

    pid = (project.project_id or "").strip()
    if pid:
        keys.append(pid)
        keys.append(f"id:{pid}")

    row_index = getattr(project, "row_index", None)
    if isinstance(row_index, int) and row_index > 1:
        keys.append(f"row:{row_index}")

    for case_key in CASE_NUMBER_KEYS:
        case_num = str(project.fields.get(case_key, "")).strip()
        if case_num:
            keys.append(f"case:{case_num}")

    return list(dict.fromkeys(keys))


def resolve_project_docs_path(
    project: Project | None,
    *,
    default_docs_dir: str,
    project_docs_dirs: Mapping[str, str],
) -> str:
    """Choose the saved documents path for a project, falling back to default."""
    if project is None:
        return ""

    path = default_docs_dir
    for key in project_docs_keys(project):
        saved = project_docs_dirs.get(key, "").strip()
        if saved:
            path = saved
            break
    return path


def remember_project_docs_path(
    project: Project | None,
    path: str,
    project_docs_dirs: MutableMapping[str, str],
) -> None:
    """Persist a documents path under every stable key known for a project."""
    if project is None:
        return

    for key in project_docs_keys(project):
        project_docs_dirs[key] = path

