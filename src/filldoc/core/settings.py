from __future__ import annotations

import json
from dataclasses import dataclass, asdict, field
from pathlib import Path

from .errors import SettingsError


def _settings_dir() -> Path:
    return Path.home() / ".filldoc"


def _settings_path() -> Path:
    return _settings_dir() / "settings.json"


@dataclass
class AppSettings:
    excel_path: str = ""
    templates_dir: str = ""
    output_dir: str = ""
    docs_dir: str = ""
    # Пути к папкам документов для каждого проекта: {project_id: path}
    project_docs_dirs: dict = field(default_factory=dict)

    @staticmethod
    def load() -> "AppSettings":
        path = _settings_path()
        if not path.exists():
            return AppSettings()
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            raw_dirs = data.get("project_docs_dirs", {})
            return AppSettings(
                excel_path=str(data.get("excel_path", "")),
                templates_dir=str(data.get("templates_dir", "")),
                output_dir=str(data.get("output_dir", "")),
                docs_dir=str(data.get("docs_dir", "")),
                project_docs_dirs=raw_dirs if isinstance(raw_dirs, dict) else {},
            )
        except Exception as e:  # noqa: BLE001
            raise SettingsError(f"Не удалось прочитать настройки: {e}") from e

    def save(self) -> None:
        try:
            _settings_dir().mkdir(parents=True, exist_ok=True)
            _settings_path().write_text(
                json.dumps(asdict(self), ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:  # noqa: BLE001
            raise SettingsError(f"Не удалось сохранить настройки: {e}") from e

    def validate_paths(self) -> list[str]:
        errors: list[str] = []
        if not self.excel_path:
            errors.append("Не указан путь к Excel-файлу проектов.")
        elif not Path(self.excel_path).exists():
            errors.append("Excel-файл недоступен или не существует.")

        if not self.templates_dir:
            errors.append("Не указан путь к библиотеке шаблонов.")
        elif not Path(self.templates_dir).exists():
            errors.append("Папка библиотеки шаблонов недоступна или не существует.")

        if not self.output_dir:
            errors.append("Не указан путь к папке выгрузки.")
        elif not Path(self.output_dir).exists():
            errors.append("Папка выгрузки недоступна или не существует.")

        return errors

