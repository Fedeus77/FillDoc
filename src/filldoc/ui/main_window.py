from __future__ import annotations

from pathlib import Path

from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

from filldoc.ui.tabs.projects_tab import ProjectsTab
from filldoc.ui.tabs.templates_tab import TemplatesTab
from filldoc.ui.tabs.variables_tab import VariablesTab
from filldoc.ui.tabs.settings_tab import SettingsTab
from filldoc.core.settings import AppSettings
from filldoc.ui.theme import ThemeManager, build_global_stylesheet


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("FillDoc")
        self.resize(1100, 700)

        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        self.settings_tab = SettingsTab(self)
        self.projects_tab = ProjectsTab(self)
        self.templates_tab = TemplatesTab(self)
        self.variables_tab = VariablesTab(self)

        self.tabs.addTab(self.projects_tab, "Проекты")
        self.tabs.addTab(self.templates_tab, "Шаблоны")
        self.tabs.addTab(self.variables_tab, "Переменные")
        self.tabs.addTab(self.settings_tab, "Настройки")

        self.statusBar().showMessage("Готово")

        # Применяем начальную тему
        self._apply_theme_to_all()

        # ── Горячие клавиши ───────────────────────────────────────────────────
        save_sc = QShortcut(QKeySequence("Ctrl+S"), self)
        save_sc.activated.connect(self._hotkey_save)
        refresh_sc = QShortcut(QKeySequence("Ctrl+R"), self)
        refresh_sc.activated.connect(self._hotkey_refresh)

        self.settings_tab.settings_changed.connect(self._on_settings_changed)
        self.settings_tab.theme_changed.connect(self._on_theme_changed)
        self._on_settings_changed()

    def show_status(self, message: str, timeout_ms: int = 4000) -> None:
        """Показывает сообщение в статус-баре вместо модального QMessageBox."""
        self.statusBar().showMessage(message, timeout_ms)

    def _on_theme_changed(self, theme_name: str) -> None:
        """Применяет новую тему ко всему приложению."""
        tm = ThemeManager.instance()
        tm.set_theme(theme_name)
        app = QApplication.instance()
        if app:
            app.setStyleSheet(build_global_stylesheet(tm.colors))
        self._apply_theme_to_all()

    def _apply_theme_to_all(self) -> None:
        """Обновляет тему во всех вкладках."""
        c = ThemeManager.instance().colors
        for tab in (self.projects_tab, self.templates_tab, self.variables_tab, self.settings_tab):
            if hasattr(tab, "apply_theme"):
                tab.apply_theme(c)

    def _hotkey_save(self) -> None:
        """Ctrl+S: сохранить в зависимости от активной вкладки."""
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.projects_tab._save_all()
        elif idx == 1:
            self.templates_tab._save_to_excel()

    def _hotkey_refresh(self) -> None:
        """Ctrl+R: обновить в зависимости от активной вкладки."""
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.projects_tab._load_projects()
        elif idx == 1:
            self.templates_tab._reload_all()

    def _on_settings_changed(self) -> None:
        s: AppSettings = self.settings_tab.get_settings()

        # Сохраняем project_docs_dirs из живого состояния projects_tab
        existing = getattr(self.projects_tab, "_settings", None)
        if existing and existing.project_docs_dirs:
            merged = dict(existing.project_docs_dirs)
            merged.update(s.project_docs_dirs)
            s.project_docs_dirs = merged
            self.settings_tab._settings.project_docs_dirs = dict(s.project_docs_dirs)

        self.projects_tab.set_settings(s)
        self.templates_tab.set_settings(s)
        self.variables_tab.set_settings(s)

        excel_ok = bool(s.excel_path) and Path(s.excel_path).is_file()
        templates_ok = bool(s.templates_dir) and Path(s.templates_dir).is_dir()

        if excel_ok:
            self.projects_tab._load_projects()
            self.variables_tab._reload()
        if excel_ok or templates_ok:
            self.templates_tab._load_projects()
        if templates_ok:
            self.templates_tab._scan_templates()
