from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QEvent, QObject
from PySide6.QtGui import QKeySequence, QShortcut
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget

from filldoc.ui.tabs.projects_tab import ProjectsTab
from filldoc.ui.tabs.templates_tab import TemplatesTab
from filldoc.ui.tabs.variables_tab import VariablesTab
from filldoc.ui.tabs.settings_tab import SettingsTab
from filldoc.core.settings import AppSettings
from filldoc.ui.icons import SVG_SETTINGS, icon_btn, update_icon_btn
from filldoc.ui.theme import ThemeManager, build_global_stylesheet


class _SettingsButtonPositioner(QObject):
    def __init__(self, window: "MainWindow") -> None:
        super().__init__(window)
        self._window = window

    def eventFilter(self, watched: QObject, event: QEvent) -> bool:
        if event.type() in (QEvent.Type.Resize, QEvent.Type.Show):
            self._window._position_settings_button()
        return super().eventFilter(watched, event)


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
        self.settings_btn = icon_btn(SVG_SETTINGS, "Settings", icon_size=16, button_size=26)
        self.settings_btn.clicked.connect(self._show_settings)

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

        self.settings_tab_index = self.tabs.indexOf(self.settings_tab)
        self.settings_btn.setToolTip(self.tabs.tabText(self.settings_tab_index))
        self.tabs.tabBar().setTabVisible(self.settings_tab_index, False)
        self.settings_btn.setParent(self.tabs)
        self.settings_btn.show()
        self._settings_button_positioner = _SettingsButtonPositioner(self)
        self.tabs.tabBar().installEventFilter(self._settings_button_positioner)
        self.tabs.installEventFilter(self._settings_button_positioner)
        self._position_settings_button()

        self.settings_tab.settings_changed.connect(self._on_settings_changed)
        self.settings_tab.theme_changed.connect(self._on_theme_changed)
        self.variables_tab.dictionary_changed.connect(self._on_dictionary_changed)
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
        update_icon_btn(
            self.settings_btn,
            SVG_SETTINGS,
            icon_color=c.icon_color,
            bg=c.bg_tab,
            hover=c.bg_hover,
            pressed=c.bg_tab_selected,
            icon_size=16,
            button_size=26,
        )
        self._position_settings_button()

    def _show_settings(self) -> None:
        self.tabs.setCurrentIndex(self.settings_tab_index)

    def _position_settings_button(self) -> None:
        tab_bar = self.tabs.tabBar()
        margin = 8
        x = max(margin, self.tabs.width() - self.settings_btn.width() - margin)
        y = tab_bar.y() + max(0, (tab_bar.height() - self.settings_btn.height()) // 2)
        self.settings_btn.move(x, y)
        self.settings_btn.raise_()

    def _hotkey_save(self) -> None:
        """Ctrl+S: сохранить в зависимости от активной вкладки."""
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.projects_tab._save_all()
        elif idx == 1:
            self.templates_tab._save_to_excel()
        elif idx == 2:
            self.variables_tab._save_dictionary()

    def _hotkey_refresh(self) -> None:
        """Ctrl+R: обновить в зависимости от активной вкладки."""
        idx = self.tabs.currentIndex()
        if idx == 0:
            self.projects_tab._load_projects()
        elif idx == 1:
            self.templates_tab._reload_all()
        elif idx == 2:
            self.variables_tab.reload_all_variables()

    def _on_dictionary_changed(self) -> None:
        self.templates_tab.reload_dictionary()

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
            self.variables_tab._reload_dictionary(show_errors=False)
