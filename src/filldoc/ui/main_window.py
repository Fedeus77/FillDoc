from __future__ import annotations

from PySide6.QtWidgets import QMainWindow, QTabWidget

from filldoc.ui.tabs.projects_tab import ProjectsTab
from filldoc.ui.tabs.templates_tab import TemplatesTab
from filldoc.ui.tabs.fill_tab import FillTab
from filldoc.ui.tabs.settings_tab import SettingsTab
from filldoc.core.settings import AppSettings


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("FillDoc (MVP)")
        self.resize(1100, 700)

        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        self.settings_tab = SettingsTab(self)
        self.projects_tab = ProjectsTab(self)
        self.templates_tab = TemplatesTab(self)
        self.fill_tab = FillTab(self)

        self.tabs.addTab(self.projects_tab, "Проекты")
        self.tabs.addTab(self.templates_tab, "Шаблоны")
        self.tabs.addTab(self.fill_tab, "Заполнение")
        self.tabs.addTab(self.settings_tab, "Настройки")

        # MVP: вкладки общаются через окно (простая связка без сложной архитектуры)
        self.settings_tab.settings_changed.connect(self._on_settings_changed)
        self._on_settings_changed()

    def _on_settings_changed(self) -> None:
        s: AppSettings = self.settings_tab.get_settings()
        self.projects_tab.set_settings(s)
        self.templates_tab.set_settings(s)
        self.fill_tab.set_settings(s)

