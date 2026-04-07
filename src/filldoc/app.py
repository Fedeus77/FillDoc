from __future__ import annotations

import logging
import sys

from PySide6.QtWidgets import QApplication

from filldoc.core.logging_setup import setup_logging
from filldoc.core.settings import AppSettings
from filldoc.ui.theme import ThemeManager, build_global_stylesheet
from filldoc.ui.main_window import MainWindow

log = logging.getLogger("filldoc.app")


def run() -> None:
    setup_logging()
    log.info("FillDoc starting")
    app = QApplication(sys.argv)
    app.setApplicationName("FillDoc")

    # Загружаем настройки и применяем сохранённую тему
    settings = AppSettings.load()
    tm = ThemeManager.instance()
    tm.set_theme(settings.theme)
    app.setStyleSheet(build_global_stylesheet(tm.colors))

    w = MainWindow()
    w.show()
    exit_code = app.exec()
    log.info("FillDoc exiting (code %d)", exit_code)
    sys.exit(exit_code)
