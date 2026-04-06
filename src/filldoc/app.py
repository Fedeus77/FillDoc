from __future__ import annotations

import logging
import sys

from PySide6.QtWidgets import QApplication

from filldoc.core.logging_setup import setup_logging
from filldoc.ui.main_window import MainWindow

log = logging.getLogger("filldoc.app")


def run() -> None:
    setup_logging()
    log.info("FillDoc starting")
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    exit_code = app.exec()
    log.info("FillDoc exiting (code %d)", exit_code)
    sys.exit(exit_code)

