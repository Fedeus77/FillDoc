"""Инициализация логирования FillDoc в файл ~/.filldoc/filldoc.log."""
from __future__ import annotations

import logging
import logging.handlers
from pathlib import Path


def setup_logging(level: int = logging.DEBUG) -> None:
    """
    Настраивает корневой логгер filldoc:
    - Файл ~/.filldoc/filldoc.log, ротация 5 МБ × 3 резервных копии
    - Консоль (WARNING и выше)
    """
    log_dir = Path.home() / ".filldoc"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "filldoc.log"

    logger = logging.getLogger("filldoc")
    if logger.handlers:
        return  # уже настроен (например, при перезапуске в тестах)

    logger.setLevel(level)

    fmt = logging.Formatter(
        "%(asctime)s  %(levelname)-8s  %(name)s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # Файловый обработчик с ротацией по размеру
    fh = logging.handlers.RotatingFileHandler(
        log_path,
        maxBytes=5 * 1024 * 1024,  # 5 МБ
        backupCount=3,
        encoding="utf-8",
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # Консольный обработчик — только WARNING+
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(fmt)
    logger.addHandler(ch)
