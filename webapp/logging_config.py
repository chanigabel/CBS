"""Central logging configuration for the web application."""

from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

_LOG_FORMAT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
_DEFAULT_MAX_BYTES = 5 * 1024 * 1024
_DEFAULT_BACKUP_COUNT = 3


def get_app_data_dir() -> Path:
    """Return the internal writable application data directory."""
    if getattr(sys, "frozen", False):
        return Path(os.environ.get("LOCALAPPDATA", Path.home())) / "Excelstandardization"
    return Path.cwd()


def get_log_dir() -> Path:
    """Return the internal application-managed log directory."""
    return get_app_data_dir() / "logs"


def configure_logging(
    *,
    log_dir: Path | None = None,
    level: int = logging.INFO,
    force: bool = False,
) -> Path:
    """Configure console and rotating-file logging for the web app.

    Returns the path of the active log file.
    """
    target_dir = Path(log_dir) if log_dir is not None else get_log_dir()
    target_dir.mkdir(parents=True, exist_ok=True)
    log_file = target_dir / "app.log"

    root = logging.getLogger()
    if getattr(root, "_excel_standardization_logging_configured", False) and not force:
        return log_file

    if force:
        for handler in list(root.handlers):
            root.removeHandler(handler)
            handler.close()

    formatter = logging.Formatter(_LOG_FORMAT)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)

    file_handler = RotatingFileHandler(
        log_file,
        maxBytes=_DEFAULT_MAX_BYTES,
        backupCount=_DEFAULT_BACKUP_COUNT,
        encoding="utf-8",
    )
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)

    root.setLevel(level)
    root.addHandler(console_handler)
    root.addHandler(file_handler)
    root._excel_standardization_logging_configured = True # type: ignore

    logging.getLogger(__name__).info(
        "logging_configured",
        extra={"event": "logging_configured", "log_file": str(log_file)},
    )
    return log_file
