import logging
import sys
from logging.handlers import RotatingFileHandler

from webapp.logging_config import configure_logging


def test_logging_configuration_initializes_without_crashing(tmp_path):
    log_file = configure_logging(log_dir=tmp_path / "logs", force=True)

    logging.getLogger("tests.webapp.logging").info("logging smoke")

    assert log_file.exists()
    assert log_file.parent.name == "logs"
    assert "logging smoke" in log_file.read_text(encoding="utf-8")


def test_logging_configuration_uses_rotating_file_handler(tmp_path):
    configure_logging(log_dir=tmp_path / "logs", force=True)

    root_logger = logging.getLogger()
    file_handlers = [h for h in root_logger.handlers if isinstance(h, RotatingFileHandler)]

    assert len(file_handlers) == 1
    rotating_handler = file_handlers[0]
    assert rotating_handler.baseFilename == str(tmp_path / "logs" / "app.log")
    assert rotating_handler.maxBytes == 5 * 1024 * 1024
    assert rotating_handler.backupCount == 3


def test_logging_configuration_uses_localappdata_in_frozen_mode(tmp_path, monkeypatch):
    local_app_data = tmp_path / "LocalAppData"
    monkeypatch.setenv("LOCALAPPDATA", str(local_app_data))
    monkeypatch.setattr(sys, "frozen", True, raising=False)

    log_file = configure_logging(force=True)

    assert log_file == local_app_data / "Excelstandardization" / "logs" / "app.log"
    assert log_file.exists()
    assert log_file.parent.name == "logs"
