import logging

from webapp.logging_config import configure_logging


def test_logging_configuration_initializes_without_crashing(tmp_path):
    log_file = configure_logging(log_dir=tmp_path / "logs", force=True)

    logging.getLogger("tests.webapp.logging").info("logging smoke")

    assert log_file.exists()
    assert log_file.parent.name == "logs"
    assert "logging smoke" in log_file.read_text(encoding="utf-8")
