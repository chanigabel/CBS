from datetime import date

from src.excel_standardization.data_types import SheetDataset
from src.excel_standardization.engine_management import BaseEngine, EngineManager
from src.excel_standardization.engines.date_engine import DateEngine
from src.excel_standardization.engines.gender_engine import GenderEngine
from src.excel_standardization.engines.identifier_engine import IdentifierEngine
from src.excel_standardization.engines.name_engine import NameEngine
from src.excel_standardization.engines.text_processor import TextProcessor
from src.excel_standardization.processing.standardization_pipeline import StandardizationPipeline


class CustomAppendEngine(BaseEngine):
    engine_key = "custom_append"
    display_name = "Custom Append Engine"
    version = "1.0.0"
    description = "Adds a marker field for dynamic engine tests."
    supported_fields = ["custom_marker"]

    def run(self, payload, context):
        payload["custom_marker"] = context["engine_config"].settings.get("marker", "ok")
        return payload


def _manager(tmp_path):
    return EngineManager(tmp_path / "engine_config.json")


def _pipeline(manager):
    reference_date = date(2026, 5, 14)
    return StandardizationPipeline(
        name_engine=NameEngine(TextProcessor()),
        gender_engine=GenderEngine(),
        date_engine=DateEngine(reference_date=reference_date),
        identifier_engine=IdentifierEngine(),
        reference_date=reference_date,
        engine_manager=manager,
    )


def _dataset(row):
    return SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=list(row.keys()),
        rows=[row],
        metadata={},
    )


def test_dynamic_config_disables_builtin_engine(tmp_path):
    manager = _manager(tmp_path)
    manager.disable("gender", role="engine_admin", user="test")

    normalized = _pipeline(manager).normalize_dataset(_dataset({"gender": "male"}))

    assert "gender_corrected" not in normalized.rows[0]
    assert normalized.metadata["standardization_engines"]["gender"] is False


def test_dynamic_runner_executes_registered_custom_engine(tmp_path):
    manager = _manager(tmp_path)
    manager.add_engine(
        {
            "engine_key": "custom_append",
            "display_name": "Custom Append Engine",
            "class": "tests.test_dynamic_engine_management.CustomAppendEngine",
            "enabled": True,
            "priority": 5,
            "run_mode": "sequential",
            "version": "1.0.0",
            "on_error": "stop",
            "settings": {"marker": "dynamic"},
        },
        role="system_admin",
        user="test",
    )

    normalized = _pipeline(manager).normalize_row({"source": "value"})

    assert normalized["custom_marker"] == "dynamic"


def test_add_engine_without_class_uses_passthrough_engine(tmp_path):
    manager = _manager(tmp_path)

    summary = manager.add_engine(
        {
            "engine_key": "location",
            "enabled": True,
            "priority": 50,
            "run_mode": "sequential",
        },
        role="system_admin",
        user="test",
    )

    assert summary["engine_key"] == "location"
    assert summary["enabled"] is True
    assert manager.registry.has("location") is True
