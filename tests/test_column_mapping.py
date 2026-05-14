from pathlib import Path

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.session import SessionRecord
from webapp.services.column_mapping_schema import ColumnMappingSchemaService
from webapp.services.session_service import SessionService
from webapp.services.workbook_service import WorkbookService


def _record(session_id: str, sheet: SheetDataset) -> SessionRecord:
    return SessionRecord(
        session_id=session_id,
        source_file_path="source.xlsx",
        working_copy_path="work.xlsx",
        original_filename="source.xlsx",
        status="uploaded",
        workbook_dataset=WorkbookDataset(
            source_file=str(Path("source.xlsx")),
            sheets=[sheet],
            metadata={},
        ),
    )


def test_update_column_mapping_renames_field_names_and_rows():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["שם פרטי", "last_name"],
        rows=[{"שם פרטי": "Dana", "last_name": "Cohen"}],
        metadata={},
    )
    session_service.create(_record("s1", sheet))

    schema = ColumnMappingSchemaService(Path("config/column_mapping_schema.json"))
    response = WorkbookService(session_service, schema).update_column_mapping(
        "s1",
        "Sheet1",
        "שם פרטי",
        "first name",
    )

    assert response.field_names == ["first_name", "last_name"]
    assert response.column_mappings == {"שם פרטי": "first_name"}
    assert sheet.rows[0] == {"first_name": "Dana", "last_name": "Cohen"}


def test_get_sheet_data_includes_active_column_mappings():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["custom_first"],
        rows=[{"custom_first": "Dana"}],
        metadata={},
    )
    record = _record("s2", sheet)
    record.column_mappings = {"Sheet1": {"custom_first": "first_name"}}
    session_service.create(record)

    service = WorkbookService(session_service)
    service.apply_column_mappings_to_sheet(sheet, record.column_mappings["Sheet1"])
    payload = service.get_sheet_data("s2", "Sheet1")

    assert "first_name" in payload.field_names
    assert payload.column_mappings == {"custom_first": "first_name"}
    assert payload.rows[0]["first_name"] == "Dana"


def test_column_mapping_schema_adds_and_removes_synonym(tmp_path):
    schema = ColumnMappingSchemaService(tmp_path / "column_mapping_schema.json")

    schema.add_mapping("first_name", "given name")
    assert schema.resolve("given name") == "first_name"

    schema.remove_mapping("first_name", "given name")
    assert "given name" not in schema.mappings()["first_name"]


def test_column_schema_response_uses_canonical_fields_only(tmp_path):
    schema = ColumnMappingSchemaService(tmp_path / "column_mapping_schema.json")
    response = WorkbookService(SessionService(), schema).get_column_schema()

    assert "first_name" in response.fields
    assert "first name" not in response.fields
    assert "first name" in response.suggestions


def test_reload_column_mapping_returns_current_schema(tmp_path):
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name"],
        rows=[{"first_name": "Dana"}],
        metadata={},
    )
    session_service.create(_record("s3", sheet))
    schema = ColumnMappingSchemaService(tmp_path / "column_mapping_schema.json")

    response = WorkbookService(session_service, schema).reload_column_mapping("s3", "Sheet1")

    assert "first_name" in response.fields
