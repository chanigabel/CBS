from pathlib import Path

import pytest
from fastapi import HTTPException
from openpyxl import Workbook

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from src.excel_standardization.io_layer.excel_reader import ExcelReader
from src.excel_standardization.io_layer.excel_to_json_extractor import ExcelToJsonExtractor
from webapp.models.session import SessionRecord
from webapp.services.column_mapping_schema import ColumnMappingSchemaService
from webapp.services.session_service import SessionService
from webapp.services.standardization_service import StandardizationService
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


def test_update_column_mapping_records_mapping_without_mutating_rows():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["source_first", "last_name"],
        rows=[{"source_first": "Dana", "last_name": "Cohen"}],
        metadata={},
    )
    session_service.create(_record("s1", sheet))

    response = WorkbookService(session_service).update_column_mapping(
        "s1",
        "Sheet1",
        "source_first",
        "first_name",
    )

    assert response.field_names == ["source_first", "last_name"]
    assert response.column_mappings == {"source_first": "first_name"}
    assert sheet.rows[0] == {"source_first": "Dana", "last_name": "Cohen"}


def test_editing_allows_duplicate_column_mappings():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["source_first", "source_last"],
        rows=[{"source_first": "Dana", "source_last": "Cohen"}],
        metadata={},
    )
    session_service.create(_record("duplicate-edit-session", sheet))

    service = WorkbookService(session_service)
    service.update_column_mapping("duplicate-edit-session", "Sheet1", "source_first", "first_name")
    response = service.update_column_mapping(
        "duplicate-edit-session",
        "Sheet1",
        "source_last",
        "first_name",
    )

    assert response.column_mappings == {
        "source_first": "first_name",
        "source_last": "first_name",
    }
    assert sheet.field_names == ["source_first", "source_last"]
    assert sheet.rows[0] == {"source_first": "Dana", "source_last": "Cohen"}


def test_update_column_mapping_after_standardization_marks_working_dataset_dirty():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["custom_first", "last_name"],
        rows=[{"custom_first": "Dana", "last_name": "Cohen"}],
        metadata={},
    )
    record = _record("dirty-map-session", sheet)
    record.status = "standardized"
    record.working_dataset_dirty = False
    session_service.create(record)

    WorkbookService(session_service).update_column_mapping(
        "dirty-map-session",
        "Sheet1",
        "custom_first",
        "first_name",
    )

    assert session_service.get("dirty-map-session").working_dataset_dirty is True


def test_noop_column_mapping_does_not_mark_working_dataset_dirty():
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
    record = _record("noop-map-session", sheet)
    record.status = "standardized"
    record.working_dataset_dirty = False
    session_service.create(record)

    WorkbookService(session_service).update_column_mapping(
        "noop-map-session",
        "Sheet1",
        "first_name",
        "first_name",
    )

    assert session_service.get("noop-map-session").working_dataset_dirty is False


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

    payload = WorkbookService(session_service).get_sheet_data("s2", "Sheet1")

    assert "custom_first" in payload.field_names
    assert payload.column_mappings == {"custom_first": "first_name"}
    assert payload.column_display_names == {"custom_first": "custom_first \u2192 first_name"}
    assert payload.rows[0]["custom_first"] == "Dana"


def test_standardization_blocks_duplicate_column_mappings():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="People",
        header_row=1,
        header_rows_count=1,
        field_names=["source_first", "source_last"],
        rows=[{"source_first": "Dana", "source_last": "Cohen"}],
        metadata={},
    )
    record = _record("duplicate-standardize-session", sheet)
    record.column_mappings = {
        "People": {
            "source_first": "first_name",
            "source_last": "first_name",
        }
    }
    session_service.create(record)
    workbook_service = WorkbookService(session_service)

    with pytest.raises(HTTPException) as exc_info:
        StandardizationService(session_service, workbook_service=workbook_service).standardize(
            "duplicate-standardize-session",
            sheet_name="People",
        )

    assert exc_info.value.status_code == 400
    assert "Cannot start standardization" in exc_info.value.detail
    assert sheet.field_names == ["source_first", "source_last"]
    assert sheet.rows[0] == {"source_first": "Dana", "source_last": "Cohen"}


def test_standardization_applies_valid_mapping_before_pipeline():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="People",
        header_row=1,
        header_rows_count=1,
        field_names=["source_first", "source_last", "gender"],
        rows=[{"source_first": " Dana ", "source_last": "Cohen", "gender": "F"}],
        metadata={},
    )
    record = _record("valid-standardize-session", sheet)
    record.column_mappings = {
        "People": {
            "source_first": "first_name",
            "source_last": "last_name",
        }
    }
    session_service.create(record)
    workbook_service = WorkbookService(session_service)

    response = StandardizationService(
        session_service,
        workbook_service=workbook_service,
    ).standardize("valid-standardize-session", sheet_name="People")

    standardized_sheet = session_service.get(
        "valid-standardize-session"
    ).workbook_dataset.get_sheet_by_name("People")
    assert response.status == "standardized"
    assert "first_name" in standardized_sheet.field_names
    assert "last_name" in standardized_sheet.field_names
    assert standardized_sheet.rows[0]["first_name"] == " Dana "
    assert standardized_sheet.rows[0]["first_name_corrected"] == "Dana"
    assert session_service.get("valid-standardize-session").column_mappings == {}


def test_swapping_two_column_mappings_works():
    session_service = SessionService()
    session_service.clear_all()
    sheet = SheetDataset(
        sheet_name="People",
        header_row=1,
        header_rows_count=1,
        field_names=["col_a", "col_b", "gender"],
        rows=[{"col_a": "Cohen", "col_b": "Dana", "gender": "F"}],
        metadata={},
    )
    record = _record("swap-session", sheet)
    record.column_mappings = {"People": {"col_a": "last_name", "col_b": "first_name"}}
    session_service.create(record)
    workbook_service = WorkbookService(session_service)

    StandardizationService(session_service, workbook_service=workbook_service).standardize(
        "swap-session",
        sheet_name="People",
    )

    standardized_sheet = session_service.get("swap-session").workbook_dataset.get_sheet_by_name("People")
    assert standardized_sheet.rows[0]["first_name"] == "Dana"
    assert standardized_sheet.rows[0]["last_name"] == "Cohen"


def test_apply_column_mappings_fails_atomically_without_overwriting_data():
    session_service = SessionService()
    sheet = SheetDataset(
        sheet_name="People",
        header_row=1,
        header_rows_count=1,
        field_names=["source_first", "first_name"],
        rows=[{"source_first": "Dana", "first_name": "Existing"}],
        metadata={},
    )
    service = WorkbookService(session_service)

    with pytest.raises(HTTPException):
        service.apply_column_mappings_to_sheet(sheet, {"source_first": "first_name"})

    assert sheet.field_names == ["source_first", "first_name"]
    assert sheet.rows[0] == {"source_first": "Dana", "first_name": "Existing"}


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


def test_duplicate_synonym_columns_are_preserved_in_extraction():
    hebrew_first_name = "\u05e9\u05dd \u05e4\u05e8\u05d8\u05d9"
    hebrew_first_name_key = "\u05e9\u05dd_\u05e4\u05e8\u05d8\u05d9"
    hebrew_dana = "\u05d3\u05e0\u05d4"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Sheet1"
    worksheet.append(["first_name", hebrew_first_name, "last_name"])
    worksheet.append(["Dana", hebrew_dana, "Cohen"])

    dataset = ExcelToJsonExtractor(ExcelReader()).extract_sheet_to_json(worksheet)

    assert "first_name" in dataset.field_names
    assert hebrew_first_name_key in dataset.field_names
    assert dataset.rows[0]["first_name"] == "Dana"
    assert dataset.rows[0][hebrew_first_name_key] == hebrew_dana
