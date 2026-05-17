"""Unit tests for EditService."""

import pytest
from fastapi import HTTPException

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.requests import CellEditRequest
from webapp.models.session import SessionRecord
from webapp.services.edit_service import EditService
from webapp.services.session_service import SessionService


def make_session_with_sheet(session_id="edit-session"):
    svc = SessionService()
    svc.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name"],
        rows=[
            {"first_name": "Alice", "last_name": "Smith", "_row_uid": "uid-alice-001"},
            {"first_name": "Bob", "last_name": "Jones", "_row_uid": "uid-bob-002"},
        ],
    )
    wb = WorkbookDataset(source_file="test.xlsx", sheets=[sheet])
    record = SessionRecord(
        session_id=session_id,
        source_file_path="uploads/edit-session.xlsx",
        working_copy_path="work/edit-session.xlsx",
        original_filename="test.xlsx",
        status="uploaded",
        workbook_dataset=wb,
    )
    svc.create(record)
    return svc, EditService(svc)


@pytest.fixture(autouse=True)
def clear_registry():
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


def test_valid_edit_mutates_in_memory_row():
    svc, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="uid-alice-001", field_name="first_name", new_value="Carol")
    response = edit_svc.edit_cell("edit-session", "Sheet1", req)

    assert response.row_uid == "uid-alice-001"
    assert response.updated_row["first_name"] == "Carol"

    # Verify in-memory mutation
    record = svc.get("edit-session")
    assert record.workbook_dataset.get_sheet_by_name("Sheet1").rows[0]["first_name"] == "Carol"


def test_valid_edit_returns_updated_row():
    _, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="uid-bob-002", field_name="last_name", new_value="Williams")
    response = edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert response.updated_row["last_name"] == "Williams"
    assert response.updated_row["first_name"] == "Bob"


def test_unknown_row_uid_raises_404():
    _, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="nonexistent-uid-99999", field_name="first_name", new_value="X")
    with pytest.raises(HTTPException) as exc_info:
        edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert exc_info.value.status_code == 404


def test_unknown_field_name_raises_400():
    _, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="uid-alice-001", field_name="nonexistent_field", new_value="X")
    with pytest.raises(HTTPException) as exc_info:
        edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert exc_info.value.status_code == 400


def test_unknown_sheet_raises_404():
    _, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="uid-alice-001", field_name="first_name", new_value="X")
    with pytest.raises(HTTPException) as exc_info:
        edit_svc.edit_cell("edit-session", "NonExistentSheet", req)
    assert exc_info.value.status_code == 404


def test_edit_is_recorded_in_session_edits():
    svc, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="uid-alice-001", field_name="first_name", new_value="Dave")
    edit_svc.edit_cell("edit-session", "Sheet1", req)

    record = svc.get("edit-session")
    assert ("Sheet1", "uid-alice-001", "first_name") in record.edits
    assert record.edits[("Sheet1", "uid-alice-001", "first_name")] == "Dave"


@pytest.mark.parametrize(
    "field_name",
    [
        "_standardization_failures",
        "_row_uid",
        "row_uid",
    ],
)
def test_blocks_computed_and_system_fields(field_name):
    _, edit_svc = make_session_with_sheet()
    record = edit_svc.session_service.get("edit-session")
    sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
    sheet.field_names.extend(["first_name_corrected", "gender_status"])
    sheet.rows[0].update(
        {
            "first_name_corrected": "Alice",
            "gender_status": "",
            "_validation_status": "",
            "_standardization_failures": [],
            "row_uid": "public-row-id",
        }
    )

    req = CellEditRequest(row_uid="uid-alice-001", field_name=field_name, new_value="Blocked")
    with pytest.raises(HTTPException) as exc_info:
        edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert exc_info.value.status_code == 400


def test_edit_corrected_field_succeeds():
    svc, edit_svc = make_session_with_sheet()
    record = edit_svc.session_service.get("edit-session")
    sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
    sheet.field_names.append("first_name_corrected")
    sheet.rows[0]["first_name_corrected"] = "Alice"

    req = CellEditRequest(row_uid="uid-alice-001", field_name="first_name_corrected", new_value="Alicia")
    resp = edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert resp.updated_row["first_name_corrected"] == "Alicia"


def test_edit_status_field_succeeds():
    svc, edit_svc = make_session_with_sheet()
    record = edit_svc.session_service.get("edit-session")
    sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
    sheet.field_names.append("birth_date_status")
    sheet.rows[0]["birth_date_status"] = ""

    req = CellEditRequest(row_uid="uid-alice-001", field_name="birth_date_status", new_value="תקין")
    resp = edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert resp.updated_row["birth_date_status"] == "תקין"


def test_edit_prefixed_validation_status_succeeds():
    svc, edit_svc = make_session_with_sheet()
    record = edit_svc.session_service.get("edit-session")
    sheet = record.workbook_dataset.get_sheet_by_name("Sheet1")
    sheet.field_names.append("_validation_status")
    sheet.rows[0]["_validation_status"] = "ok"

    req = CellEditRequest(row_uid="uid-alice-001", field_name="_validation_status", new_value="failed")
    resp = edit_svc.edit_cell("edit-session", "Sheet1", req)
    assert resp.updated_row["_validation_status"] == "failed"


def test_edit_marks_working_dataset_dirty():
    svc, edit_svc = make_session_with_sheet()
    req = CellEditRequest(row_uid="uid-alice-001", field_name="first_name", new_value="Dirty")
    edit_svc.edit_cell("edit-session", "Sheet1", req)

    assert svc.get("edit-session").working_dataset_dirty is True
