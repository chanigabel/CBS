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
