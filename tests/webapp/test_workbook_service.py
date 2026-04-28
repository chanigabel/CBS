"""Unit tests for WorkbookService."""

import pytest
from fastapi import HTTPException

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.session import SessionRecord
from webapp.services.session_service import SessionService
from webapp.services.workbook_service import WorkbookService


def make_workbook_dataset():
    sheet1 = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name"],
        rows=[
            {"first_name": "Alice", "last_name": "Smith"},
            {"first_name": "Bob", "last_name": "Jones"},
        ],
    )
    sheet2 = SheetDataset(
        sheet_name="Sheet2",
        header_row=1,
        header_rows_count=1,
        field_names=["gender"],
        rows=[{"gender": "F"}],
    )
    return WorkbookDataset(source_file="test.xlsx", sheets=[sheet1, sheet2])


@pytest.fixture(autouse=True)
def clear_registry():
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


@pytest.fixture
def session_with_workbook():
    svc = SessionService()
    wb = make_workbook_dataset()
    record = SessionRecord(
        session_id="wb-session",
        source_file_path="uploads/wb-session.xlsx",
        working_copy_path="work/wb-session.xlsx",
        original_filename="test.xlsx",
        status="uploaded",
        workbook_dataset=wb,
    )
    svc.create(record)
    return svc, WorkbookService(svc)


def test_get_summary_returns_correct_sheet_names(session_with_workbook):
    _, wb_svc = session_with_workbook
    summary = wb_svc.get_summary("wb-session")
    assert summary.session_id == "wb-session"
    assert len(summary.sheets) == 2
    names = [s.sheet_name for s in summary.sheets]
    assert "Sheet1" in names
    assert "Sheet2" in names


def test_get_summary_returns_correct_row_counts(session_with_workbook):
    _, wb_svc = session_with_workbook
    summary = wb_svc.get_summary("wb-session")
    sheet1_summary = next(s for s in summary.sheets if s.sheet_name == "Sheet1")
    assert sheet1_summary.row_count == 2
    sheet2_summary = next(s for s in summary.sheets if s.sheet_name == "Sheet2")
    assert sheet2_summary.row_count == 1


def test_get_summary_returns_correct_field_names(session_with_workbook):
    _, wb_svc = session_with_workbook
    summary = wb_svc.get_summary("wb-session")
    sheet1_summary = next(s for s in summary.sheets if s.sheet_name == "Sheet1")
    assert "first_name" in sheet1_summary.field_names
    assert "last_name" in sheet1_summary.field_names


def test_get_sheet_data_returns_rows_for_valid_sheet(session_with_workbook):
    _, wb_svc = session_with_workbook
    response = wb_svc.get_sheet_data("wb-session", "Sheet1")
    assert response.sheet_name == "Sheet1"
    assert len(response.rows) == 2
    assert response.rows[0]["first_name"] == "Alice"


def test_get_sheet_data_raises_404_for_unknown_sheet(session_with_workbook):
    _, wb_svc = session_with_workbook
    with pytest.raises(HTTPException) as exc_info:
        wb_svc.get_sheet_data("wb-session", "NonExistentSheet")
    assert exc_info.value.status_code == 404


def test_get_summary_raises_404_for_unknown_session(session_with_workbook):
    _, wb_svc = session_with_workbook
    with pytest.raises(HTTPException) as exc_info:
        wb_svc.get_summary("ghost-session")
    assert exc_info.value.status_code == 404
