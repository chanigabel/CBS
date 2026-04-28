"""Unit tests for ExportService."""

import pytest
from pathlib import Path
from fastapi import HTTPException
from unittest.mock import patch

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.session import SessionRecord
from webapp.services.export_service import ExportService
from webapp.services.session_service import SessionService


def make_session_with_workbook(session_id="export-session"):
    svc = SessionService()
    svc.clear_all()
    # Use the actual VBA sheet names that ExportEngine expects
    sheet = SheetDataset(
        sheet_name="דיירים יחידים",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name"],
        rows=[
            {"first_name": "Alice", "last_name": "Smith",
             "first_name_corrected": "Alice", "last_name_corrected": "Smith"},
        ],
    )
    wb = WorkbookDataset(source_file="test.xlsx", sheets=[sheet])
    record = SessionRecord(
        session_id=session_id,
        source_file_path="uploads/export-session.xlsx",
        working_copy_path="work/export-session.xlsx",
        original_filename="test.xlsx",
        status="standardized",
        workbook_dataset=wb,
    )
    svc.create(record)
    return svc, record


@pytest.fixture(autouse=True)
def clear_registry():
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


def test_successful_export_returns_path_with_normalized_suffix(tmp_path):
    svc, _ = make_session_with_workbook()
    export_svc = ExportService(svc, tmp_path / "output")
    output_path = export_svc.export("export-session")
    assert "_standardized_" in output_path.name
    assert output_path.suffix == ".xlsx"
    assert output_path.exists()


def test_export_failure_raises_500_and_preserves_session(tmp_path):
    svc, record = make_session_with_workbook()
    original_dataset = record.workbook_dataset
    export_svc = ExportService(svc, tmp_path / "output")

    with patch(
        "webapp.services.export_service.Workbook",
        side_effect=RuntimeError("disk full"),
    ):
        with pytest.raises(HTTPException) as exc_info:
            export_svc.export("export-session")
        assert exc_info.value.status_code == 500

    # Session state must be preserved
    record_after = svc.get("export-session")
    assert record_after.workbook_dataset is original_dataset


def test_export_raises_404_for_unknown_session(tmp_path):
    svc = SessionService()
    export_svc = ExportService(svc, tmp_path / "output")
    with pytest.raises(HTTPException) as exc_info:
        export_svc.export("ghost-session")
    assert exc_info.value.status_code == 404


def test_export_raises_500_when_no_workbook_dataset(tmp_path):
    svc = SessionService()
    record = SessionRecord(
        session_id="no-wb-session",
        source_file_path="uploads/no-wb.xlsx",
        working_copy_path="work/no-wb.xlsx",
        original_filename="test.xlsx",
        status="uploaded",
        workbook_dataset=None,
    )
    svc.create(record)
    export_svc = ExportService(svc, tmp_path / "output")
    with pytest.raises(HTTPException) as exc_info:
        export_svc.export("no-wb-session")
    assert exc_info.value.status_code == 500
