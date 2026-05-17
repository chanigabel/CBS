import logging

import pytest
from fastapi import HTTPException

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from tests.webapp.conftest import make_xlsx_bytes
from webapp.models.requests import CellEditRequest
from webapp.models.session import SessionRecord
from webapp.services.edit_service import EditService
from webapp.services.processing_report_service import ProcessingReportService
from webapp.services.session_service import SessionService
from webapp.services.upload_service import UploadService


def test_upload_service_emits_success_logs(tmp_path, caplog):
    svc = SessionService()
    svc.clear_all()
    upload_service = UploadService(
        svc,
        tmp_path / "uploads",
        tmp_path / "work",
        ProcessingReportService(svc),
    )

    with caplog.at_level(logging.INFO):
        upload_service.handle_upload("test.xlsx", make_xlsx_bytes(["Sheet1"]))

    events = {record.__dict__.get("event") for record in caplog.records}
    assert "upload_started" in events
    assert "upload_saved_internal_files" in events
    assert "upload_session_created" in events
    assert "upload_successful" in events


def test_upload_service_emits_invalid_extension_warning(tmp_path, caplog):
    svc = SessionService()
    svc.clear_all()
    upload_service = UploadService(svc, tmp_path / "uploads", tmp_path / "work")

    with caplog.at_level(logging.WARNING), pytest.raises(HTTPException):
        upload_service.handle_upload("test.csv", b"not excel")

    events = {record.__dict__.get("event") for record in caplog.records}
    assert "upload_rejected_invalid_extension" in events


def test_edit_service_emits_success_and_rejection_logs(caplog):
    svc = SessionService()
    svc.clear_all()
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "first_name_corrected"],
        rows=[{"_row_uid": "row-1", "first_name": "Alice", "first_name_corrected": "Alice"}],
    )
    svc.create(
        SessionRecord(
            session_id="edit-log-session",
            source_file_path="uploads/edit-log-session.xlsx",
            working_copy_path="work/edit-log-session.xlsx",
            original_filename="edit.xlsx",
            status="standardized",
            workbook_dataset=WorkbookDataset(source_file="work/edit-log-session.xlsx", sheets=[sheet]),
        )
    )

    edit_service = EditService(svc)
    with caplog.at_level(logging.INFO):
        edit_service.edit_cell(
            "edit-log-session",
            "Sheet1",
            CellEditRequest(row_uid="row-1", field_name="first_name", new_value="Alicia"),
        )
    # Editing corrected fields is allowed and should succeed.
    with caplog.at_level(logging.INFO):
        edit_service.edit_cell(
            "edit-log-session",
            "Sheet1",
            CellEditRequest(row_uid="row-1", field_name="first_name_corrected", new_value="X"),
        )

    events = {record.__dict__.get("event") for record in caplog.records}
    assert "cell_edit_requested" in events
    assert "cell_edit_succeeded" in events
    assert "cell_edit_rejected_system_field" not in events
