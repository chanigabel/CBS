import copy

import pytest

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.session import SessionRecord
from webapp.services.report_service import ReportService
from webapp.services.session_service import SessionService


@pytest.fixture
def session_service():
    svc = SessionService()
    svc.clear_all()
    yield svc
    svc.clear_all()


def _standardized_record(dirty=False):
    sheet = SheetDataset(
        sheet_name="Sheet1",
        header_row=1,
        header_rows_count=1,
        field_names=[
            "first_name",
            "first_name_corrected",
            "gender_status",
            "identifier_status",
            "birth_date_status",
            "_validation_status",
            "_standardization_failures",
        ],
        rows=[
            {
                "_row_uid": "row-1",
                "first_name": "Alice",
                "first_name_corrected": "Alice",
                "gender_status": "תקין",
                "identifier_status": "",
                "birth_date_status": "שנה חסרה והושלמה",
                "_validation_status": "ok",
                "_standardization_failures": [],
            },
            {
                "_row_uid": "row-2",
                "first_name": "Bob",
                "first_name_corrected": "Robert",
                "gender_status": "לא תקין",
                "identifier_status": "ת.ז. לא תקינה",
                "birth_date_status": "",
                "_validation_status": "failed",
                "_standardization_failures": ["identifier failed"],
            },
        ],
    )
    return SessionRecord(
        session_id="report-session",
        source_file_path="uploads/report-session.xlsx",
        working_copy_path="work/report-session.xlsx",
        original_filename="source.xlsx",
        status="standardized",
        workbook_dataset=WorkbookDataset(source_file="work/report-session.xlsx", sheets=[sheet]),
        edits={("Sheet1", "row-2", "first_name"): "Robert"},
        working_dataset_dirty=dirty,
    )


def test_report_works_for_standardized_dataset(session_service):
    session_service.create(_standardized_record())

    report = ReportService(session_service).build("report-session")

    assert report.session_id == "report-session"
    assert report.file_name == "source.xlsx"
    assert report.status == "standardized"
    assert report.export_ready is True
    assert report.dirty is False
    assert report.summary.total_sheets == 1
    assert report.summary.total_rows == 2
    assert report.summary.edited_cells == 1
    assert report.summary.corrected_fields == 1
    assert report.manual_edits.edited_sheets == ["Sheet1"]
    assert report.manual_edits.edited_fields == ["first_name"]


def test_report_counts_status_values(session_service):
    session_service.create(_standardized_record())

    sheet_report = ReportService(session_service).build("report-session").sheets[0]

    assert sheet_report.status_counts["birth_date_status"]["שנה חסרה והושלמה"] == 1
    assert sheet_report.status_counts["gender_status"]["לא תקין"] == 1
    assert sheet_report.status_counts["identifier_status"]["ת.ז. לא תקינה"] == 1
    assert sheet_report.status_counts["_validation_status"]["failed"] == 1
    assert sheet_report.status_counts["_standardization_failures"]["identifier failed"] == 1
    assert sheet_report.rows_with_warnings == 1
    assert sheet_report.rows_with_errors == 1


def test_report_includes_dirty_export_state(session_service):
    session_service.create(_standardized_record(dirty=True))

    report = ReportService(session_service).build("report-session")

    # Even if manual edits were made after standardization, the report should
    # indicate the dataset is dirty/stale but allow export because a
    # standardized result exists.
    assert report.export_ready is True
    assert report.dirty is True
    assert report.stale is True
    assert "Run Standardization again" not in (report.export_blocked_reason or "")


def test_report_does_not_mutate_workbook_dataset(session_service):
    record = _standardized_record()
    session_service.create(record)
    before = copy.deepcopy(record.workbook_dataset)

    ReportService(session_service).build("report-session")

    assert record.workbook_dataset == before


def test_report_does_not_reextract_when_dataset_exists(session_service, monkeypatch):
    record = _standardized_record()
    session_service.create(record)

    def fail_if_called(*args, **kwargs):
        raise AssertionError("report must not read or re-extract workbook files")

    monkeypatch.setattr("webapp.services.workbook_loader.extract_workbook_dataset", fail_if_called)

    report = ReportService(session_service).build("report-session")

    assert report.summary.total_rows == 2


def test_report_returns_useful_empty_report_before_standardization(session_service):
    record = SessionRecord(
        session_id="uploaded-session",
        source_file_path="uploads/uploaded-session.xlsx",
        working_copy_path="work/uploaded-session.xlsx",
        original_filename="uploaded.xlsx",
        status="uploaded",
    )
    session_service.create(record)

    report = ReportService(session_service).build("uploaded-session")

    assert report.session_id == "uploaded-session"
    assert report.file_name == "uploaded.xlsx"
    assert report.export_ready is False
    assert report.summary.total_rows == 0
    assert report.export_blocked_reason == "Workbook data is not loaded yet."
