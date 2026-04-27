"""Unit tests for NormalizationService."""

import io
import pytest
from pathlib import Path
from openpyxl import Workbook
from fastapi import HTTPException

from src.excel_normalization.data_types import SheetDataset, WorkbookDataset
from webapp.models.session import SessionRecord
from webapp.services.session_service import SessionService
from webapp.services.normalization_service import NormalizationService


def make_xlsx_bytes(sheet_names=None) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    for name in (sheet_names or ["Sheet1"]):
        ws = wb.create_sheet(name)
        ws.append(["first_name", "last_name", "gender"])
        ws.append(["Alice", "Smith", "F"])
        ws.append(["Bob", "Jones", "M"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture(autouse=True)
def clear_registry():
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


@pytest.fixture
def session_with_file(tmp_path):
    """Create a session with a real xlsx file on disk and a pre-loaded dataset."""
    from src.excel_normalization.io_layer.excel_to_json_extractor import ExcelToJsonExtractor
    from src.excel_normalization.io_layer.excel_reader import ExcelReader

    svc = SessionService()
    file_bytes = make_xlsx_bytes(["Sheet1"])
    working_path = tmp_path / "work" / "test-session.xlsx"
    working_path.parent.mkdir(parents=True, exist_ok=True)
    working_path.write_bytes(file_bytes)

    source_path = tmp_path / "uploads" / "test-session.xlsx"
    source_path.parent.mkdir(parents=True, exist_ok=True)
    source_path.write_bytes(file_bytes)

    # Pre-load the workbook dataset so normalization has data to work with
    extractor = ExcelToJsonExtractor(ExcelReader(), skip_empty_rows=False,
                                     handle_formulas=True, preserve_types=True)
    wbd = extractor.extract_workbook_to_json(str(working_path))

    record = SessionRecord(
        session_id="test-session",
        source_file_path=str(source_path),
        working_copy_path=str(working_path),
        original_filename="test.xlsx",
        status="uploaded",
        workbook_dataset=wbd,
    )
    svc.create(record)
    return svc, NormalizationService(svc)


def test_successful_normalization_sets_status_to_normalized(session_with_file):
    svc, norm_svc = session_with_file
    response = norm_svc.normalize("test-session")
    assert response.status == "normalized"
    assert response.session_id == "test-session"

    record = svc.get("test-session")
    assert record.status == "normalized"


def test_successful_normalization_returns_stats(session_with_file):
    _, norm_svc = session_with_file
    response = norm_svc.normalize("test-session")
    assert response.sheets_processed >= 1
    assert response.total_rows >= 0
    assert len(response.per_sheet_stats) >= 1


def test_normalization_updates_workbook_dataset(session_with_file):
    svc, norm_svc = session_with_file
    norm_svc.normalize("test-session")
    record = svc.get("test-session")
    assert record.workbook_dataset is not None
    assert len(record.workbook_dataset.sheets) >= 1


def test_normalization_raises_404_for_unknown_session(session_with_file):
    _, norm_svc = session_with_file
    with pytest.raises(HTTPException) as exc_info:
        norm_svc.normalize("ghost-session")
    assert exc_info.value.status_code == 404


def test_normalization_raises_500_for_invalid_working_copy(tmp_path):
    svc = SessionService()
    record = SessionRecord(
        session_id="bad-session",
        source_file_path="uploads/bad-session.xlsx",
        working_copy_path=str(tmp_path / "nonexistent.xlsx"),
        original_filename="test.xlsx",
        status="uploaded",
    )
    svc.create(record)
    norm_svc = NormalizationService(svc)
    with pytest.raises(HTTPException) as exc_info:
        norm_svc.normalize("bad-session")
    assert exc_info.value.status_code == 500
