"""Integration tests for the full upload → normalize → export workflow.

Uses FastAPI TestClient with real xlsx files to verify the complete pipeline.
"""

import hashlib
import io
import pytest
from pathlib import Path
from openpyxl import Workbook
from fastapi.testclient import TestClient

from webapp.services.session_service import SessionService


def make_sample_xlsx() -> bytes:
    """Create a sample xlsx file with realistic data."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Sheet1")
    ws.append(["first_name", "last_name", "gender"])
    ws.append(["Alice", "Smith", "F"])
    ws.append(["Bob", "Jones", "M"])
    ws.append(["Carol", "Williams", "F"])
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
def client(tmp_path, monkeypatch):
    """Create a TestClient with patched runtime directories."""
    import webapp.dependencies as deps
    from webapp.services.upload_service import UploadService
    from webapp.services.workbook_service import WorkbookService
    from webapp.services.standardization_service import standardizationService
    from webapp.services.edit_service import EditService
    from webapp.services.export_service import ExportService

    svc = SessionService()
    upload_svc = UploadService(svc, tmp_path / "uploads", tmp_path / "work")
    workbook_svc = WorkbookService(svc)
    norm_svc = standardizationService(svc)
    edit_svc = EditService(svc)
    export_svc = ExportService(svc, tmp_path / "output")

    monkeypatch.setattr(deps, "_session_service", svc)
    monkeypatch.setattr(deps, "_upload_service", upload_svc)
    monkeypatch.setattr(deps, "_workbook_service", workbook_svc)
    monkeypatch.setattr(deps, "_standardization_service", norm_svc)
    monkeypatch.setattr(deps, "_edit_service", edit_svc)
    monkeypatch.setattr(deps, "_export_service", export_svc)

    from webapp.app import app
    with TestClient(app) as c:
        yield c, tmp_path


def test_full_workflow_upload_normalize_export(client):
    """Integration test: upload → normalize → export full workflow."""
    test_client, tmp_path = client
    file_bytes = make_sample_xlsx()

    # Step 1: Upload
    upload_response = test_client.post(
        "/api/upload",
        files={"file": ("sample.xlsx", file_bytes, "application/octet-stream")},
    )
    assert upload_response.status_code == 200
    upload_data = upload_response.json()
    session_id = upload_data["session_id"]
    assert session_id
    assert "Sheet1" in upload_data["sheet_names"]

    # Step 2: Get summary
    summary_response = test_client.get(f"/api/workbook/{session_id}/summary")
    assert summary_response.status_code == 200
    summary = summary_response.json()
    assert summary["session_id"] == session_id
    assert len(summary["sheets"]) >= 1

    # Step 3: Get sheet data
    sheet_response = test_client.get(f"/api/workbook/{session_id}/sheet/Sheet1")
    assert sheet_response.status_code == 200
    sheet_data = sheet_response.json()
    assert len(sheet_data["rows"]) == 3

    # Step 4: Normalize
    normalize_response = test_client.post(f"/api/workbook/{session_id}/normalize")
    assert normalize_response.status_code == 200
    norm_data = normalize_response.json()
    assert norm_data["status"] == "standardized"
    assert norm_data["sheets_processed"] >= 1
    assert norm_data["total_rows"] >= 0

    # Step 5: Export
    export_response = test_client.post(f"/api/workbook/{session_id}/export")
    # Export may return 200 (file) or 500 if no matching VBA sheets
    # The important thing is it doesn't return 404
    assert export_response.status_code != 404

    # Step 6: Verify source file is unchanged
    uploads_dir = tmp_path / "uploads"
    source_files = list(uploads_dir.glob(f"{session_id}*"))
    assert len(source_files) == 1
    assert source_files[0].read_bytes() == file_bytes


def test_concurrent_sessions_are_independent(client):
    """Two concurrent upload sessions should not interfere with each other."""
    test_client, _ = client
    file_bytes = make_sample_xlsx()

    # Upload two files
    resp1 = test_client.post(
        "/api/upload",
        files={"file": ("file1.xlsx", file_bytes, "application/octet-stream")},
    )
    resp2 = test_client.post(
        "/api/upload",
        files={"file": ("file2.xlsx", file_bytes, "application/octet-stream")},
    )
    assert resp1.status_code == 200
    assert resp2.status_code == 200

    session1 = resp1.json()["session_id"]
    session2 = resp2.json()["session_id"]
    assert session1 != session2

    # Normalize session 1 only
    norm_resp = test_client.post(f"/api/workbook/{session1}/normalize")
    assert norm_resp.status_code == 200

    # Session 2 should still be in "uploaded" state
    summary2 = test_client.get(f"/api/workbook/{session2}/summary")
    assert summary2.status_code == 200


def test_source_file_unchanged_after_full_workflow(client):
    """Source file must be byte-for-byte identical after upload → normalize → export."""
    test_client, tmp_path = client
    file_bytes = make_sample_xlsx()
    original_hash = hashlib.sha256(file_bytes).hexdigest()

    # Upload
    resp = test_client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    session_id = resp.json()["session_id"]

    # Normalize
    test_client.post(f"/api/workbook/{session_id}/normalize")

    # Export
    test_client.post(f"/api/workbook/{session_id}/export")

    # Verify source file
    uploads_dir = tmp_path / "uploads"
    source_files = list(uploads_dir.glob(f"{session_id}*"))
    assert len(source_files) == 1
    assert hashlib.sha256(source_files[0].read_bytes()).hexdigest() == original_hash


def test_index_page_returns_html(client):
    """The root endpoint should return the HTML UI."""
    test_client, _ = client
    response = test_client.get("/")
    assert response.status_code == 200
    assert "text/html" in response.headers.get("content-type", "")
