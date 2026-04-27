"""Shared fixtures for webapp API tests."""

import io
import pytest
from pathlib import Path
from openpyxl import Workbook
from fastapi.testclient import TestClient

from webapp.services.session_service import SessionService


def make_xlsx_bytes(sheet_names=None) -> bytes:
    """Create a minimal valid .xlsx file in memory."""
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
def clear_session_registry():
    """Clear the session registry before each test."""
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


@pytest.fixture
def client(tmp_path, monkeypatch):
    """Create a TestClient with patched runtime directories."""
    # Patch the runtime directories to use tmp_path
    import webapp.dependencies as deps
    import webapp.services.upload_service as us

    monkeypatch.setattr(deps, "UPLOADS_DIR", tmp_path / "uploads")
    monkeypatch.setattr(deps, "WORK_DIR", tmp_path / "work")
    monkeypatch.setattr(deps, "OUTPUT_DIR", tmp_path / "output")

    # Rebuild service instances with patched dirs
    svc = SessionService()
    from webapp.services.upload_service import UploadService
    from webapp.services.workbook_service import WorkbookService
    from webapp.services.normalization_service import NormalizationService
    from webapp.services.edit_service import EditService
    from webapp.services.export_service import ExportService

    upload_svc = UploadService(svc, tmp_path / "uploads", tmp_path / "work")
    workbook_svc = WorkbookService(svc)
    norm_svc = NormalizationService(svc)
    edit_svc = EditService(svc)
    export_svc = ExportService(svc, tmp_path / "output")

    monkeypatch.setattr(deps, "_session_service", svc)
    monkeypatch.setattr(deps, "_upload_service", upload_svc)
    monkeypatch.setattr(deps, "_workbook_service", workbook_svc)
    monkeypatch.setattr(deps, "_normalization_service", norm_svc)
    monkeypatch.setattr(deps, "_edit_service", edit_svc)
    monkeypatch.setattr(deps, "_export_service", export_svc)

    from webapp.app import app
    with TestClient(app) as c:
        yield c
