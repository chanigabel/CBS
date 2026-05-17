from pathlib import Path

from fastapi.testclient import TestClient

from tests.webapp.conftest import make_xlsx_bytes
from webapp.services.cleanup_service import CleanupService
from webapp.services.edit_service import EditService
from webapp.services.export_service import ExportService
from webapp.services.processing_report_service import ProcessingReportService
from webapp.services.session_service import SessionService
from webapp.services.standardization_service import StandardizationService
from webapp.services.upload_service import UploadService
from webapp.services.workbook_service import WorkbookService


def test_cleanup_deletes_files_inside_allowed_runtime_directories(tmp_path):
    uploads = tmp_path / "uploads"
    work = tmp_path / "work"
    output = tmp_path / "output"
    nested = work / "session-1"

    uploads.mkdir()
    work.mkdir()
    output.mkdir()
    nested.mkdir()
    (uploads / "source.xlsx").write_bytes(b"upload")
    (nested / "working.xlsx").write_bytes(b"work")
    (output / "export.xlsx.tmp").write_bytes(b"export")

    CleanupService([uploads, work, output]).cleanup_runtime_files(reason="test")

    assert uploads.exists()
    assert work.exists()
    assert output.exists()
    assert list(uploads.iterdir()) == []
    assert list(work.iterdir()) == []
    assert list(output.iterdir()) == []


def test_cleanup_does_not_delete_files_outside_allowed_directories(tmp_path):
    uploads = tmp_path / "uploads"
    external = tmp_path / "outside"
    uploads.mkdir()
    external.mkdir()
    outside_file = external / "original.xlsx"
    outside_file.write_bytes(b"real user file")
    (uploads / "source.xlsx").write_bytes(b"upload")

    CleanupService([uploads]).cleanup_runtime_files(reason="test")

    assert outside_file.exists()
    assert outside_file.read_bytes() == b"real user file"
    assert list(uploads.iterdir()) == []


def test_cleanup_handles_missing_directories_safely(tmp_path):
    missing = tmp_path / "uploads"

    CleanupService([missing]).cleanup_runtime_files(reason="test")

    assert not missing.exists()


def test_cleanup_refuses_downloads_desktop_documents_and_home(tmp_path, monkeypatch):
    fake_home = tmp_path / "home"
    protected_dirs = [
        fake_home,
        fake_home / "Downloads",
        fake_home / "Downloads" / "app-runtime",
        fake_home / "Desktop",
        fake_home / "Desktop" / "app-runtime",
        fake_home / "Documents",
        fake_home / "Documents" / "app-runtime",
    ]

    for directory in protected_dirs:
        directory.mkdir(parents=True, exist_ok=True)
        (directory / "must-stay.xlsx").write_bytes(b"user file")

    monkeypatch.setattr(Path, "home", staticmethod(lambda: fake_home))

    CleanupService(protected_dirs).cleanup_runtime_files(reason="test")

    for directory in protected_dirs:
        assert (directory / "must-stay.xlsx").exists()


def test_cleanup_refuses_symlink_that_resolves_outside_allowlist(tmp_path):
    uploads = tmp_path / "uploads"
    external = tmp_path / "external"
    uploads.mkdir()
    external.mkdir()
    external_target = external / "original.xlsx"
    external_target.write_bytes(b"user file")

    link = uploads / "linked-original.xlsx"
    try:
        link.symlink_to(external_target)
    except OSError:
        return

    CleanupService([uploads]).cleanup_runtime_files(reason="test")

    assert link.is_symlink()
    assert external_target.exists()
    assert external_target.read_bytes() == b"user file"


def test_browser_refresh_like_request_does_not_clean_session_files(client):
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", make_xlsx_bytes(["Sheet1"]), "application/octet-stream")},
    )
    assert response.status_code == 200
    session_id = response.json()["session_id"]

    first_summary = client.get(f"/api/workbook/{session_id}/summary")
    refresh_like_request = client.get("/")
    second_summary = client.get(f"/api/workbook/{session_id}/summary")

    assert first_summary.status_code == 200
    assert refresh_like_request.status_code == 200
    assert second_summary.status_code == 200
    assert second_summary.json()["session_id"] == session_id


def test_startup_cleanup_only_removes_stale_internal_runtime_files(tmp_path, monkeypatch):
    import webapp.dependencies as deps
    from webapp.app import app

    uploads = tmp_path / "uploads"
    work = tmp_path / "work"
    output = tmp_path / "output"
    external = tmp_path / "outside"
    for directory in (uploads, work, output, external):
        directory.mkdir()

    (uploads / "stale-upload.xlsx").write_bytes(b"upload")
    (work / "stale-work.xlsx").write_bytes(b"work")
    (output / "stale-export.xlsx").write_bytes(b"output")
    external_file = external / "original.xlsx"
    external_file.write_bytes(b"real user file")

    svc = SessionService()
    report_svc = ProcessingReportService(svc)
    upload_svc = UploadService(svc, uploads, work, report_svc)
    workbook_svc = WorkbookService(svc)
    norm_svc = StandardizationService(svc, report_svc)
    edit_svc = EditService(svc)
    export_svc = ExportService(svc, output, report_svc)

    monkeypatch.setattr(deps, "_session_service", svc)
    monkeypatch.setattr(deps, "_upload_service", upload_svc)
    monkeypatch.setattr(deps, "_workbook_service", workbook_svc)
    monkeypatch.setattr(deps, "_standardization_service", norm_svc)
    monkeypatch.setattr(deps, "_edit_service", edit_svc)
    monkeypatch.setattr(deps, "_export_service", export_svc)

    with TestClient(app) as test_client:
        assert test_client.get("/").status_code == 200
        assert list(uploads.iterdir()) == []
        assert list(work.iterdir()) == []
        assert list(output.iterdir()) == []
        assert external_file.exists()
