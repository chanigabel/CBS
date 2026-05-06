"""Tests for POST /api/process-file."""

import io

from openpyxl import load_workbook

from tests.webapp.conftest import make_xlsx_bytes


def test_process_file_returns_exported_workbook(client):
    file_bytes = make_xlsx_bytes(["Sheet1"])

    response = client.post(
        "/api/process-file",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "attachment;" in response.headers["content-disposition"]

    wb = load_workbook(io.BytesIO(response.content))
    assert wb.sheetnames
    wb.close()


def test_process_file_reuses_existing_services(client, monkeypatch):
    import webapp.dependencies as deps

    file_bytes = make_xlsx_bytes(["Sheet1"])
    calls = []

    original_upload = deps._upload_service.handle_upload
    original_standardize = deps._standardization_service.standardize
    original_export = deps._export_service.export

    def spy_upload(filename, uploaded_bytes):
        calls.append("upload")
        return original_upload(filename, uploaded_bytes)

    def spy_standardize(session_id, sheet_name=None):
        calls.append(("standardize", session_id, sheet_name))
        return original_standardize(session_id, sheet_name=sheet_name)

    def spy_export(session_id):
        calls.append(("export", session_id))
        return original_export(session_id)

    monkeypatch.setattr(deps._upload_service, "handle_upload", spy_upload)
    monkeypatch.setattr(deps._standardization_service, "standardize", spy_standardize)
    monkeypatch.setattr(deps._export_service, "export", spy_export)

    response = client.post(
        "/api/process-file",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )

    assert response.status_code == 200
    assert calls[0] == "upload"
    assert calls[1][0] == "standardize"
    assert calls[1][2] is None
    assert calls[2] == ("export", calls[1][1])


def test_process_file_invalid_extension_returns_400(client):
    response = client.post(
        "/api/process-file",
        files={"file": ("test.csv", b"col1,col2\nval1,val2", "text/csv")},
    )

    assert response.status_code == 400
    assert "xlsx" in response.json()["detail"].lower()
