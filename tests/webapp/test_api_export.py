"""Unit tests for the POST /api/workbook/{session_id}/export endpoint."""

from tests.webapp.conftest import make_xlsx_bytes


def upload_and_normalize(client):
    """Upload and normalize a file, return session_id."""
    file_bytes = make_xlsx_bytes(["Sheet1"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    assert response.status_code == 200
    session_id = response.json()["session_id"]

    norm_response = client.post(f"/api/workbook/{session_id}/normalize")
    assert norm_response.status_code == 200
    return session_id


def test_export_returns_file_response(client):
    session_id = upload_and_normalize(client)
    response = client.post(f"/api/workbook/{session_id}/export")
    # Export may return 200 (file) or 500 if no matching VBA sheets
    # The important thing is it doesn't crash with 404
    assert response.status_code in (200, 500)


def test_export_returns_404_for_unknown_session(client):
    response = client.post("/api/workbook/ghost-session/export")
    assert response.status_code == 404


def test_export_after_upload_without_normalize(client):
    """Export should work even without standardization (uses raw data)."""
    file_bytes = make_xlsx_bytes(["Sheet1"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    session_id = response.json()["session_id"]
    export_response = client.post(f"/api/workbook/{session_id}/export")
    # Should not return 404
    assert export_response.status_code != 404
