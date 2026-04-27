"""Unit tests for the POST /api/workbook/{session_id}/normalize endpoint."""

from tests.webapp.conftest import make_xlsx_bytes


def upload_file(client, sheet_names=None):
    file_bytes = make_xlsx_bytes(sheet_names or ["Sheet1"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    assert response.status_code == 200
    return response.json()["session_id"]


def test_normalize_returns_200_with_stats(client):
    session_id = upload_file(client)
    response = client.post(f"/api/workbook/{session_id}/normalize")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "normalized"
    assert data["session_id"] == session_id
    assert "sheets_processed" in data
    assert "total_rows" in data
    assert "per_sheet_stats" in data


def test_normalize_returns_404_for_unknown_session(client):
    response = client.post("/api/workbook/ghost-session/normalize")
    assert response.status_code == 404
