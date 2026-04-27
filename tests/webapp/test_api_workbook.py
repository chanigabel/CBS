"""Unit tests for workbook API endpoints."""

import pytest
from tests.webapp.conftest import make_xlsx_bytes


def upload_file(client, sheet_names=None):
    """Helper: upload a file and return the session_id."""
    file_bytes = make_xlsx_bytes(sheet_names or ["Sheet1"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    assert response.status_code == 200
    return response.json()["session_id"]


def test_summary_returns_correct_structure(client):
    session_id = upload_file(client, ["Sheet1", "Sheet2"])
    response = client.get(f"/api/workbook/{session_id}/summary")
    assert response.status_code == 200
    data = response.json()
    assert data["session_id"] == session_id
    assert len(data["sheets"]) == 2
    sheet_names = [s["sheet_name"] for s in data["sheets"]]
    assert "Sheet1" in sheet_names
    assert "Sheet2" in sheet_names


def test_summary_returns_404_for_unknown_session(client):
    response = client.get("/api/workbook/nonexistent-session/summary")
    assert response.status_code == 404


def test_sheet_data_returns_rows_for_valid_sheet(client):
    session_id = upload_file(client, ["Sheet1"])
    response = client.get(f"/api/workbook/{session_id}/sheet/Sheet1")
    assert response.status_code == 200
    data = response.json()
    assert data["sheet_name"] == "Sheet1"
    assert len(data["rows"]) >= 1
    assert "field_names" in data


def test_sheet_data_returns_404_for_unknown_sheet(client):
    session_id = upload_file(client)
    response = client.get(f"/api/workbook/{session_id}/sheet/NonExistentSheet")
    assert response.status_code == 404


def test_sheet_data_returns_404_for_unknown_session(client):
    response = client.get("/api/workbook/ghost-session/sheet/Sheet1")
    assert response.status_code == 404
