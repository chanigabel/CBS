"""Unit tests for the POST /api/upload endpoint."""

import pytest
from tests.webapp.conftest import make_xlsx_bytes


def test_valid_upload_returns_200_with_session_id_and_sheet_names(client):
    file_bytes = make_xlsx_bytes(["Sheet1", "Sheet2"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    assert response.status_code == 200
    data = response.json()
    assert "session_id" in data
    assert set(data["sheet_names"]) == {"Sheet1", "Sheet2"}


def test_invalid_extension_returns_400(client):
    response = client.post(
        "/api/upload",
        files={"file": ("test.csv", b"col1,col2\nval1,val2", "text/csv")},
    )
    assert response.status_code == 400
    assert "xlsx" in response.json()["detail"].lower()


def test_corrupted_file_returns_422(client):
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", b"not a valid xlsx", "application/octet-stream")},
    )
    assert response.status_code == 422


def test_xlsm_extension_is_accepted(client):
    file_bytes = make_xlsx_bytes(["Sheet1"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsm", file_bytes, "application/octet-stream")},
    )
    assert response.status_code == 200
    assert "session_id" in response.json()
