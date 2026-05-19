"""Unit tests for workbook API endpoints."""

import pytest
from openpyxl import Workbook
from tests.webapp.conftest import make_xlsx_bytes


def make_identifier_only_xlsx_bytes() -> bytes:
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["id_number"])
    ws.append(["ABC123"])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def make_mapping_xlsx_bytes() -> bytes:
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "People"
    ws.append(["custom_a", "custom_b", "gender"])
    ws.append(["Alice", "Smith", "F"])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


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


def test_sheet_data_includes_generated_passport_corrected_without_source_passport(client):
    response = client.post(
        "/api/upload",
        files={
            "file": (
                "identifier_only.xlsx",
                make_identifier_only_xlsx_bytes(),
                "application/octet-stream",
            )
        },
    )
    assert response.status_code == 200
    session_id = response.json()["session_id"]

    norm_response = client.post(f"/api/workbook/{session_id}/normalize")
    assert norm_response.status_code == 200

    sheet_response = client.get(f"/api/workbook/{session_id}/sheet/Sheet1")
    assert sheet_response.status_code == 200
    data = sheet_response.json()

    fields = data["field_names"]
    row = data["rows"][0]
    id_index = fields.index("id_number")
    assert fields[id_index + 1] == "id_number_corrected"
    assert fields[id_index + 2] == "passport_corrected"
    assert "passport" not in fields
    assert "passport" not in row
    assert row["id_number"] == "ABC123"
    assert row["id_number_corrected"] == ""
    assert row["passport_corrected"] == "ABC123"


def test_column_mapping_edit_allows_duplicates_and_grid_shows_effective_mapping(client):
    response = client.post(
        "/api/upload",
        files={"file": ("mapping.xlsx", make_mapping_xlsx_bytes(), "application/octet-stream")},
    )
    assert response.status_code == 200
    session_id = response.json()["session_id"]

    sheet_response = client.get(f"/api/workbook/{session_id}/sheet/People")
    assert sheet_response.status_code == 200

    first_response = client.post(
        f"/api/workbook/{session_id}/sheet/People/column-mapping",
        json={"old_name": "custom_a", "new_name": "first_name"},
    )
    assert first_response.status_code == 200
    duplicate_response = client.post(
        f"/api/workbook/{session_id}/sheet/People/column-mapping",
        json={"old_name": "custom_b", "new_name": "first_name"},
    )
    assert duplicate_response.status_code == 200

    data_response = client.get(f"/api/workbook/{session_id}/sheet/People")
    assert data_response.status_code == 200
    data = data_response.json()

    assert "custom_a" in data["field_names"]
    assert "custom_b" in data["field_names"]
    assert "gender" in data["field_names"]
    assert data["rows"][0]["custom_a"] == "Alice"
    assert data["rows"][0]["custom_b"] == "Smith"
    assert data["column_mappings"] == {
        "custom_a": "first_name",
        "custom_b": "first_name",
    }
    assert data["column_display_names"]["custom_a"] == "custom_a \u2192 first_name"
    assert data["column_display_names"]["custom_b"] == "custom_b \u2192 first_name"
