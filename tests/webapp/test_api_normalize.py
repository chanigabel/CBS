"""Unit tests for the POST /api/workbook/{session_id}/normalize endpoint."""

from tests.webapp.conftest import make_xlsx_bytes
from openpyxl import Workbook


def make_mapping_xlsx_bytes() -> bytes:
    from io import BytesIO

    wb = Workbook()
    ws = wb.active
    ws.title = "People"
    ws.append(["custom_a", "custom_b", "gender"])
    ws.append([" Alice ", "Smith", "F"])
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


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
    assert data["status"] == "standardized"
    assert data["session_id"] == session_id
    assert "sheets_processed" in data
    assert "total_rows" in data
    assert "per_sheet_stats" in data


def test_normalize_returns_404_for_unknown_session(client):
    response = client.post("/api/workbook/ghost-session/normalize")
    assert response.status_code == 404


def test_normalize_blocks_duplicate_mappings_then_runs_after_fix(client):
    upload_response = client.post(
        "/api/upload",
        files={"file": ("mapping.xlsx", make_mapping_xlsx_bytes(), "application/octet-stream")},
    )
    assert upload_response.status_code == 200
    session_id = upload_response.json()["session_id"]

    assert client.get(f"/api/workbook/{session_id}/sheet/People").status_code == 200
    assert client.post(
        f"/api/workbook/{session_id}/sheet/People/column-mapping",
        json={"old_name": "custom_a", "new_name": "first_name"},
    ).status_code == 200
    assert client.post(
        f"/api/workbook/{session_id}/sheet/People/column-mapping",
        json={"old_name": "custom_b", "new_name": "first_name"},
    ).status_code == 200

    blocked_response = client.post(f"/api/workbook/{session_id}/normalize?sheet=People")
    assert blocked_response.status_code == 400
    assert "Cannot start standardization" in blocked_response.json()["detail"]

    assert client.post(
        f"/api/workbook/{session_id}/sheet/People/column-mapping",
        json={"old_name": "custom_b", "new_name": "last_name"},
    ).status_code == 200

    normalized_response = client.post(f"/api/workbook/{session_id}/normalize?sheet=People")
    assert normalized_response.status_code == 200

    sheet_response = client.get(f"/api/workbook/{session_id}/sheet/People")
    assert sheet_response.status_code == 200
    data = sheet_response.json()
    assert "first_name" in data["field_names"]
    assert data["rows"][0]["first_name"] == " Alice "
    assert data["rows"][0]["first_name_corrected"] == "Alice"
