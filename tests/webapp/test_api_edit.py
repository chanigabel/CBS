"""Unit tests for the PATCH /api/workbook/{session_id}/sheet/{sheet_name}/cell endpoint."""

from tests.webapp.conftest import make_xlsx_bytes


def upload_and_get_sheet(client):
    """Upload a file and return (session_id, sheet_name, first_field_name, first_row_uid).

    Picks the first field that is actually present in the row data (skipping
    derived display columns like _serial and MosadID that are not in sheet.rows).
    """
    file_bytes = make_xlsx_bytes(["Sheet1"])
    response = client.post(
        "/api/upload",
        files={"file": ("test.xlsx", file_bytes, "application/octet-stream")},
    )
    assert response.status_code == 200
    session_id = response.json()["session_id"]

    # Get sheet data to find field names and row UIDs
    sheet_response = client.get(f"/api/workbook/{session_id}/sheet/Sheet1")
    assert sheet_response.status_code == 200
    data = sheet_response.json()
    rows = data["rows"]
    field_names = data["field_names"]

    # Pick the first field that actually exists in the raw sheet rows
    # (skip synthetic derived columns like _serial and MosadID that are
    # injected into display rows but not stored in sheet.rows)
    _DERIVED_COLS = {'_serial', 'MosadID'}
    real_field = None
    if rows:
        for fn in field_names:
            if fn not in _DERIVED_COLS and fn in rows[0]:
                real_field = fn
                break
    if real_field is None:
        real_field = field_names[0]

    first_row_uid = rows[0]["_row_uid"] if rows else None

    return session_id, "Sheet1", real_field, first_row_uid


def test_valid_edit_returns_200_with_updated_row(client):
    session_id, sheet_name, field_name, row_uid = upload_and_get_sheet(client)
    response = client.patch(
        f"/api/workbook/{session_id}/sheet/{sheet_name}/cell",
        json={"row_uid": row_uid, "field_name": field_name, "new_value": "NewValue"},
    )
    assert response.status_code == 200
    data = response.json()
    assert data["row_uid"] == row_uid
    assert data["updated_row"][field_name] == "NewValue"


def test_unknown_row_uid_returns_404(client):
    session_id, sheet_name, field_name, _ = upload_and_get_sheet(client)
    response = client.patch(
        f"/api/workbook/{session_id}/sheet/{sheet_name}/cell",
        json={"row_uid": "nonexistent-uid-99999999999999999999", "field_name": field_name, "new_value": "X"},
    )
    assert response.status_code == 404


def test_unknown_field_name_returns_400(client):
    session_id, sheet_name, _, row_uid = upload_and_get_sheet(client)
    response = client.patch(
        f"/api/workbook/{session_id}/sheet/{sheet_name}/cell",
        json={"row_uid": row_uid, "field_name": "nonexistent_field", "new_value": "X"},
    )
    assert response.status_code == 400


def test_edit_returns_404_for_unknown_session(client):
    response = client.patch(
        "/api/workbook/ghost-session/sheet/Sheet1/cell",
        json={"row_uid": "some-uid-12345678901234567890", "field_name": "first_name", "new_value": "X"},
    )
    assert response.status_code == 404
