"""Unit tests for SessionService."""

import pytest
from fastapi import HTTPException
from webapp.services.session_service import SessionService
from webapp.models.session import SessionRecord


def make_record(session_id: str = "test-session-1") -> SessionRecord:
    return SessionRecord(
        session_id=session_id,
        source_file_path=f"uploads/{session_id}.xlsx",
        working_copy_path=f"work/{session_id}.xlsx",
        original_filename="test.xlsx",
        status="uploaded",
    )


@pytest.fixture(autouse=True)
def clear_registry():
    """Clear the session registry before each test."""
    svc = SessionService()
    svc.clear_all()
    yield
    svc.clear_all()


def test_create_then_get_returns_same_record():
    svc = SessionService()
    record = make_record("abc-123")
    svc.create(record)
    retrieved = svc.get("abc-123")
    assert retrieved is record
    assert retrieved.session_id == "abc-123"
    assert retrieved.status == "uploaded"


def test_get_unknown_session_raises_404():
    svc = SessionService()
    with pytest.raises(HTTPException) as exc_info:
        svc.get("nonexistent-uuid")
    assert exc_info.value.status_code == 404


def test_update_mutates_stored_record():
    svc = SessionService()
    record = make_record("upd-session")
    svc.create(record)
    svc.update("upd-session", status="normalized")
    updated = svc.get("upd-session")
    assert updated.status == "normalized"


def test_update_nonexistent_session_raises_404():
    svc = SessionService()
    with pytest.raises(HTTPException) as exc_info:
        svc.update("ghost-session", status="normalized")
    assert exc_info.value.status_code == 404


def test_delete_removes_session():
    svc = SessionService()
    record = make_record("del-session")
    svc.create(record)
    svc.delete("del-session")
    with pytest.raises(HTTPException):
        svc.get("del-session")


def test_multiple_sessions_are_independent():
    svc = SessionService()
    r1 = make_record("session-1")
    r2 = make_record("session-2")
    svc.create(r1)
    svc.create(r2)
    svc.update("session-1", status="normalized")
    assert svc.get("session-1").status == "normalized"
    assert svc.get("session-2").status == "uploaded"
