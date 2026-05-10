"""SessionService: in-memory session registry for tracking upload/processing state."""

import logging
from typing import Dict
from fastapi import HTTPException
from webapp.models.session import SessionRecord

logger = logging.getLogger(__name__)

# Module-level registry — single shared instance for the process lifetime.
# No locking needed for single-threaded Uvicorn.
_registry: Dict[str, SessionRecord] = {}


# מנהל את מצב ה־session בזיכרון עבור הקובץ הפעיל של המשתמש.
class SessionService:
    """CRUD operations on the in-memory session registry.

    The registry is a module-level dict so all instances share the same state.
    This matches the design requirement for a single process-lifetime store.
    """

    # שומר session חדש שנוצר בשלב upload.
    def create(self, record: SessionRecord) -> None:
        """Store a new session record.

        Args:
            record: The SessionRecord to store
        """
        _registry[record.session_id] = record
        logger.info(f"Session created: {record.session_id}")

    # מחזיר session קיים עבור פעולות workbook/standardize/export.
    def get(self, session_id: str) -> SessionRecord:
        """Retrieve a session record by ID.

        Args:
            session_id: UUID string of the session

        Returns:
            The SessionRecord for this session

        Raises:
            HTTPException: 404 if session_id is not found
        """
        record = _registry.get(session_id)
        if record is None:
            logger.info(f"Session not found: {session_id}")
            raise HTTPException(
                status_code=404,
                detail=f"Session '{session_id}' not found. Please upload a file first.",
            )
        return record

    # מעדכן שדות session לאחר חילוץ, עריכה או סטנדרטיזציה.
    def update(self, session_id: str, **kwargs) -> None:
        """Update fields on an existing session record.

        Args:
            session_id: UUID string of the session to update
            **kwargs: Field names and new values to set

        Raises:
            HTTPException: 404 if session_id is not found
        """
        record = self.get(session_id)
        for key, value in kwargs.items():
            if hasattr(record, key):
                setattr(record, key, value)
            else:
                logger.warning(f"SessionRecord has no attribute '{key}' — skipping")
        logger.debug(f"Session updated: {session_id}, fields: {list(kwargs.keys())}")

    # מוחק session ואת קבצי העבודה שלו כשאין בהם צורך.
    def delete(self, session_id: str) -> None:
        """Remove a session record from the registry.

        Args:
            session_id: UUID string of the session to delete
        """
        if session_id in _registry:
            del _registry[session_id]
            logger.info(f"Session deleted: {session_id}")

    # מנקה את כל ה־sessions, בעיקר לשימושי בדיקה או איפוס מקומי.
    def clear_all(self) -> None:
        """Remove all sessions (used in tests)."""
        _registry.clear()
