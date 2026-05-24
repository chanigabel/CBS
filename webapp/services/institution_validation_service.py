"""Service for validating institution codes and types against a DB.

This service provides infrastructure for validating institution-related data.
When a DB connection and data are available, it validates against DB records.
When not available, it returns appropriate fallback status without blocking.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from enum import Enum
from typing import Optional, Callable, Any

logger = logging.getLogger(__name__)


class ValidationStatus(str, Enum):
    """Status of a validation check."""
    VALID = "valid"
    INVALID = "invalid"
    UNKNOWN = "unknown"  # DB not available or validation skipped
    NOT_CHECKED = "not_checked"  # Validation not requested


@dataclass
class ValidationResult:
    """Result of a single validation check.

    Attributes:
        status: ValidationStatus indicating result
        field_name: Name of the field being validated (e.g., 'institution_code')
        value: The value that was validated
        message: Human-readable message (Hebrew-ready)
        is_warning: True if this is a warning (validation skipped), not an error
    """
    status: ValidationStatus
    field_name: str
    value: Any
    message: str = ""
    is_warning: bool = False

    def is_valid(self) -> bool:
        """Return True if validation passed."""
        return self.status == ValidationStatus.VALID

    def is_warning_or_unknown(self) -> bool:
        """Return True if status is UNKNOWN or this is a warning."""
        return self.status == ValidationStatus.UNKNOWN or self.is_warning


class InstitutionValidationService:
    """Service layer for validating institution codes and types.

    This service encapsulates institution validation logic. It supports:
    - Checking if an institution code exists
    - Checking if an institution type exists
    - Validating institution type matches code (future)
    - Safe fallback when DB is not available
    - Clear status reporting for integration with processing reports
    """

    def __init__(
        self,
        institution_code_validator: Optional[Callable[[str], bool]] = None,
        institution_type_validator: Optional[Callable[[str], bool]] = None,
    ):
        """Initialize the validation service.

        Args:
            institution_code_validator: Optional callable that takes a code string
                and returns True if valid. If None, DB is considered unavailable.
            institution_type_validator: Optional callable that takes a type string
                and returns True if valid. If None, DB is considered unavailable.
        """
        self.institution_code_validator = institution_code_validator
        self.institution_type_validator = institution_type_validator
        self._db_available = (
            institution_code_validator is not None or
            institution_type_validator is not None
        )

    @property
    def db_available(self) -> bool:
        """Return True if DB validators are configured."""
        return self._db_available

    def validate_institution_code(self, code: str) -> ValidationResult:
        """Validate that an institution code exists.

        Args:
            code: Institution code to validate (usually numeric string)

        Returns:
            ValidationResult with status, message, and is_warning flag
        """
        if not code or not str(code).strip():
            return ValidationResult(
                status=ValidationStatus.INVALID,
                field_name="institution_code",
                value=code,
                message="קוד המוסד ריק או חסר",
                is_warning=False,
            )

        if not self._db_available or self.institution_code_validator is None:
            return ValidationResult(
                status=ValidationStatus.UNKNOWN,
                field_name="institution_code",
                value=code,
                message="לא ניתן לאמת קוד מוסד - אין חיבור לבסיס נתונים",
                is_warning=True,
            )

        try:
            is_valid = self.institution_code_validator(str(code).strip())
            if is_valid:
                return ValidationResult(
                    status=ValidationStatus.VALID,
                    field_name="institution_code",
                    value=code,
                    message="",
                    is_warning=False,
                )
            else:
                return ValidationResult(
                    status=ValidationStatus.INVALID,
                    field_name="institution_code",
                    value=code,
                    message=f"קוד מוסד '{code}' לא נמצא בבסיס הנתונים",
                    is_warning=False,
                )
        except Exception as e:
            logger.warning(
                "institution_code_validation_error",
                extra={
                    "event": "institution_code_validation_error",
                    "code": code,
                    "error": str(e),
                },
            )
            return ValidationResult(
                status=ValidationStatus.UNKNOWN,
                field_name="institution_code",
                value=code,
                message="שגיאה בזמן אימות קוד המוסד",
                is_warning=True,
            )

    def validate_institution_type(self, type_val: str) -> ValidationResult:
        """Validate that an institution type exists.

        Args:
            type_val: Institution type to validate (usually numeric string)

        Returns:
            ValidationResult with status, message, and is_warning flag
        """
        if not type_val or not str(type_val).strip():
            return ValidationResult(
                status=ValidationStatus.INVALID,
                field_name="institution_type",
                value=type_val,
                message="סוג המוסד ריק או חסר",
                is_warning=False,
            )

        if not self._db_available or self.institution_type_validator is None:
            return ValidationResult(
                status=ValidationStatus.UNKNOWN,
                field_name="institution_type",
                value=type_val,
                message="לא ניתן לאמת סוג מוסד - אין חיבור לבסיס נתונים",
                is_warning=True,
            )

        try:
            is_valid = self.institution_type_validator(str(type_val).strip())
            if is_valid:
                return ValidationResult(
                    status=ValidationStatus.VALID,
                    field_name="institution_type",
                    value=type_val,
                    message="",
                    is_warning=False,
                )
            else:
                return ValidationResult(
                    status=ValidationStatus.INVALID,
                    field_name="institution_type",
                    value=type_val,
                    message=f"סוג מוסד '{type_val}' לא נמצא בבסיס הנתונים",
                    is_warning=False,
                )
        except Exception as e:
            logger.warning(
                "institution_type_validation_error",
                extra={
                    "event": "institution_type_validation_error",
                    "type_val": type_val,
                    "error": str(e),
                },
            )
            return ValidationResult(
                status=ValidationStatus.UNKNOWN,
                field_name="institution_type",
                value=type_val,
                message="שגיאה בזמן אימות סוג המוסד",
                is_warning=True,
            )

    def validate_institution_type_matches_code(
        self, code: str, type_val: str
    ) -> ValidationResult:
        """Validate that an institution type matches the given code.

        This is a future extension point. Currently not implemented.

        Args:
            code: Institution code
            type_val: Institution type to check

        Returns:
            ValidationResult with match status
        """
        # Future implementation: check DB for code-type relationship
        return ValidationResult(
            status=ValidationStatus.NOT_CHECKED,
            field_name="institution_type_matches_code",
            value=(code, type_val),
            message="אימות התאמה בין קוד וסוג מוסד עדיין לא מיושם",
            is_warning=True,
        )

    def create_default_service() -> InstitutionValidationService:
        """Create a default service with no DB validators (safe fallback mode)."""
        return InstitutionValidationService(
            institution_code_validator=None,
            institution_type_validator=None,
        )
