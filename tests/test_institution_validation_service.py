"""Tests for InstitutionValidationService."""

import pytest
from webapp.services.institution_validation_service import (
    InstitutionValidationService,
    ValidationStatus,
    ValidationResult,
)


class TestInstitutionValidationService:
    """Test suite for institution validation."""

    def test_validate_existing_institution_code(self):
        """Test validation of an existing institution code."""
        def mock_validator(code: str) -> bool:
            return code in ["12345", "67890"]

        service = InstitutionValidationService(
            institution_code_validator=mock_validator
        )
        result = service.validate_institution_code("12345")

        assert result.status == ValidationStatus.VALID
        assert result.is_valid() is True
        assert result.field_name == "institution_code"
        assert result.is_warning is False

    def test_validate_missing_institution_code(self):
        """Test validation of a missing institution code."""
        def mock_validator(code: str) -> bool:
            return code in ["12345", "67890"]

        service = InstitutionValidationService(
            institution_code_validator=mock_validator
        )
        result = service.validate_institution_code("99999")

        assert result.status == ValidationStatus.INVALID
        assert result.is_valid() is False
        assert "לא נמצא" in result.message or "99999" in result.message
        assert result.is_warning is False

    def test_validate_empty_institution_code(self):
        """Test validation rejects empty code."""
        def mock_validator(code: str) -> bool:
            return True

        service = InstitutionValidationService(
            institution_code_validator=mock_validator
        )
        result = service.validate_institution_code("")

        assert result.status == ValidationStatus.INVALID
        assert result.is_valid() is False
        assert "ריק" in result.message or "חסר" in result.message

    def test_validate_existing_institution_type(self):
        """Test validation of an existing institution type."""
        def mock_validator(type_val: str) -> bool:
            return type_val in ["1", "2", "3"]

        service = InstitutionValidationService(
            institution_type_validator=mock_validator
        )
        result = service.validate_institution_type("1")

        assert result.status == ValidationStatus.VALID
        assert result.is_valid() is True
        assert result.field_name == "institution_type"
        assert result.is_warning is False

    def test_validate_missing_institution_type(self):
        """Test validation of a missing institution type."""
        def mock_validator(type_val: str) -> bool:
            return type_val in ["1", "2", "3"]

        service = InstitutionValidationService(
            institution_type_validator=mock_validator
        )
        result = service.validate_institution_type("99")

        assert result.status == ValidationStatus.INVALID
        assert result.is_valid() is False
        assert "לא נמצא" in result.message or "99" in result.message

    def test_validate_empty_institution_type(self):
        """Test validation rejects empty type."""
        def mock_validator(type_val: str) -> bool:
            return True

        service = InstitutionValidationService(
            institution_type_validator=mock_validator
        )
        result = service.validate_institution_type("")

        assert result.status == ValidationStatus.INVALID
        assert result.is_valid() is False
        assert "ריק" in result.message or "חסר" in result.message

    def test_db_unavailable_returns_unknown_for_code(self):
        """Test that missing DB returns UNKNOWN status for code (not blocking)."""
        service = InstitutionValidationService(
            institution_code_validator=None,
            institution_type_validator=None,
        )
        result = service.validate_institution_code("12345")

        assert result.status == ValidationStatus.UNKNOWN
        assert result.is_warning is True
        assert "בסיס נתונים" in result.message or "אין" in result.message

    def test_db_unavailable_returns_unknown_for_type(self):
        """Test that missing DB returns UNKNOWN status for type (not blocking)."""
        service = InstitutionValidationService(
            institution_code_validator=None,
            institution_type_validator=None,
        )
        result = service.validate_institution_type("1")

        assert result.status == ValidationStatus.UNKNOWN
        assert result.is_warning is True
        assert "בסיס נתונים" in result.message or "אין" in result.message

    def test_db_unavailable_flag(self):
        """Test db_available property reflects validator availability."""
        # No validators
        service_no_db = InstitutionValidationService()
        assert service_no_db.db_available is False

        # With validators
        service_with_db = InstitutionValidationService(
            institution_code_validator=lambda x: True,
        )
        assert service_with_db.db_available is True

    def test_validation_result_is_warning_or_unknown(self):
        """Test is_warning_or_unknown method."""
        warning_result = ValidationResult(
            status=ValidationStatus.UNKNOWN,
            field_name="test",
            value="test",
            is_warning=True,
        )
        assert warning_result.is_warning_or_unknown() is True

        valid_result = ValidationResult(
            status=ValidationStatus.VALID,
            field_name="test",
            value="test",
            is_warning=False,
        )
        assert valid_result.is_warning_or_unknown() is False

    def test_validation_error_handling(self):
        """Test that exceptions in validators are caught and reported."""
        def failing_validator(code: str) -> bool:
            raise ValueError("DB connection failed")

        service = InstitutionValidationService(
            institution_code_validator=failing_validator
        )
        result = service.validate_institution_code("12345")

        assert result.status == ValidationStatus.UNKNOWN
        assert result.is_warning is True
        assert "שגיאה" in result.message

    def test_whitespace_trimming(self):
        """Test that validators receive trimmed whitespace."""
        received_values = []

        def capture_validator(value: str) -> bool:
            received_values.append(value)
            return True

        service = InstitutionValidationService(
            institution_code_validator=capture_validator
        )
        result = service.validate_institution_code("  12345  ")

        assert result.is_valid() is True
        assert "12345" in received_values
        assert len(received_values[0].strip()) > 0

    def test_type_matching_not_yet_implemented(self):
        """Test that type-code matching is a future feature."""
        service = InstitutionValidationService()
        result = service.validate_institution_type_matches_code("12345", "1")

        assert result.status == ValidationStatus.NOT_CHECKED
        assert "עדיין לא מיושם" in result.message or "future" in result.message.lower()

    def test_default_service_factory(self):
        """Test that create_default_service creates a no-DB fallback."""
        service = InstitutionValidationService.create_default_service()

        assert service.db_available is False
        result = service.validate_institution_code("any_code")
        assert result.status == ValidationStatus.UNKNOWN
        assert result.is_warning is True


class TestValidationIntegrationWithReport:
    """Test integration of validation results with processing reports."""

    def test_validation_warnings_can_be_added_to_report(self):
        """Test that validation results can be converted to report warnings."""
        service = InstitutionValidationService()  # No DB
        result = service.validate_institution_code("12345")

        # Simulate adding to report
        warning_messages = []
        if result.is_warning_or_unknown():
            warning_messages.append(result.message)

        assert len(warning_messages) == 1
        assert "בסיס נתונים" in warning_messages[0]

    def test_multiple_validation_failures_aggregate(self):
        """Test aggregating multiple validation failures."""
        def strict_validator(value: str) -> bool:
            return False

        service = InstitutionValidationService(
            institution_code_validator=strict_validator,
            institution_type_validator=strict_validator,
        )

        code_result = service.validate_institution_code("invalid")
        type_result = service.validate_institution_type("invalid")

        # Both should be invalid
        assert code_result.status == ValidationStatus.INVALID
        assert type_result.status == ValidationStatus.INVALID

        # Both can be aggregated for a report
        all_failures = [r for r in [code_result, type_result] if not r.is_valid()]
        assert len(all_failures) == 2
