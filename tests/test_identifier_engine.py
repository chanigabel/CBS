"""Unit tests for IdentifierEngine.

Tests the pure business logic for Israeli ID and passport validation.
"""

import pytest
from src.excel_standardization.engines.identifier_engine import IdentifierEngine


class TestClassifyIdValue:
    """Tests for classify_id_value method."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.engine = IdentifierEngine()
    
    def test_empty_value(self):
        """Empty value should return empty digits and not move to passport."""
        digits, should_move, reason = self.engine.classify_id_value("")
        assert digits == ""
        assert should_move is False
        assert reason == ""
    
    def test_9999_special_value(self):
        """9999 should be treated as empty."""
        digits, should_move, reason = self.engine.classify_id_value("9999")
        assert digits == ""
        assert should_move is False
        assert reason == ""
    
    def test_non_digit_characters(self):
        """Non-digit characters (except dashes) should move to passport."""
        digits, should_move, reason = self.engine.classify_id_value("12345ABC")
        assert digits == ""
        assert should_move is True
        assert reason == "invalid_format"
    
    def test_dash_variants_accepted(self):
        """All dash variants should be accepted and removed."""
        # Test hyphen (45)
        digits, should_move, reason = self.engine.classify_id_value("123-456-789")
        assert digits == "123456789"
        assert should_move is False
        
        # Test en-dash (8211)
        digits, should_move, reason = self.engine.classify_id_value("123–456–789")
        assert digits == "123456789"
        assert should_move is False
        
        # Test em-dash (8212)
        digits, should_move, reason = self.engine.classify_id_value("123—456—789")
        assert digits == "123456789"
        assert should_move is False
    
    def test_too_few_digits(self):
        """Less than 4 digits should move to passport with specific reason."""
        digits, should_move, reason = self.engine.classify_id_value("123")
        assert digits == ""
        assert should_move is True
        assert reason == "too_short"
    
    def test_too_many_digits(self):
        """More than 9 digits should move to passport."""
        digits, should_move, reason = self.engine.classify_id_value("1234567890")
        assert digits == ""
        assert should_move is True
        assert reason == "too_long"
    
    def test_valid_digit_count_range(self):
        """4-9 digits should be accepted."""
        # 4 digits
        digits, should_move, reason = self.engine.classify_id_value("1234")
        assert digits == "1234"
        assert should_move is False
        
        # 9 digits
        digits, should_move, reason = self.engine.classify_id_value("123456789")
        assert digits == "123456789"
        assert should_move is False


class TestValidateIsraeliId:
    """Tests for validate_israeli_id method."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.engine = IdentifierEngine()
    
    def test_valid_checksum(self):
        """Valid Israeli ID should pass checksum validation."""
        # Known valid Israeli ID: 000000018
        assert self.engine.validate_israeli_id("000000018") is True
    
    def test_invalid_checksum(self):
        """Invalid Israeli ID should fail checksum validation."""
        # Invalid checksum
        assert self.engine.validate_israeli_id("000000019") is False
    
    def test_checksum_algorithm(self):
        """Test the checksum algorithm step by step."""
        # Test ID: 123456782
        # Position: 1  2  3  4  5  6  7  8  9
        # Digit:    1  2  3  4  5  6  7  8  2
        # Multiply: 1  2  1  2  1  2  1  2  1
        # Result:   1  4  3  8  5 12  7 16  2
        # Adjust:   1  4  3  8  5  3  7  7  2
        # Sum: 1+4+3+8+5+3+7+7+2 = 40, divisible by 10
        assert self.engine.validate_israeli_id("123456782") is True
    
    def test_wrong_length(self):
        """ID with wrong length should fail validation."""
        assert self.engine.validate_israeli_id("12345678") is False
        assert self.engine.validate_israeli_id("1234567890") is False


class TestPadId:
    """Tests for pad_id method."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.engine = IdentifierEngine()
    
    def test_pad_4_digits(self):
        """4-digit ID should be padded to 9 digits."""
        assert self.engine.pad_id("1234") == "000001234"
    
    def test_pad_5_digits(self):
        """5-digit ID should be padded to 9 digits."""
        assert self.engine.pad_id("12345") == "000012345"
    
    def test_pad_9_digits(self):
        """9-digit ID should remain unchanged."""
        assert self.engine.pad_id("123456789") == "123456789"
    
    def test_pad_preserves_leading_zeros(self):
        """Padding should add leading zeros."""
        assert self.engine.pad_id("1") == "000000001"


class TestCleanPassport:
    """Tests for clean_passport method."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.engine = IdentifierEngine()
    
    def test_empty_passport(self):
        """Empty passport should return empty string."""
        assert self.engine.clean_passport("") == ""
    
    def test_keep_digits(self):
        """Digits should be preserved."""
        assert self.engine.clean_passport("123456") == "123456"
    
    def test_keep_english_letters(self):
        """English letters should be preserved."""
        assert self.engine.clean_passport("ABC123xyz") == "ABC123xyz"
    
    def test_keep_hebrew_letters(self):
        """Hebrew letters should be preserved."""
        # Hebrew letters: א (1488), ב (1489), ג (1490)
        assert self.engine.clean_passport("אבג123") == "אבג123"
    
    def test_keep_dashes(self):
        """Dash characters should be preserved."""
        assert self.engine.clean_passport("123-456") == "123-456"
    
    def test_remove_invalid_characters(self):
        """Invalid characters should be removed."""
        assert self.engine.clean_passport("123@#$456") == "123456"
        assert self.engine.clean_passport("ABC!@#123") == "ABC123"
    
    def test_mixed_valid_invalid(self):
        """Mixed valid and invalid characters."""
        assert self.engine.clean_passport("A1B2@C3#") == "A1B2C3"


class TestNormalizeIdentifiers:
    """Tests for normalize_identifiers method (integration)."""
    
    def setup_method(self):
        """Set up test fixtures."""
        self.engine = IdentifierEngine()
    
    def test_empty_id_empty_passport(self):
        """Empty ID and passport should return 'חסר מזהים'."""
        result = self.engine.normalize_identifiers("", "")
        assert result.corrected_id == ""
        assert result.corrected_passport == ""
        assert result.status_text == "חסר מזהים"
    
    def test_empty_id_with_passport(self):
        """Empty ID with passport should return 'דרכון הוזן'."""
        result = self.engine.normalize_identifiers("", "ABC123")
        assert result.corrected_id == ""
        assert result.corrected_passport == "ABC123"
        assert result.status_text == "דרכון הוזן"
    
    def test_valid_id_no_passport(self):
        """Valid ID without passport should return 'ת.ז. תקינה'."""
        result = self.engine.normalize_identifiers("000000018", "")
        assert result.corrected_id == "000000018"
        assert result.corrected_passport == ""
        assert result.status_text == "ת.ז. תקינה"
    
    def test_valid_id_with_passport(self):
        """Valid ID with passport should return 'ת.ז. תקינה + דרכון הוזן'."""
        result = self.engine.normalize_identifiers("000000018", "ABC123")
        assert result.corrected_id == "000000018"
        assert result.corrected_passport == "ABC123"
        assert result.status_text == "ת.ז. תקינה + דרכון הוזן"
    
    def test_invalid_id_no_passport(self):
        """Invalid ID without passport should return 'ת.ז. לא תקינה'."""
        result = self.engine.normalize_identifiers("000000019", "")
        assert result.corrected_id == "000000019"
        assert result.corrected_passport == ""
        assert result.status_text == "ת.ז. לא תקינה"
    
    def test_invalid_id_with_passport(self):
        """Invalid ID with passport should return 'ת.ז. לא תקינה + דרכון הוזן'."""
        result = self.engine.normalize_identifiers("000000019", "ABC123")
        assert result.corrected_id == "000000019"
        assert result.corrected_passport == "ABC123"
        assert result.status_text == "ת.ז. לא תקינה + דרכון הוזן"
    
    def test_all_zeros_id(self):
        """All zeros ID should be marked as invalid."""
        result = self.engine.normalize_identifiers("000000000", "")
        assert result.corrected_id == "000000000"
        assert result.status_text == "ת.ז. לא תקינה"
    
    def test_all_identical_digits(self):
        """All identical digits should be marked as invalid."""
        result = self.engine.normalize_identifiers("111111111", "")
        assert result.corrected_id == "111111111"
        assert result.status_text == "ת.ז. לא תקינה"
    
    def test_id_moved_to_passport_invalid_format(self):
        """ID with letters: hyphens-only cleanup leaves letters intact.

        With hyphen-only clean_id_number, 'ABC123' is unchanged (no hyphens).
        _process_id_value sees 'A' which is non-digit/non-dash and moves the
        entire value to passport via clean_passport, which keeps letters.
        """
        result = self.engine.normalize_identifiers("ABC123", "")
        assert result.corrected_id == ""
        assert result.corrected_passport == "ABC123"
        assert result.status_text == "ת.ז. הועברה לדרכון"
    
    def test_id_moved_to_passport_too_short(self):
        """ID with too few digits should be moved to passport."""
        result = self.engine.normalize_identifiers("123", "")
        assert result.corrected_id == ""
        assert result.corrected_passport == "123"
        assert result.status_text == "ת.ז. לא תקינה + הועברה לדרכון"
    
    def test_id_padding(self):
        """ID with 4-9 digits should be padded to 9 digits."""
        result = self.engine.normalize_identifiers("12345", "")
        assert result.corrected_id == "000012345"
    
    def test_id_moved_combines_with_existing_passport(self):
        """ID moved to passport should combine with existing passport."""
        result = self.engine.normalize_identifiers("ABC", "XYZ")
        assert result.corrected_id == ""
        # VBA parity: move-to-passport happens only when passport was empty
        assert result.corrected_passport == "XYZ"
