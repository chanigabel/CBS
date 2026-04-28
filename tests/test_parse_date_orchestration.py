"""Unit tests for DateEngine parse_date orchestration method."""

import pytest
from datetime import datetime
from src.excel_standardization.engines.date_engine import DateEngine
from src.excel_standardization.data_types import DateFormatPattern, DateFieldType


class TestParseDateOrchestration:
    """Tests for parse_date orchestration method."""
    
    def test_parse_from_split_columns_when_all_present(self):
        """Should parse from split columns when all three values are present."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val="1990",
            month_val="12",
            day_val="25",
            main_val="01/01/2000",  # Should be ignored
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
        assert result.status_text == ""
    
    def test_parse_from_main_value_when_split_columns_empty(self):
        """Should parse from main value when split columns are empty."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val=None,
            month_val=None,
            day_val=None,
            main_val="25/12/1990",
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
        assert result.status_text == ""
    
    def test_parse_from_main_value_when_split_columns_have_empty_strings(self):
        """Should parse from main value when split columns are empty strings."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val="",
            month_val="",
            day_val="",
            main_val="25/12/1990",
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
        assert result.status_text == ""
    
    def test_parse_from_main_value_when_split_columns_partially_filled(self):
        """Should parse from main value when split columns are partially filled."""
        engine = DateEngine()
        
        # Only year is present, month and day are None
        result = engine.parse_date(
            year_val="1990",
            month_val=None,
            day_val=None,
            main_val="25/12/1990",
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_business_rules_applied_to_split_column_result(self):
        """Business rules should be applied to split column parsing result."""
        engine = DateEngine()
        today = datetime.now()
        future_year = today.year + 1
        
        result = engine.parse_date(
            year_val=str(future_year),
            month_val="6",
            day_val="15",
            main_val=None,
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.is_valid is False
        assert result.status_text == "תאריך לידה עתידי"
    
    def test_business_rules_applied_to_main_value_result(self):
        """Business rules should be applied to main value parsing result."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val=None,
            month_val=None,
            day_val=None,
            main_val="25/12/1850",
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.is_valid is False
        assert result.status_text == "שנה לפני 1900"
    
    def test_entry_date_empty_clears_status(self):
        """Entry date with empty main value should have cleared status."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val=None,
            month_val=None,
            day_val=None,
            main_val=None,
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.ENTRY_DATE
        )
        
        assert result.is_valid is False
        assert result.status_text == ""  # Should be cleared for entry date
    
    def test_birth_date_empty_keeps_status(self):
        """Birth date with empty main value should keep 'תא ריק' status."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val=None,
            month_val=None,
            day_val=None,
            main_val=None,
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.is_valid is False
        assert result.status_text == "תא ריק"
    
    def test_age_over_100_warning_from_split_columns(self):
        """Age over 100 warning should be applied to split column result."""
        engine = DateEngine()
        today = datetime.now()
        birth_year = today.year - 105
        
        result = engine.parse_date(
            year_val=str(birth_year),
            month_val="1",
            day_val="1",
            main_val=None,
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.is_valid is True  # Should remain valid
        assert "גיל מעל 100" in result.status_text
        assert "105 שנים" in result.status_text
    
    def test_two_digit_year_expansion_in_split_columns(self):
        """Two-digit year in split columns should be expanded."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val="90",
            month_val="12",
            day_val="25",
            main_val=None,
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_numeric_string_in_main_value(self):
        """Numeric string in main value should be parsed correctly."""
        engine = DateEngine()
        
        result = engine.parse_date(
            year_val=None,
            month_val=None,
            day_val=None,
            main_val="25121990",
            pattern=DateFormatPattern.DDMM,
            field_type=DateFieldType.BIRTH_DATE
        )
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
