"""Unit tests for DateEngine core methods."""

import pytest
from datetime import datetime, date
from src.excel_standardization.engines.date_engine import DateEngine
from src.excel_standardization.data_types import DateFormatPattern


class TestExpandTwoDigitYear:
    """Tests for expand_two_digit_year method."""
    
    def test_year_less_than_or_equal_current(self):
        """Years <= current year % 100 should be in 2000s."""
        engine = DateEngine()
        current_two_digit = datetime.now().year % 100
        
        # Test with current year's two-digit
        result = engine.expand_two_digit_year(current_two_digit)
        assert result == 2000 + current_two_digit
        
        # Test with year less than current
        if current_two_digit > 0:
            result = engine.expand_two_digit_year(current_two_digit - 1)
            assert result == 2000 + current_two_digit - 1
    
    def test_year_greater_than_current(self):
        """Years > current year % 100 should be in 1900s."""
        engine = DateEngine()
        current_two_digit = datetime.now().year % 100
        
        # Test with year greater than current
        if current_two_digit < 99:
            result = engine.expand_two_digit_year(current_two_digit + 1)
            assert result == 1900 + current_two_digit + 1
    
    def test_boundary_values(self):
        """Test boundary values 0 and 99."""
        engine = DateEngine()
        
        # 0 should always be 2000
        result = engine.expand_two_digit_year(0)
        assert result == 2000
        
        # 99 depends on current year
        result = engine.expand_two_digit_year(99)
        current_two_digit = datetime.now().year % 100
        if 99 <= current_two_digit:
            assert result == 2099
        else:
            assert result == 1999


class TestParseFromSplitColumns:
    """Tests for parse_from_split_columns method."""
    
    def test_valid_date(self):
        """Valid date components should parse successfully."""
        engine = DateEngine()
        result = engine.parse_from_split_columns(1990, 12, 25)
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
        assert result.status_text == ""
    
    def test_two_digit_year_expansion(self):
        """Two-digit years should be expanded."""
        engine = DateEngine()
        result = engine.parse_from_split_columns(90, 12, 25)
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_empty_values(self):
        """Empty values should return invalid result."""
        engine = DateEngine()
        
        result = engine.parse_from_split_columns(None, 12, 25)
        assert result.is_valid is False
        
        result = engine.parse_from_split_columns(1990, None, 25)
        assert result.is_valid is False
        
        result = engine.parse_from_split_columns(1990, 12, None)
        assert result.is_valid is False
    
    def test_invalid_day_range(self):
        """Day outside 1-31 should be invalid."""
        engine = DateEngine()
        
        result = engine.parse_from_split_columns(1990, 12, 0)
        assert result.is_valid is False
        assert "יום לא תקין" in result.status_text
        
        result = engine.parse_from_split_columns(1990, 12, 32)
        assert result.is_valid is False
        assert "יום לא תקין" in result.status_text
    
    def test_invalid_month_range(self):
        """Month outside 1-12 should be invalid."""
        engine = DateEngine()
        
        result = engine.parse_from_split_columns(1990, 0, 25)
        assert result.is_valid is False
        assert "חודש לא תקין" in result.status_text
        
        result = engine.parse_from_split_columns(1990, 13, 25)
        assert result.is_valid is False
        assert "חודש לא תקין" in result.status_text
    
    def test_nonexistent_date(self):
        """Dates that don't exist should be invalid."""
        engine = DateEngine()
        
        # February 30 doesn't exist
        result = engine.parse_from_split_columns(1990, 2, 30)
        assert result.is_valid is False
        assert "תאריך לא קיים" in result.status_text
        
        # February 29 in non-leap year
        result = engine.parse_from_split_columns(1990, 2, 29)
        assert result.is_valid is False
        assert "תאריך לא קיים" in result.status_text
    
    def test_leap_year_date(self):
        """February 29 in leap year should be valid."""
        engine = DateEngine()
        
        result = engine.parse_from_split_columns(2000, 2, 29)
        assert result.is_valid is True
        assert result.year == 2000
        assert result.month == 2
        assert result.day == 29


class TestParseNumericDateString:
    """Tests for parse_numeric_date_string method."""
    
    def test_8_digit_format(self):
        """8-digit DDMMYYYY format should parse correctly."""
        engine = DateEngine()
        
        result = engine.parse_numeric_date_string("25121990")
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_6_digit_format(self):
        """6-digit DDMMYY format should parse correctly."""
        engine = DateEngine()
        
        result = engine.parse_numeric_date_string("251290")
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_4_digit_format(self):
        """4-digit DMYY format should parse correctly."""
        engine = DateEngine()
        
        result = engine.parse_numeric_date_string("5190")
        assert result.year == 1990
        assert result.month == 1
        assert result.day == 5
        assert result.is_valid is True
    
    def test_invalid_length(self):
        """Invalid length should return error."""
        engine = DateEngine()
        
        result = engine.parse_numeric_date_string("123")
        assert result.is_valid is False
        assert "אורך תאריך לא תקין" in result.status_text
    
    def test_non_numeric(self):
        """Non-numeric strings should return error."""
        engine = DateEngine()
        
        result = engine.parse_numeric_date_string("25/12/1990")
        assert result.is_valid is False
        assert "פורמט תאריך לא תקין" in result.status_text


class TestParseSeparatedDateString:
    """Tests for parse_separated_date_string method."""
    
    def test_ddmm_with_slash(self):
        """DDMM format with / separator should parse correctly."""
        engine = DateEngine()
        
        result = engine.parse_separated_date_string("25/12/1990", DateFormatPattern.DDMM)
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_ddmm_with_dot(self):
        """DDMM format with . separator should parse correctly."""
        engine = DateEngine()
        
        result = engine.parse_separated_date_string("25.12.1990", DateFormatPattern.DDMM)
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_mmdd_with_slash(self):
        """MMDD format with / separator should parse correctly."""
        engine = DateEngine()
        
        result = engine.parse_separated_date_string("12/25/1990", DateFormatPattern.MMDD)
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_two_part_date_uses_current_year(self):
        """Two-part date should use current year."""
        engine = DateEngine()
        
        result = engine.parse_separated_date_string("25/12", DateFormatPattern.DDMM)
        assert result.year == datetime.now().year
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_no_separator(self):
        """String without separator should return error."""
        engine = DateEngine()
        
        result = engine.parse_separated_date_string("25121990", DateFormatPattern.DDMM)
        assert result.is_valid is False
        assert "אין מפריד בתאריך" in result.status_text
    
    def test_single_part(self):
        """Single part should return error."""
        engine = DateEngine()
        
        result = engine.parse_separated_date_string("25", DateFormatPattern.DDMM)
        assert result.is_valid is False
        # Single part without separator returns "no separator" error
        assert "אין מפריד בתאריך" in result.status_text


class TestParseFromMainValue:
    """Tests for parse_from_main_value method."""
    
    def test_empty_value(self):
        """Empty value should return 'תא ריק'."""
        engine = DateEngine()
        
        result = engine.parse_from_main_value(None, DateFormatPattern.DDMM)
        assert result.is_valid is False
        assert result.status_text == "תא ריק"
        
        result = engine.parse_from_main_value("", DateFormatPattern.DDMM)
        assert result.is_valid is False
        assert result.status_text == "תא ריק"
    
    def test_datetime_object(self):
        """datetime object should extract components directly."""
        engine = DateEngine()
        
        dt = datetime(1990, 12, 25, 10, 30, 0)
        result = engine.parse_from_main_value(dt, DateFormatPattern.DDMM)
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
        assert result.status_text == ""
    
    def test_date_object(self):
        """date object should extract components directly."""
        engine = DateEngine()
        
        d = date(1990, 12, 25)
        result = engine.parse_from_main_value(d, DateFormatPattern.DDMM)
        
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
        assert result.status_text == ""
    
    def test_numeric_string(self):
        """Numeric string should use parse_numeric_date_string."""
        engine = DateEngine()
        
        result = engine.parse_from_main_value("25121990", DateFormatPattern.DDMM)
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_separated_string(self):
        """Separated string should use parse_separated_date_string."""
        engine = DateEngine()
        
        result = engine.parse_from_main_value("25/12/1990", DateFormatPattern.DDMM)
        assert result.year == 1990
        assert result.month == 12
        assert result.day == 25
        assert result.is_valid is True
    
    def test_unknown_format(self):
        """Unknown format should return error."""
        engine = DateEngine()
        
        result = engine.parse_from_main_value("invalid-date", DateFormatPattern.DDMM)
        assert result.is_valid is False
        assert "פורמט תאריך לא מזוהה" in result.status_text

    def test_iso_datetime_string(self):
        """ISO datetime string like '1997-09-04T00:00:00' should parse correctly."""
        engine = DateEngine()
        result = engine.parse_from_main_value("1997-09-04T00:00:00", DateFormatPattern.DDMM)
        assert result.is_valid is True
        assert result.year == 1997
        assert result.month == 9
        assert result.day == 4

    def test_iso_datetime_string_future(self):
        """ISO datetime string with future date should parse (business rules applied separately)."""
        engine = DateEngine()
        result = engine.parse_from_main_value("2025-09-20T00:00:00", DateFormatPattern.DDMM)
        # parse_from_main_value itself just parses — business rules are applied by parse_date
        assert result.year == 2025
        assert result.month == 9
        assert result.day == 20


if __name__ == "__main__":
    pytest.main([__file__, "-v"])



class TestCalculateAge:
    """Tests for calculate_age method."""
    
    def test_age_birthday_already_passed(self):
        """Age calculation when birthday already passed this year."""
        engine = DateEngine()
        today = datetime.now()
        
        # Birth date 3 months ago (or January if we're in Jan-Mar)
        birth_year = today.year - 30
        if today.month > 3:
            birth_month = today.month - 3
        else:
            # If we're in Jan-Mar, use January to ensure birthday passed
            birth_month = 1
        birth_day = min(today.day, 28)  # Use day 28 to avoid month-end issues
        
        age = engine.calculate_age(birth_year, birth_month, birth_day)
        assert age == 30
    
    def test_age_birthday_not_yet_this_year(self):
        """Age calculation when birthday hasn't occurred yet this year."""
        engine = DateEngine()
        today = datetime.now()
        
        # Birth date 3 months from now (or December if we're in Oct-Dec)
        birth_year = today.year - 30
        if today.month <= 9:
            birth_month = today.month + 3
        else:
            # If we're in Oct-Dec, use December to ensure birthday hasn't passed
            birth_month = 12
        birth_day = min(today.day, 28)  # Use day 28 to avoid month-end issues
        
        age = engine.calculate_age(birth_year, birth_month, birth_day)
        assert age == 29
    
    def test_age_today_is_birthday(self):
        """Age calculation when today is the birthday."""
        engine = DateEngine()
        today = datetime.now()
        
        birth_year = today.year - 25
        birth_month = today.month
        birth_day = today.day
        
        age = engine.calculate_age(birth_year, birth_month, birth_day)
        assert age == 25
    
    def test_age_over_100(self):
        """Age calculation for someone over 100 years old."""
        engine = DateEngine()
        today = datetime.now()
        
        birth_year = today.year - 105
        birth_month = today.month
        birth_day = today.day
        
        age = engine.calculate_age(birth_year, birth_month, birth_day)
        assert age == 105


class TestValidateBusinessRules:
    """Tests for validate_business_rules method."""
    
    def test_entry_date_empty_clears_status(self):
        """Entry date with 'תא ריק' status should be cleared."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        result = DateParseResult(
            year=None,
            month=None,
            day=None,
            is_valid=False,
            status_text="תא ריק"
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.ENTRY_DATE)
        assert validated.status_text == ""
        assert validated.is_valid is False
    
    def test_birth_date_empty_keeps_status(self):
        """Birth date with 'תא ריק' status should keep the status."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        result = DateParseResult(
            year=None,
            month=None,
            day=None,
            is_valid=False,
            status_text="תא ריק"
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert validated.status_text == "תא ריק"
        assert validated.is_valid is False
    
    def test_invalid_date_keeps_status(self):
        """Invalid date should keep existing status."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        result = DateParseResult(
            year=1990,
            month=13,
            day=25,
            is_valid=False,
            status_text="חודש לא תקין: 13"
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert validated.status_text == "חודש לא תקין: 13"
        assert validated.is_valid is False
    
    def test_year_before_1900(self):
        """Year before 1900 should be marked invalid."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        result = DateParseResult(
            year=1850,
            month=12,
            day=25,
            is_valid=True,
            status_text=""
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert validated.status_text == "שנה לפני 1900"
        assert validated.is_valid is False
    
    def test_future_birth_date(self):
        """Future birth date should be marked invalid."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        today = datetime.now()
        future_year = today.year + 1
        
        result = DateParseResult(
            year=future_year,
            month=6,
            day=15,
            is_valid=True,
            status_text=""
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert validated.status_text == "תאריך לידה עתידי"
        assert validated.is_valid is False
    
    def test_future_entry_date(self):
        """Future entry date should be marked invalid with the cutoff-year status."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        today = datetime.now()
        future_year = today.year + 1
        
        result = DateParseResult(
            year=future_year,
            month=6,
            day=15,
            is_valid=True,
            status_text=""
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.ENTRY_DATE)
        # The cutoff rule (year >= current_year) fires before the generic future-date
        # check, so future entry dates receive the cutoff status, not "תאריך כניסה עתידי".
        assert validated.status_text == "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is False
    
    def test_age_over_100_warning(self):
        """Birth date with age > 100 should show warning but remain valid."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        today = datetime.now()
        birth_year = today.year - 105
        
        # Use a date that has already passed this year to ensure age is 105
        result = DateParseResult(
            year=birth_year,
            month=1,  # January to ensure birthday has passed
            day=1,
            is_valid=True,
            status_text=""
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert "גיל מעל 100" in validated.status_text
        assert "105 שנים" in validated.status_text
        assert validated.is_valid is True  # Should remain valid
    
    def test_age_exactly_100_no_warning(self):
        """Birth date with age exactly 100 should not show warning."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        today = datetime.now()
        birth_year = today.year - 100
        
        result = DateParseResult(
            year=birth_year,
            month=today.month,
            day=today.day,
            is_valid=True,
            status_text=""
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert validated.status_text == ""
        assert validated.is_valid is True
    
    def test_valid_date_passes_all_rules(self):
        """Valid date within acceptable range should pass all rules."""
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        
        engine = DateEngine()
        
        result = DateParseResult(
            year=1990,
            month=12,
            day=25,
            is_valid=True,
            status_text=""
        )
        
        validated = engine.validate_business_rules(result, DateFieldType.BIRTH_DATE)
        assert validated.status_text == ""
        assert validated.is_valid is True
        assert validated.year == 1990
        assert validated.month == 12
        assert validated.day == 25


# ---------------------------------------------------------------------------
# New base completion algorithm (current-year-relative threshold)
# ---------------------------------------------------------------------------

class TestExpandTwoDigitYearNewAlgorithm:
    """Verify the current-year-relative threshold algorithm.

    With current year 2026 (threshold = 26):
      30 > 26  → 1930
      27 > 26  → 1927
      26 <= 26 → 2026
      12 <= 26 → 2012
       5 <= 26 → 2005
       7 <= 26 → 2007
    """

    def setup_method(self):
        self.engine = DateEngine()
        self.threshold = date.today().year % 100
        self.century_2000 = (date.today().year // 100) * 100
        self.century_1900 = self.century_2000 - 100

    def test_year_greater_than_threshold_goes_to_1900s(self):
        # Any 2-digit year strictly greater than threshold → 1900s
        yr = self.threshold + 1
        if yr > 99:
            return  # skip if threshold is 99
        result = self.engine.expand_two_digit_year(yr)
        assert result == self.century_1900 + yr

    def test_year_equal_to_threshold_goes_to_2000s(self):
        # Equal to threshold → 2000s
        yr = self.threshold
        result = self.engine.expand_two_digit_year(yr)
        assert result == self.century_2000 + yr

    def test_year_less_than_threshold_goes_to_2000s(self):
        # Less than threshold → 2000s
        if self.threshold == 0:
            return  # nothing less than 0 in valid range
        yr = max(0, self.threshold - 5)
        result = self.engine.expand_two_digit_year(yr)
        assert result == self.century_2000 + yr

    def test_1_digit_year_goes_to_2000s(self):
        # Single-digit year (e.g. 7) is always <= any reasonable threshold
        result = self.engine.expand_two_digit_year(7)
        assert result == self.century_2000 + 7

    def test_zero_goes_to_2000(self):
        result = self.engine.expand_two_digit_year(0)
        assert result == self.century_2000

    def test_99_goes_to_1900s(self):
        # 99 is always > any current threshold (threshold < 99 for decades to come)
        result = self.engine.expand_two_digit_year(99)
        assert result == self.century_1900 + 99


class TestParseFromSplitColumnsAutoCompletedFlag:
    """Verify year_was_auto_completed is set correctly."""

    def setup_method(self):
        self.engine = DateEngine()

    def test_4_digit_year_not_auto_completed(self):
        result = self.engine.parse_from_split_columns(1990, 6, 15)
        assert result.year_was_auto_completed is False

    def test_2_digit_year_is_auto_completed(self):
        result = self.engine.parse_from_split_columns(30, 6, 15)
        assert result.year_was_auto_completed is True

    def test_1_digit_year_is_auto_completed(self):
        result = self.engine.parse_from_split_columns(7, 6, 15)
        assert result.year_was_auto_completed is True

    def test_explicit_2026_not_auto_completed(self):
        result = self.engine.parse_from_split_columns(2026, 1, 1)
        assert result.year_was_auto_completed is False

    def test_2_digit_year_expands_correctly(self):
        # 30 > threshold (26 in 2026) → 1930
        result = self.engine.parse_from_split_columns(30, 6, 15)
        assert result.year == 1930
        assert result.year_was_auto_completed is True

    def test_threshold_year_expands_to_2000s(self):
        # threshold (e.g. 26) → 2026
        threshold = date.today().year % 100
        century_2000 = (date.today().year // 100) * 100
        result = self.engine.parse_from_split_columns(threshold, 6, 15)
        assert result.year == century_2000 + threshold
        assert result.year_was_auto_completed is True


# ---------------------------------------------------------------------------
# Entry-date cutoff-year validation (current_year - 1 is the latest allowed)
# ---------------------------------------------------------------------------

class TestEntryDateCutoffYear:
    """Verify that entry_date.year >= current_year triggers the late-date status.

    The allowed cutoff is current_year - 1.  With current year 2026:
      2025-01-01  → valid (year <= 2025)
      2025-12-31  → valid (year <= 2025)
      2026-01-01  → invalid: "תאריך כניסה מאוחר מהתאריך שנקבע"
      2027-05-10  → invalid: "תאריך כניסה מאוחר מהתאריך שנקבע"
    """

    def setup_method(self):
        from src.excel_standardization.engines.date_engine import DateEngine
        from src.excel_standardization.data_types import DateParseResult, DateFieldType
        self.engine = DateEngine()
        self.DateParseResult = DateParseResult
        self.DateFieldType = DateFieldType
        self.current_year = date.today().year
        self.cutoff_year = self.current_year - 1  # latest allowed year

    def _make_entry_result(self, year, month=1, day=1):
        return self.DateParseResult(
            year=year, month=month, day=day,
            is_valid=True, status_text=""
        )

    # --- valid dates (year <= cutoff_year) ---

    def test_cutoff_year_jan1_is_valid(self):
        """First day of the cutoff year must be valid."""
        result = self._make_entry_result(self.cutoff_year, 1, 1)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text != "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is True

    def test_cutoff_year_dec31_is_valid(self):
        """Last day of the cutoff year must be valid."""
        result = self._make_entry_result(self.cutoff_year, 12, 31)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text != "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is True

    def test_well_past_date_is_valid(self):
        """A date many years in the past must be valid."""
        result = self._make_entry_result(2000, 6, 15)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text != "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is True

    # --- invalid dates (year >= current_year) ---

    def test_current_year_jan1_is_late(self):
        """First day of the current year must be flagged as late."""
        result = self._make_entry_result(self.current_year, 1, 1)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text == "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is False

    def test_current_year_mid_year_is_late(self):
        """A mid-year date in the current year must be flagged as late."""
        result = self._make_entry_result(self.current_year, 6, 15)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text == "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is False

    def test_future_year_is_late(self):
        """A date two years in the future must be flagged as late."""
        result = self._make_entry_result(self.current_year + 1, 5, 10)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text == "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert validated.is_valid is False

    # --- split-column path ---

    def test_split_columns_cutoff_year_valid(self):
        """Split-column parse of a cutoff-year date must be valid."""
        result = self.engine.parse_date(
            str(self.cutoff_year), "6", "15", None, None,
            self.DateFieldType.ENTRY_DATE
        )
        assert result.status_text != "תאריך כניסה מאוחר מהתאריך שנקבע"

    def test_split_columns_current_year_late(self):
        """Split-column parse of a current-year date must be flagged as late."""
        result = self.engine.parse_date(
            str(self.current_year), "1", "1", None, None,
            self.DateFieldType.ENTRY_DATE
        )
        assert result.status_text == "תאריך כניסה מאוחר מהתאריך שנקבע"
        assert result.is_valid is False

    # --- birth-date field is NOT affected ---

    def test_birth_date_current_year_not_flagged_as_late(self):
        """The cutoff rule must NOT apply to birth dates."""
        # A birth date in the current year that is not in the future should
        # not receive the entry-date late status.
        today = date.today()
        # Use Jan 1 of current year — it's in the past, so not "future birth"
        result = self._make_entry_result(self.current_year, 1, 1)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.BIRTH_DATE)
        assert validated.status_text != "תאריך כניסה מאוחר מהתאריך שנקבע"

    # --- interaction: late status is not silently overwritten by other checks ---

    def test_late_status_is_set_not_empty(self):
        """The late status must be a non-empty string."""
        result = self._make_entry_result(self.current_year, 3, 1)
        validated = self.engine.validate_business_rules(result, self.DateFieldType.ENTRY_DATE)
        assert validated.status_text.strip() != ""
