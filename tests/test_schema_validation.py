"""Tests for JSON schema validation utilities.

This module tests the schema validation functions and field naming convention utilities.

Requirements:
    - Validates: Requirements 19.1-19.5
"""

import pytest
from src.excel_standardization.schema_validation import (
    validate_json_row,
    validate_sheet_dataset_schema,
    validate_workbook_dataset_schema,
    validate_field_naming_convention,
    get_corrected_field_name,
    get_original_field_name,
    is_corrected_field,
    get_field_pairs,
    is_jsonschema_available,
    load_schema
)
from src.excel_standardization.data_types import SheetDataset, WorkbookDataset, JsonRow


# ============================================================================
# Field Naming Convention Tests
# ============================================================================

def test_get_corrected_field_name():
    """Test getting corrected field name from original field name."""
    assert get_corrected_field_name("first_name") == "first_name_corrected"
    assert get_corrected_field_name("gender") == "gender_corrected"
    assert get_corrected_field_name("birth_year") == "birth_year_corrected"
    
    # Already corrected field should return itself
    assert get_corrected_field_name("first_name_corrected") == "first_name_corrected"


def test_get_original_field_name():
    """Test getting original field name from corrected field name."""
    assert get_original_field_name("first_name_corrected") == "first_name"
    assert get_original_field_name("gender_corrected") == "gender"
    assert get_original_field_name("birth_year_corrected") == "birth_year"
    
    # Original field should return itself
    assert get_original_field_name("first_name") == "first_name"


def test_is_corrected_field():
    """Test checking if a field name is a corrected field."""
    assert is_corrected_field("first_name_corrected") is True
    assert is_corrected_field("gender_corrected") is True
    assert is_corrected_field("first_name") is False
    assert is_corrected_field("gender") is False


def test_get_field_pairs():
    """Test getting field pairs from a JsonRow."""
    row = {
        "first_name": "John",
        "first_name_corrected": "John",
        "gender": "M",
        "gender_corrected": "1"
    }
    
    pairs = get_field_pairs(row)
    assert len(pairs) == 2
    assert ("first_name", "first_name_corrected") in pairs
    assert ("gender", "gender_corrected") in pairs


def test_get_field_pairs_missing_corrected():
    """Test getting field pairs when corrected field is missing."""
    row = {
        "first_name": "John",
        "gender": "M",
        "gender_corrected": "1"
    }
    
    pairs = get_field_pairs(row)
    assert len(pairs) == 1
    assert ("gender", "gender_corrected") in pairs
    assert ("first_name", "first_name_corrected") not in pairs


def test_validate_field_naming_convention_valid():
    """Test validating a JsonRow with correct naming convention."""
    row = {
        "first_name": "John",
        "first_name_corrected": "John",
        "last_name": "Doe",
        "last_name_corrected": "Doe"
    }
    
    is_valid, errors = validate_field_naming_convention(row)
    assert is_valid
    assert errors is None


def test_validate_field_naming_convention_missing_corrected():
    """Test validating a JsonRow missing corrected field."""
    row = {
        "first_name": "John",
        "last_name": "Doe",
        "last_name_corrected": "Doe"
    }
    
    is_valid, errors = validate_field_naming_convention(row)
    assert not is_valid
    assert errors is not None
    assert any("first_name" in err and "missing" in err for err in errors)


def test_validate_field_naming_convention_orphaned_corrected():
    """Test validating a JsonRow with orphaned corrected field."""
    row = {
        "first_name": "John",
        "first_name_corrected": "John",
        "last_name_corrected": "Doe"  # No original last_name
    }
    
    is_valid, errors = validate_field_naming_convention(row)
    assert not is_valid
    assert errors is not None
    assert any("last_name_corrected" in err and "no corresponding" in err for err in errors)


# ============================================================================
# Schema Loading Tests
# ============================================================================

@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_load_schema_json_row():
    """Test loading JsonRow schema."""
    schema = load_schema("json_row.schema.json")
    assert schema is not None
    assert schema["$schema"] == "http://json-schema.org/draft-07/schema#"
    assert schema["title"] == "JsonRow"
    assert schema["type"] == "object"


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_load_schema_sheet_dataset():
    """Test loading SheetDataset schema."""
    schema = load_schema("sheet_dataset.schema.json")
    assert schema is not None
    assert schema["$schema"] == "http://json-schema.org/draft-07/schema#"
    assert schema["title"] == "SheetDataset"
    assert schema["type"] == "object"
    assert "sheet_name" in schema["properties"]
    assert "header_row" in schema["properties"]
    assert "rows" in schema["properties"]


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_load_schema_workbook_dataset():
    """Test loading WorkbookDataset schema."""
    schema = load_schema("workbook_dataset.schema.json")
    assert schema is not None
    assert schema["$schema"] == "http://json-schema.org/draft-07/schema#"
    assert schema["title"] == "WorkbookDataset"
    assert schema["type"] == "object"
    assert "source_file" in schema["properties"]
    assert "sheets" in schema["properties"]


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_load_schema_not_found():
    """Test loading non-existent schema."""
    with pytest.raises(FileNotFoundError):
        load_schema("nonexistent.schema.json")


# ============================================================================
# JsonRow Validation Tests
# ============================================================================

@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_json_row_valid():
    """Test validating a valid JsonRow."""
    row = {
        "first_name": "John",
        "first_name_corrected": "John",
        "last_name": "Doe",
        "last_name_corrected": "Doe",
        "gender": "M",
        "gender_corrected": "1"
    }
    
    is_valid, errors = validate_json_row(row)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_json_row_with_null_values():
    """Test validating a JsonRow with null values."""
    row = {
        "first_name": None,
        "first_name_corrected": None,
        "last_name": "Doe",
        "last_name_corrected": "Doe"
    }
    
    is_valid, errors = validate_json_row(row)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_json_row_with_split_dates():
    """Test validating a JsonRow with split date fields."""
    row = {
        "first_name": "John",
        "first_name_corrected": "John",
        "birth_year": 1980,
        "birth_year_corrected": 1980,
        "birth_month": 5,
        "birth_month_corrected": 5,
        "birth_day": 15,
        "birth_day_corrected": 15
    }
    
    is_valid, errors = validate_json_row(row)
    assert is_valid
    assert errors is None


# ============================================================================
# SheetDataset Validation Tests
# ============================================================================

@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_sheet_dataset_valid():
    """Test validating a valid SheetDataset."""
    dataset = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name"],
        rows=[
            {
                "first_name": "John",
                "first_name_corrected": "John",
                "last_name": "Doe",
                "last_name_corrected": "Doe"
            }
        ],
        metadata={}
    )
    
    is_valid, errors = validate_sheet_dataset_schema(dataset)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_sheet_dataset_with_metadata():
    """Test validating a SheetDataset with metadata."""
    dataset = SheetDataset(
        sheet_name="Students",
        header_row=2,
        header_rows_count=1,
        field_names=["first_name"],
        rows=[
            {
                "first_name": "John",
                "first_name_corrected": "John"
            }
        ],
        metadata={
            "source_file": "data.xlsx",
            "extraction_date": "2024-01-15",
            "total_rows": 1
        }
    )
    
    is_valid, errors = validate_sheet_dataset_schema(dataset)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_sheet_dataset_invalid_header_rows_count():
    """Test validating a SheetDataset with invalid header_rows_count."""
    dataset = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=3,  # Invalid: must be 1 or 2
        field_names=["first_name"],
        rows=[],
        metadata={}
    )
    
    is_valid, errors = validate_sheet_dataset_schema(dataset)
    assert not is_valid
    assert errors is not None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_sheet_dataset_empty_field_names():
    """Test validating a SheetDataset with empty field_names."""
    dataset = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=1,
        field_names=[],  # Invalid: must have at least one field
        rows=[],
        metadata={}
    )
    
    is_valid, errors = validate_sheet_dataset_schema(dataset)
    assert not is_valid
    assert errors is not None


# ============================================================================
# WorkbookDataset Validation Tests
# ============================================================================

@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_workbook_dataset_valid():
    """Test validating a valid WorkbookDataset."""
    sheet = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name"],
        rows=[
            {
                "first_name": "John",
                "first_name_corrected": "John"
            }
        ],
        metadata={}
    )
    
    dataset = WorkbookDataset(
        source_file="data.xlsx",
        sheets=[sheet],
        metadata={}
    )
    
    is_valid, errors = validate_workbook_dataset_schema(dataset)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_workbook_dataset_empty_sheets():
    """Test validating a WorkbookDataset with no sheets."""
    dataset = WorkbookDataset(
        source_file="data.xlsx",
        sheets=[],
        metadata={}
    )
    
    is_valid, errors = validate_workbook_dataset_schema(dataset)
    assert is_valid  # Empty sheets is valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_workbook_dataset_with_metadata():
    """Test validating a WorkbookDataset with metadata."""
    dataset = WorkbookDataset(
        source_file="data.xlsx",
        sheets=[],
        metadata={
            "extraction_date": "2024-01-15",
            "total_sheets": 2,
            "processed_sheets": 0,
            "skipped_sheets": ["Summary", "Notes"]
        }
    )
    
    is_valid, errors = validate_workbook_dataset_schema(dataset)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_workbook_dataset_invalid_sheet():
    """Test validating a WorkbookDataset with invalid sheet."""
    # Create a sheet with invalid header_rows_count
    sheet = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=5,  # Invalid
        field_names=["first_name"],
        rows=[],
        metadata={}
    )
    
    dataset = WorkbookDataset(
        source_file="data.xlsx",
        sheets=[sheet],
        metadata={}
    )
    
    is_valid, errors = validate_workbook_dataset_schema(dataset)
    assert not is_valid
    assert errors is not None


# ============================================================================
# Raise on Error Tests
# ============================================================================

@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_sheet_dataset_raise_on_error():
    """Test that raise_on_error parameter works for SheetDataset."""
    dataset = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=3,  # Invalid
        field_names=["first_name"],
        rows=[],
        metadata={}
    )
    
    # Should raise ValidationError
    with pytest.raises(Exception):  # ValidationError from jsonschema
        validate_sheet_dataset_schema(dataset, raise_on_error=True)


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_workbook_dataset_raise_on_error():
    """Test that raise_on_error parameter works for WorkbookDataset."""
    dataset = WorkbookDataset(
        source_file="",  # Invalid: empty string
        sheets=[],
        metadata={}
    )
    
    # Should raise ValidationError
    with pytest.raises(Exception):  # ValidationError from jsonschema
        validate_workbook_dataset_schema(dataset, raise_on_error=True)


# ============================================================================
# Integration Tests
# ============================================================================

@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_complete_dataset():
    """Test validating a complete dataset with multiple rows and fields."""
    dataset = SheetDataset(
        sheet_name="Students",
        header_row=2,
        header_rows_count=1,
        field_names=["first_name", "last_name", "gender", "birth_year"],
        rows=[
            {
                "first_name": "יוסי",
                "first_name_corrected": "יוסי",
                "last_name": "כהן",
                "last_name_corrected": "כהן",
                "gender": "ז",
                "gender_corrected": "2",
                "birth_year": 1980,
                "birth_year_corrected": 1980
            },
            {
                "first_name": "שרה",
                "first_name_corrected": "שרה",
                "last_name": "לוי",
                "last_name_corrected": "לוי",
                "gender": "נ",
                "gender_corrected": "1",
                "birth_year": 1985,
                "birth_year_corrected": 1985
            }
        ],
        metadata={
            "source_file": "data.xlsx",
            "extraction_date": "2024-01-15",
            "total_rows": 2,
            "date_field_structure": {
                "birth_date": "split"
            }
        }
    )
    
    is_valid, errors = validate_sheet_dataset_schema(dataset)
    assert is_valid
    assert errors is None


@pytest.mark.skipif(not is_jsonschema_available(), reason="jsonschema not installed")
def test_validate_workbook_with_multiple_sheets():
    """Test validating a WorkbookDataset with multiple sheets."""
    sheet1 = SheetDataset(
        sheet_name="Students",
        header_row=1,
        header_rows_count=1,
        field_names=["first_name"],
        rows=[
            {
                "first_name": "John",
                "first_name_corrected": "John"
            }
        ],
        metadata={}
    )
    
    sheet2 = SheetDataset(
        sheet_name="Teachers",
        header_row=2,
        header_rows_count=2,
        field_names=["first_name", "last_name"],
        rows=[
            {
                "first_name": "Jane",
                "first_name_corrected": "Jane",
                "last_name": "Smith",
                "last_name_corrected": "Smith"
            }
        ],
        metadata={}
    )
    
    dataset = WorkbookDataset(
        source_file="data.xlsx",
        sheets=[sheet1, sheet2],
        metadata={
            "extraction_date": "2024-01-15",
            "total_sheets": 2,
            "processed_sheets": 2,
            "skipped_sheets": []
        }
    )
    
    is_valid, errors = validate_workbook_dataset_schema(dataset)
    assert is_valid
    assert errors is None
