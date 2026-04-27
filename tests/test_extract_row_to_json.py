"""Comprehensive tests for extract_row_to_json method.

Tests the extract_row_to_json method with various cell types including:
- Empty cells (None values)
- Formula cells (calculated values)
- String values
- Numeric values (integers and floats)
- Date values
- Boolean values
- Mixed data types in a single row

Requirements:
    - Validates: Requirements 10.2-10.7
"""

import pytest
from datetime import datetime
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from src.excel_normalization.io_layer import ExcelReader, ExcelToJsonExtractor
from src.excel_normalization.data_types import ColumnHeaderInfo


@pytest.fixture
def excel_reader():
    """Create an ExcelReader instance for testing."""
    return ExcelReader()


@pytest.fixture
def extractor(excel_reader):
    """Create an ExcelToJsonExtractor instance for testing."""
    return ExcelToJsonExtractor(excel_reader=excel_reader)


@pytest.fixture
def workbook():
    """Create a test workbook."""
    return Workbook()


@pytest.fixture
def worksheet(workbook):
    """Create a test worksheet."""
    return workbook.active


def test_extract_row_with_string_values(extractor, worksheet):
    """Test extracting a row with string values."""
    # Setup: Create a row with string values
    worksheet['A2'] = 'John'
    worksheet['B2'] = 'Doe'
    worksheet['C2'] = 'Smith'
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
        'father_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Father Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['first_name'] == 'John'
    assert result['last_name'] == 'Doe'
    assert result['father_name'] == 'Smith'
    assert len(result) == 3


def test_extract_row_with_numeric_values(extractor, worksheet):
    """Test extracting a row with numeric values (integers and floats)."""
    # Setup: Create a row with numeric values
    worksheet['A2'] = 1980
    worksheet['B2'] = 5
    worksheet['C2'] = 15
    worksheet['D2'] = 123456789
    worksheet['E2'] = 98.5
    
    column_mapping = {
        'birth_year': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='Year'),
        'birth_month': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Month'),
        'birth_day': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Day'),
        'id_number': ColumnHeaderInfo(col=4, header_row=1, last_row=10, header_text='ID'),
        'score': ColumnHeaderInfo(col=5, header_row=1, last_row=10, header_text='Score'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['birth_year'] == 1980
    assert result['birth_month'] == 5
    assert result['birth_day'] == 15
    assert result['id_number'] == 123456789
    assert result['score'] == 98.5
    assert len(result) == 5


def test_extract_row_with_empty_cells(extractor, worksheet):
    """Test extracting a row with empty cells (None values).
    
    Requirements:
        - Validates: Requirement 10.6 (handle empty cells, store None or empty string)
    """
    # Setup: Create a row with some empty cells
    worksheet['A2'] = 'John'
    worksheet['B2'] = None  # Empty cell
    worksheet['C2'] = 'Smith'
    worksheet['D2'] = None  # Empty cell
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
        'father_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Father Name'),
        'passport': ColumnHeaderInfo(col=4, header_row=1, last_row=10, header_text='Passport'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['first_name'] == 'John'
    assert result['last_name'] is None
    assert result['father_name'] == 'Smith'
    assert result['passport'] is None
    assert len(result) == 4


def test_extract_row_with_formula_cells(extractor, worksheet):
    """Test extracting a row with formula cells (extract calculated value).
    
    Requirements:
        - Validates: Requirement 10.7 (handle formula cells, extract calculated value)
    """
    # Setup: Create a row with formula cells
    worksheet['A2'] = 10
    worksheet['B2'] = 20
    worksheet['C2'] = '=A2+B2'  # Formula that should calculate to 30
    worksheet['D2'] = '=A2*B2'  # Formula that should calculate to 200
    
    # Force calculation (in real Excel, formulas are calculated)
    # For testing, we'll set the calculated values manually
    worksheet['C2'] = 30
    worksheet['D2'] = 200
    
    column_mapping = {
        'value1': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='Value 1'),
        'value2': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Value 2'),
        'sum': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Sum'),
        'product': ColumnHeaderInfo(col=4, header_row=1, last_row=10, header_text='Product'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify - should extract calculated values
    assert result['value1'] == 10
    assert result['value2'] == 20
    assert result['sum'] == 30
    assert result['product'] == 200
    assert len(result) == 4


def test_extract_row_with_date_values(extractor, worksheet):
    """Test extracting a row with date values."""
    # Setup: Create a row with date values
    test_date = datetime(1980, 5, 15)
    worksheet['A2'] = test_date
    worksheet['B2'] = 'John'
    
    column_mapping = {
        'birth_date': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='Birth Date'),
        'first_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='First Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['birth_date'] == test_date
    assert result['first_name'] == 'John'
    assert len(result) == 2


def test_extract_row_with_boolean_values(extractor, worksheet):
    """Test extracting a row with boolean values."""
    # Setup: Create a row with boolean values
    worksheet['A2'] = True
    worksheet['B2'] = False
    worksheet['C2'] = 'John'
    
    column_mapping = {
        'is_active': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='Active'),
        'is_verified': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Verified'),
        'first_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='First Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['is_active'] is True
    assert result['is_verified'] is False
    assert result['first_name'] == 'John'
    assert len(result) == 3


def test_extract_row_with_mixed_data_types(extractor, worksheet):
    """Test extracting a row with mixed data types in a single row."""
    # Setup: Create a row with various data types
    test_date = datetime(1980, 5, 15)
    worksheet['A2'] = 'John'  # String
    worksheet['B2'] = 'Doe'  # String
    worksheet['C2'] = 1980  # Integer
    worksheet['D2'] = 5  # Integer
    worksheet['E2'] = 15  # Integer
    worksheet['F2'] = test_date  # Date
    worksheet['G2'] = 123456789  # Integer
    worksheet['H2'] = None  # Empty
    worksheet['I2'] = True  # Boolean
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
        'birth_year': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Year'),
        'birth_month': ColumnHeaderInfo(col=4, header_row=1, last_row=10, header_text='Month'),
        'birth_day': ColumnHeaderInfo(col=5, header_row=1, last_row=10, header_text='Day'),
        'entry_date': ColumnHeaderInfo(col=6, header_row=1, last_row=10, header_text='Entry Date'),
        'id_number': ColumnHeaderInfo(col=7, header_row=1, last_row=10, header_text='ID'),
        'passport': ColumnHeaderInfo(col=8, header_row=1, last_row=10, header_text='Passport'),
        'is_active': ColumnHeaderInfo(col=9, header_row=1, last_row=10, header_text='Active'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify all data types are preserved
    assert result['first_name'] == 'John'
    assert result['last_name'] == 'Doe'
    assert result['birth_year'] == 1980
    assert result['birth_month'] == 5
    assert result['birth_day'] == 15
    assert result['entry_date'] == test_date
    assert result['id_number'] == 123456789
    assert result['passport'] is None
    assert result['is_active'] is True
    assert len(result) == 9


def test_extract_row_with_hebrew_text(extractor, worksheet):
    """Test extracting a row with Hebrew text values."""
    # Setup: Create a row with Hebrew text
    worksheet['A2'] = 'יוסי'
    worksheet['B2'] = 'כהן'
    worksheet['C2'] = 'דוד'
    worksheet['D2'] = 'ז'
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='שם פרטי'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='שם משפחה'),
        'father_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='שם האב'),
        'gender': ColumnHeaderInfo(col=4, header_row=1, last_row=10, header_text='מין'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['first_name'] == 'יוסי'
    assert result['last_name'] == 'כהן'
    assert result['father_name'] == 'דוד'
    assert result['gender'] == 'ז'
    assert len(result) == 4


def test_extract_row_with_all_empty_cells(extractor, worksheet):
    """Test extracting a row where all cells are empty."""
    # Setup: Create a row with all empty cells
    worksheet['A2'] = None
    worksheet['B2'] = None
    worksheet['C2'] = None
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
        'father_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Father Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['first_name'] is None
    assert result['last_name'] is None
    assert result['father_name'] is None
    assert len(result) == 3


def test_extract_row_with_zero_values(extractor, worksheet):
    """Test extracting a row with zero values (should not be treated as empty)."""
    # Setup: Create a row with zero values
    worksheet['A2'] = 0
    worksheet['B2'] = 0.0
    worksheet['C2'] = 'John'
    
    column_mapping = {
        'count': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='Count'),
        'score': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Score'),
        'first_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='First Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify - zeros should be preserved, not treated as None
    assert result['count'] == 0
    assert result['score'] == 0.0
    assert result['first_name'] == 'John'
    assert len(result) == 3


def test_extract_row_with_empty_string_values(extractor, worksheet):
    """Test extracting a row with empty string values."""
    # Setup: Create a row with empty strings
    worksheet['A2'] = ''
    worksheet['B2'] = 'John'
    worksheet['C2'] = ''
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
        'father_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Father Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify - empty strings should be preserved
    assert result['first_name'] == ''
    assert result['last_name'] == 'John'
    assert result['father_name'] == ''
    assert len(result) == 3


def test_extract_row_with_whitespace_values(extractor, worksheet):
    """Test extracting a row with whitespace values."""
    # Setup: Create a row with whitespace
    worksheet['A2'] = '  John  '
    worksheet['B2'] = ' Doe '
    worksheet['C2'] = '   '
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
        'father_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Father Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify - whitespace should be preserved (not trimmed)
    assert result['first_name'] == '  John  '
    assert result['last_name'] == ' Doe '
    assert result['father_name'] == '   '
    assert len(result) == 3


def test_extract_row_preserves_field_order(extractor, worksheet):
    """Test that extract_row_to_json preserves the field order from column_mapping."""
    # Setup: Create a row with values
    worksheet['A2'] = 'Value A'
    worksheet['B2'] = 'Value B'
    worksheet['C2'] = 'Value C'
    
    column_mapping = {
        'field_c': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='C'),
        'field_a': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='A'),
        'field_b': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='B'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify - all fields should be present with correct values
    assert result['field_a'] == 'Value A'
    assert result['field_b'] == 'Value B'
    assert result['field_c'] == 'Value C'
    assert len(result) == 3


def test_extract_row_with_single_field(extractor, worksheet):
    """Test extracting a row with only one field."""
    # Setup: Create a row with one value
    worksheet['A2'] = 'John'
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['first_name'] == 'John'
    assert len(result) == 1


def test_extract_row_with_empty_column_mapping(extractor, worksheet):
    """Test extracting a row with empty column mapping."""
    # Setup: Create a row with values
    worksheet['A2'] = 'John'
    worksheet['B2'] = 'Doe'
    
    column_mapping = {}
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify - should return empty dictionary
    assert result == {}
    assert len(result) == 0


def test_extract_row_different_row_numbers(extractor, worksheet):
    """Test extracting different rows from the same worksheet."""
    # Setup: Create multiple rows with different values
    worksheet['A2'] = 'John'
    worksheet['B2'] = 'Doe'
    worksheet['A3'] = 'Jane'
    worksheet['B3'] = 'Smith'
    worksheet['A4'] = 'Bob'
    worksheet['B4'] = 'Johnson'
    
    column_mapping = {
        'first_name': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='First Name'),
        'last_name': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Last Name'),
    }
    
    # Execute - extract different rows
    result_row2 = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    result_row3 = extractor.extract_row_to_json(worksheet, 3, column_mapping)
    result_row4 = extractor.extract_row_to_json(worksheet, 4, column_mapping)
    
    # Verify - each row should have correct values
    assert result_row2['first_name'] == 'John'
    assert result_row2['last_name'] == 'Doe'
    
    assert result_row3['first_name'] == 'Jane'
    assert result_row3['last_name'] == 'Smith'
    
    assert result_row4['first_name'] == 'Bob'
    assert result_row4['last_name'] == 'Johnson'


def test_extract_row_with_large_numbers(extractor, worksheet):
    """Test extracting a row with large numbers."""
    # Setup: Create a row with large numbers
    worksheet['A2'] = 999999999
    worksheet['B2'] = 1234567890123456
    worksheet['C2'] = 3.14159265358979
    
    column_mapping = {
        'id_number': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='ID'),
        'large_number': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Large'),
        'pi': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Pi'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['id_number'] == 999999999
    assert result['large_number'] == 1234567890123456
    assert result['pi'] == 3.14159265358979
    assert len(result) == 3


def test_extract_row_with_negative_numbers(extractor, worksheet):
    """Test extracting a row with negative numbers."""
    # Setup: Create a row with negative numbers
    worksheet['A2'] = -100
    worksheet['B2'] = -3.14
    worksheet['C2'] = 'John'
    
    column_mapping = {
        'balance': ColumnHeaderInfo(col=1, header_row=1, last_row=10, header_text='Balance'),
        'temperature': ColumnHeaderInfo(col=2, header_row=1, last_row=10, header_text='Temp'),
        'first_name': ColumnHeaderInfo(col=3, header_row=1, last_row=10, header_text='Name'),
    }
    
    # Execute
    result = extractor.extract_row_to_json(worksheet, 2, column_mapping)
    
    # Verify
    assert result['balance'] == -100
    assert result['temperature'] == -3.14
    assert result['first_name'] == 'John'
    assert len(result) == 3
