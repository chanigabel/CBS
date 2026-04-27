"""Basic tests for ExcelToJsonExtractor class structure.

Tests the initialization and basic structure of the ExcelToJsonExtractor class.
"""

import pytest
from src.excel_normalization.io_layer import ExcelReader, ExcelToJsonExtractor


def test_extractor_initialization_default():
    """Test ExcelToJsonExtractor can be initialized with default options."""
    reader = ExcelReader()
    extractor = ExcelToJsonExtractor(excel_reader=reader)
    
    assert extractor.excel_reader is reader
    assert extractor.skip_empty_rows is False
    assert extractor.handle_formulas is True
    assert extractor.preserve_types is True
    assert extractor.max_scan_rows == 30


def test_extractor_initialization_custom_options():
    """Test ExcelToJsonExtractor can be initialized with custom options."""
    reader = ExcelReader()
    extractor = ExcelToJsonExtractor(
        excel_reader=reader,
        skip_empty_rows=True,
        handle_formulas=False,
        preserve_types=False,
        max_scan_rows=50,
    )
    
    assert extractor.excel_reader is reader
    assert extractor.skip_empty_rows is True
    assert extractor.handle_formulas is False
    assert extractor.preserve_types is False
    assert extractor.max_scan_rows == 50


def test_extractor_has_required_methods():
    """Test ExcelToJsonExtractor has all required methods."""
    reader = ExcelReader()
    extractor = ExcelToJsonExtractor(excel_reader=reader)
    
    # Check that all required methods exist
    assert hasattr(extractor, 'extract_row_to_json')
    assert hasattr(extractor, 'extract_sheet_to_json')
    assert hasattr(extractor, 'extract_workbook_to_json')
    
    # Check methods are callable
    assert callable(extractor.extract_row_to_json)
    assert callable(extractor.extract_sheet_to_json)
    assert callable(extractor.extract_workbook_to_json)


def test_extractor_accepts_excel_reader_dependency():
    """Test ExcelToJsonExtractor properly accepts ExcelReader dependency."""
    reader = ExcelReader()
    extractor = ExcelToJsonExtractor(excel_reader=reader)
    
    # Verify the dependency is stored correctly
    assert isinstance(extractor.excel_reader, ExcelReader)
    assert extractor.excel_reader is reader
