"""Unit tests for SheetDataset dataclass."""

import pytest
from src.excel_standardization.data_types import SheetDataset, JsonRow


class TestSheetDataset:
    """Test suite for SheetDataset dataclass."""

    def test_create_basic_dataset(self):
        """Test creating a basic SheetDataset."""
        dataset = SheetDataset(
            sheet_name="Test Sheet",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name", "last_name"],
            rows=[
                {"first_name": "John", "last_name": "Doe"},
                {"first_name": "Jane", "last_name": "Smith"}
            ]
        )
        
        assert dataset.sheet_name == "Test Sheet"
        assert dataset.header_row == 1
        assert dataset.header_rows_count == 1
        assert dataset.field_names == ["first_name", "last_name"]
        assert len(dataset.rows) == 2

    def test_get_field_names(self):
        """Test get_field_names helper method."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name", "last_name", "gender"],
            rows=[]
        )
        
        assert dataset.get_field_names() == ["first_name", "last_name", "gender"]

    def test_get_row_count(self):
        """Test get_row_count helper method."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[
                {"first_name": "John"},
                {"first_name": "Jane"},
                {"first_name": "Bob"}
            ]
        )
        
        assert dataset.get_row_count() == 3

    def test_validate_valid_dataset(self):
        """Test validation of a valid dataset."""
        dataset = SheetDataset(
            sheet_name="Valid",
            header_row=2,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[{"first_name": "John"}]
        )
        
        assert dataset.validate() is True

    def test_validate_empty_sheet_name(self):
        """Test validation fails for empty sheet name."""
        dataset = SheetDataset(
            sheet_name="",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[]
        )
        
        assert dataset.validate() is False

    def test_validate_invalid_header_row(self):
        """Test validation fails for invalid header row."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=0,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[]
        )
        
        assert dataset.validate() is False

    def test_validate_invalid_header_rows_count(self):
        """Test validation fails for invalid header rows count."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=3,
            field_names=["first_name"],
            rows=[]
        )
        
        assert dataset.validate() is False

    def test_validate_empty_field_names(self):
        """Test validation fails for empty field names."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=[],
            rows=[]
        )
        
        assert dataset.validate() is False

    def test_validate_non_dict_rows(self):
        """Test validation fails for non-dictionary rows."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=["not a dict"]
        )
        
        assert dataset.validate() is False

    def test_metadata_default_empty(self):
        """Test metadata defaults to empty dict."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[]
        )
        
        assert dataset.metadata == {}

    def test_metadata_with_values(self):
        """Test creating dataset with metadata."""
        metadata = {
            "source_file": "test.xlsx",
            "total_rows": 100,
            "date_field_structure": {"birth_date": "split"}
        }
        
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[],
            metadata=metadata
        )
        
        assert dataset.metadata == metadata

    def test_get_metadata(self):
        """Test get_metadata method."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[],
            metadata={"source_file": "test.xlsx", "total_rows": 50}
        )
        
        assert dataset.get_metadata("source_file") == "test.xlsx"
        assert dataset.get_metadata("total_rows") == 50
        assert dataset.get_metadata("missing_key") is None
        assert dataset.get_metadata("missing_key", "default") == "default"

    def test_set_metadata(self):
        """Test set_metadata method."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[]
        )
        
        dataset.set_metadata("source_file", "test.xlsx")
        dataset.set_metadata("total_rows", 100)
        
        assert dataset.metadata["source_file"] == "test.xlsx"
        assert dataset.metadata["total_rows"] == 100

    def test_multi_row_headers(self):
        """Test dataset with multi-row headers."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=2,
            field_names=["birth_year", "birth_month", "birth_day"],
            rows=[
                {"birth_year": 1980, "birth_month": 5, "birth_day": 15}
            ]
        )
        
        assert dataset.header_rows_count == 2
        assert dataset.validate() is True

    def test_with_corrected_fields(self):
        """Test dataset with original and corrected fields."""
        dataset = SheetDataset(
            sheet_name="Test",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name", "gender"],
            rows=[
                {
                    "first_name": "יוסי",
                    "first_name_corrected": "יוסי",
                    "gender": "ז",
                    "gender_corrected": "2"
                }
            ]
        )
        
        assert dataset.get_row_count() == 1
        assert dataset.rows[0]["first_name"] == "יוסי"
        assert dataset.rows[0]["first_name_corrected"] == "יוסי"
        assert dataset.rows[0]["gender"] == "ז"
        assert dataset.rows[0]["gender_corrected"] == "2"
