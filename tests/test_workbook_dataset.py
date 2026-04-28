"""Unit tests for WorkbookDataset dataclass."""

import pytest
from src.excel_standardization.data_types import WorkbookDataset, SheetDataset, JsonRow


class TestWorkbookDataset:
    """Test suite for WorkbookDataset dataclass."""

    def test_create_basic_workbook_dataset(self):
        """Test creating a basic WorkbookDataset."""
        sheet1 = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name", "last_name"],
            rows=[{"first_name": "John", "last_name": "Doe"}]
        )
        
        sheet2 = SheetDataset(
            sheet_name="Sheet2",
            header_row=2,
            header_rows_count=1,
            field_names=["gender", "id_number"],
            rows=[{"gender": "M", "id_number": "123456789"}]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet1, sheet2]
        )
        
        assert dataset.source_file == "test.xlsx"
        assert len(dataset.sheets) == 2
        assert dataset.sheets[0].sheet_name == "Sheet1"
        assert dataset.sheets[1].sheet_name == "Sheet2"
        assert dataset.metadata == {}

    def test_get_sheet_by_name(self):
        """Test get_sheet_by_name method."""
        sheet1 = SheetDataset(
            sheet_name="Students",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[]
        )
        
        sheet2 = SheetDataset(
            sheet_name="Teachers",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet1, sheet2]
        )
        
        # Test finding existing sheets
        found_sheet = dataset.get_sheet_by_name("Students")
        assert found_sheet is not None
        assert found_sheet.sheet_name == "Students"
        
        found_sheet = dataset.get_sheet_by_name("Teachers")
        assert found_sheet is not None
        assert found_sheet.sheet_name == "Teachers"
        
        # Test non-existent sheet
        not_found = dataset.get_sheet_by_name("NonExistent")
        assert not_found is None

    def test_get_sheet_names(self):
        """Test get_sheet_names method."""
        sheet1 = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        sheet2 = SheetDataset(
            sheet_name="Sheet2",
            header_row=1,
            header_rows_count=1,
            field_names=["field2"],
            rows=[]
        )
        
        sheet3 = SheetDataset(
            sheet_name="Sheet3",
            header_row=1,
            header_rows_count=1,
            field_names=["field3"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet1, sheet2, sheet3]
        )
        
        names = dataset.get_sheet_names()
        assert names == ["Sheet1", "Sheet2", "Sheet3"]

    def test_get_sheet_count(self):
        """Test get_sheet_count method."""
        sheet1 = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        sheet2 = SheetDataset(
            sheet_name="Sheet2",
            header_row=1,
            header_rows_count=1,
            field_names=["field2"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet1, sheet2]
        )
        
        assert dataset.get_sheet_count() == 2
        
        # Test empty workbook
        empty_dataset = WorkbookDataset(
            source_file="empty.xlsx",
            sheets=[]
        )
        assert empty_dataset.get_sheet_count() == 0

    def test_validate_valid_workbook(self):
        """Test validation of a valid workbook dataset."""
        sheet = SheetDataset(
            sheet_name="Valid",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[{"field1": "value1"}]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet]
        )
        
        assert dataset.validate() is True

    def test_validate_empty_source_file(self):
        """Test validation fails for empty source file."""
        sheet = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="",
            sheets=[sheet]
        )
        
        assert dataset.validate() is False

    def test_validate_non_list_sheets(self):
        """Test validation fails for non-list sheets."""
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets="not a list"  # type: ignore
        )
        
        assert dataset.validate() is False

    def test_validate_non_sheetdataset_items(self):
        """Test validation fails for non-SheetDataset items in sheets list."""
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[{"not": "a SheetDataset"}]  # type: ignore
        )
        
        assert dataset.validate() is False

    def test_validate_duplicate_sheet_names(self):
        """Test validation fails for duplicate sheet names."""
        sheet1 = SheetDataset(
            sheet_name="Duplicate",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        sheet2 = SheetDataset(
            sheet_name="Duplicate",
            header_row=1,
            header_rows_count=1,
            field_names=["field2"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet1, sheet2]
        )
        
        assert dataset.validate() is False

    def test_validate_invalid_sheet(self):
        """Test validation fails when a sheet is invalid."""
        invalid_sheet = SheetDataset(
            sheet_name="",  # Invalid: empty name
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[invalid_sheet]
        )
        
        assert dataset.validate() is False

    def test_metadata_default_empty(self):
        """Test metadata defaults to empty dict."""
        sheet = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet]
        )
        
        assert dataset.metadata == {}

    def test_metadata_with_values(self):
        """Test creating dataset with metadata."""
        sheet = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        metadata = {
            "extraction_date": "2024-01-15",
            "total_sheets": 3,
            "processed_sheets": 2,
            "skipped_sheets": ["Summary"]
        }
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet],
            metadata=metadata
        )
        
        assert dataset.metadata == metadata

    def test_get_metadata(self):
        """Test get_metadata method."""
        sheet = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet],
            metadata={"key1": "value1", "key2": 42}
        )
        
        assert dataset.get_metadata("key1") == "value1"
        assert dataset.get_metadata("key2") == 42
        assert dataset.get_metadata("nonexistent") is None
        assert dataset.get_metadata("nonexistent", "default") == "default"

    def test_set_metadata(self):
        """Test set_metadata method."""
        sheet = SheetDataset(
            sheet_name="Sheet1",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet]
        )
        
        dataset.set_metadata("new_key", "new_value")
        assert dataset.metadata["new_key"] == "new_value"
        
        dataset.set_metadata("another_key", 123)
        assert dataset.metadata["another_key"] == 123

    def test_has_sheet(self):
        """Test has_sheet method."""
        sheet1 = SheetDataset(
            sheet_name="Exists",
            header_row=1,
            header_rows_count=1,
            field_names=["field1"],
            rows=[]
        )
        
        dataset = WorkbookDataset(
            source_file="test.xlsx",
            sheets=[sheet1]
        )
        
        assert dataset.has_sheet("Exists") is True
        assert dataset.has_sheet("DoesNotExist") is False

    def test_multi_sheet_workbook(self):
        """Test workbook with multiple sheets containing data."""
        sheet1 = SheetDataset(
            sheet_name="Students",
            header_row=2,
            header_rows_count=1,
            field_names=["first_name", "last_name", "gender"],
            rows=[
                {"first_name": "יוסי", "last_name": "כהן", "gender": "ז"},
                {"first_name": "שרה", "last_name": "לוי", "gender": "נ"}
            ],
            metadata={"total_rows": 2}
        )
        
        sheet2 = SheetDataset(
            sheet_name="Teachers",
            header_row=1,
            header_rows_count=1,
            field_names=["first_name", "last_name"],
            rows=[
                {"first_name": "דוד", "last_name": "מזרחי"}
            ],
            metadata={"total_rows": 1}
        )
        
        dataset = WorkbookDataset(
            source_file="school.xlsx",
            sheets=[sheet1, sheet2],
            metadata={
                "extraction_date": "2024-01-15",
                "total_sheets": 2,
                "processed_sheets": 2
            }
        )
        
        assert dataset.get_sheet_count() == 2
        assert dataset.validate() is True
        
        students = dataset.get_sheet_by_name("Students")
        assert students is not None
        assert students.get_row_count() == 2
        
        teachers = dataset.get_sheet_by_name("Teachers")
        assert teachers is not None
        assert teachers.get_row_count() == 1

    def test_empty_workbook(self):
        """Test workbook with no sheets."""
        dataset = WorkbookDataset(
            source_file="empty.xlsx",
            sheets=[]
        )
        
        assert dataset.get_sheet_count() == 0
        assert dataset.get_sheet_names() == []
        assert dataset.get_sheet_by_name("Any") is None
        assert dataset.has_sheet("Any") is False
        # Empty workbook is still valid
        assert dataset.validate() is True
