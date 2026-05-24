"""Tests for field mapping duplicate prevention."""

import pytest
from unittest.mock import MagicMock
from fastapi import HTTPException

from webapp.services.workbook_service import WorkbookService
from webapp.services.column_mapping_schema import ColumnMappingSchemaService


class TestFieldMappingDuplicatePrevention:
    """Test suite for preventing duplicate field mappings."""

    def create_mock_workbook_service_with_schema(self):
        """Create a mock workbook service with column schema."""
        mock_session = MagicMock()
        mock_schema = MagicMock(spec=ColumnMappingSchemaService)
        mock_schema.fields.return_value = [
            "first_name", "last_name", "gender", "birth_date"
        ]
        mock_schema.resolve.side_effect = lambda x: x  # Return as-is
        
        service = WorkbookService(mock_session, mock_schema)
        return service, mock_session, mock_schema

    def create_mock_sheet(self, field_names):
        """Create a mock sheet with given field names."""
        mock_sheet = MagicMock()
        mock_sheet.sheet_name = "TestSheet"
        mock_sheet.field_names = field_names
        mock_sheet.rows = [
            {"_row_uid": "row_0", **{f: f"val_{i}" for i, f in enumerate(field_names)}}
        ]
        return mock_sheet

    def create_mock_dataset(self, sheets):
        """Create a mock dataset with sheets."""
        mock_dataset = MagicMock()
        mock_dataset.sheets = sheets
        mock_dataset.get_sheet_by_name = lambda name: next(
            (s for s in sheets if s.sheet_name == name), None
        )
        return mock_dataset

    def test_single_mapping_is_allowed(self):
        """Test that a single mapping is allowed."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet = self.create_mock_sheet(["name1", "name2", "age"])
        mock_dataset = self.create_mock_dataset([mock_sheet])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.column_mappings = {}
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        mock_session.update = MagicMock()
        
        service._ensure_sheet_loaded = MagicMock()
        
        result = service.update_column_mapping(
            session_id="sess_1",
            sheet_name="TestSheet",
            old_name="name1",
            new_name="first_name",
        )
        
        assert result.sheet_name == "TestSheet"
        assert result.old_name == "name1"
        assert result.new_name == "first_name"

    def test_duplicate_mapping_is_rejected(self):
        """Test that duplicate mapping to same target is rejected."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet = self.create_mock_sheet(["name1", "name2", "age"])
        mock_dataset = self.create_mock_dataset([mock_sheet])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.column_mappings = {
            "TestSheet": {"name1": "first_name"}  # Already mapped
        }
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        service._ensure_sheet_loaded = MagicMock()
        
        with pytest.raises(HTTPException) as exc_info:
            service.update_column_mapping(
                session_id="sess_1",
                sheet_name="TestSheet",
                old_name="name2",
                new_name="first_name",  # Same target as name1
            )
        
        assert exc_info.value.status_code == 400
        assert "כפולים" in exc_info.value.detail or "duplicate" in exc_info.value.detail.lower()

    def test_hebrew_error_message_for_duplicate(self):
        """Test that error message includes Hebrew text for duplicates."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet = self.create_mock_sheet(["name1", "name2"])
        mock_dataset = self.create_mock_dataset([mock_sheet])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.column_mappings = {"TestSheet": {"name1": "first_name"}}
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        service._ensure_sheet_loaded = MagicMock()
        
        with pytest.raises(HTTPException) as exc_info:
            service.update_column_mapping(
                session_id="sess_1",
                sheet_name="TestSheet",
                old_name="name2",
                new_name="first_name",
            )
        
        detail = exc_info.value.detail
        assert "שדה סטנדרטי" in detail

    def test_mapping_change_reverted_on_duplicate_error(self):
        """Test that mapping is reverted if duplicate detected."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet = self.create_mock_sheet(["name1", "name2"])
        mock_dataset = self.create_mock_dataset([mock_sheet])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        original_mappings = {"TestSheet": {"name1": "first_name"}}
        mock_record.column_mappings = dict(original_mappings)  # Make a copy
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        service._ensure_sheet_loaded = MagicMock()
        
        with pytest.raises(HTTPException):
            service.update_column_mapping(
                session_id="sess_1",
                sheet_name="TestSheet",
                old_name="name2",
                new_name="first_name",
            )
        
        # Mappings should be unchanged
        assert mock_record.column_mappings["TestSheet"]["name1"] == "first_name"
        assert "name2" not in mock_record.column_mappings["TestSheet"]

    def test_multiple_different_mappings_allowed(self):
        """Test that different targets can be mapped."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet = self.create_mock_sheet(["name1", "name2", "name3"])
        mock_dataset = self.create_mock_dataset([mock_sheet])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.column_mappings = {
            "TestSheet": {
                "name1": "first_name",
                "name2": "last_name",
            }
        }
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        service._ensure_sheet_loaded = MagicMock()
        
        # This should work - different targets
        result = service.update_column_mapping(
            session_id="sess_1",
            sheet_name="TestSheet",
            old_name="name3",
            new_name="gender",
        )
        
        assert result.new_name == "gender"

    def test_duplicate_validation_per_sheet(self):
        """Test that duplicate validation is per-sheet."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet1 = self.create_mock_sheet(["name1", "name2"])
        mock_sheet1.sheet_name = "Sheet1"
        mock_sheet2 = self.create_mock_sheet(["name1", "name2"])
        mock_sheet2.sheet_name = "Sheet2"
        
        mock_dataset = self.create_mock_dataset([mock_sheet1, mock_sheet2])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.column_mappings = {
            "Sheet1": {"name1": "first_name"},
            "Sheet2": {},
        }
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        service._ensure_sheet_loaded = MagicMock()
        
        # Should be allowed to map "name1" to "first_name" in Sheet2
        # because it's a different sheet
        result = service.update_column_mapping(
            session_id="sess_1",
            sheet_name="Sheet2",
            old_name="name1",
            new_name="first_name",
        )
        
        assert result.new_name == "first_name"

    def test_changing_existing_mapping_is_allowed(self):
        """Test that changing an existing mapping to a different target is allowed."""
        service, mock_session, mock_schema = self.create_mock_workbook_service_with_schema()
        
        mock_sheet = self.create_mock_sheet(["name1", "name2"])
        mock_dataset = self.create_mock_dataset([mock_sheet])
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.column_mappings = {
            "TestSheet": {"name1": "first_name"}
        }
        mock_record.status = "uploaded"
        
        mock_session.get = MagicMock(return_value=mock_record)
        service._ensure_sheet_loaded = MagicMock()
        
        # Change name1 from first_name to last_name
        result = service.update_column_mapping(
            session_id="sess_1",
            sheet_name="TestSheet",
            old_name="name1",
            new_name="last_name",
        )
        
        assert result.new_name == "last_name"
