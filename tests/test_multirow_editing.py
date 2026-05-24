"""Tests for multi-row editing and undo stack functionality."""

import pytest
from unittest.mock import MagicMock, Mock
from fastapi import HTTPException

from webapp.services.multirow_edit_service import (
    MultiRowEditService,
    MultiRowEditResult,
    UndoStack,
)


class TestMultiRowEditService:
    """Test suite for multi-row editing."""

    def create_mock_session_service_with_sheet(self, rows_data):
        """Create a mock session service with sheet data."""
        mock_session = MagicMock()
        
        # Create mock rows
        mock_rows = []
        for i, row_data in enumerate(rows_data):
            row = {
                "_row_uid": f"row_{i}",
                **row_data
            }
            mock_rows.append(row)
        
        # Create mock sheet
        mock_sheet = MagicMock()
        mock_sheet.name = "TestSheet"
        mock_sheet.rows = mock_rows
        
        # Create mock workbook dataset
        mock_dataset = MagicMock()
        mock_dataset.sheets = [mock_sheet]
        mock_dataset.get_sheet_by_name = MagicMock(return_value=mock_sheet)
        
        # Create mock session record
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.edits = {}
        mock_record.working_dataset_dirty = False
        
        # Create mock session service
        mock_session_service = MagicMock()
        mock_session_service.get = MagicMock(return_value=mock_record)
        mock_session_service.update = MagicMock()
        
        return mock_session_service, mock_record

    def test_edit_multiple_rows_basic(self):
        """Test basic multi-row edit."""
        rows_data = [
            {"name": "Alice", "age": 30},
            {"name": "Bob", "age": 25},
            {"name": "Charlie", "age": 35},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        
        service = MultiRowEditService(mock_service)
        
        result = service.edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=["row_0", "row_1"],
            field_name="age",
            new_value="40",
        )
        
        assert result.edited_count == 2
        assert "row_0" in result.updated_rows
        assert "row_1" in result.updated_rows
        assert result.updated_rows["row_0"]["age"] == 40
        assert result.updated_rows["row_1"]["age"] == 40

    def test_edit_multiple_rows_with_uids_from_actual_sheet_data(self):
        rows_data = [
            {"name": "Alice", "SugMosad": "100"},
            {"name": "Bob", "SugMosad": "100"},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        sheet = mock_record.workbook_dataset.get_sheet_by_name("TestSheet")
        row_uids = [row["_row_uid"] for row in sheet.rows]

        result = MultiRowEditService(mock_service).edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=row_uids,
            field_name="SugMosad",
            new_value="123",
        )

        assert result.edited_count == 2
        assert [row["SugMosad"] for row in sheet.rows] == ["123", "123"]

    def test_edit_multiple_rows_missing_uid_returns_clear_error_without_mutation(self):
        rows_data = [
            {"name": "Alice", "SugMosad": "100"},
            {"name": "Bob", "SugMosad": "100"},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        sheet = mock_record.workbook_dataset.get_sheet_by_name("TestSheet")

        with pytest.raises(HTTPException) as exc_info:
            MultiRowEditService(mock_service).edit_multiple_rows(
                session_id="sess_1",
                sheet_name="TestSheet",
                row_uids=["row_0", "missing-row"],
                field_name="SugMosad",
                new_value="999",
            )

        assert exc_info.value.status_code == 404
        assert exc_info.value.detail["sheet_name"] == "TestSheet"
        assert exc_info.value.detail["requested_rows"] == 2
        assert exc_info.value.detail["found_rows"] == 1
        assert exc_info.value.detail["missing_rows"] == 1
        assert [row["SugMosad"] for row in sheet.rows] == ["100", "100"]
        assert mock_record.edits == {}

    def test_edit_multiple_rows_sheet_name_mismatch_returns_clear_error(self):
        rows_data = [{"name": "Alice"}]
        mock_service, _mock_record = self.create_mock_session_service_with_sheet(rows_data)
        mock_service.get.return_value.workbook_dataset.get_sheet_by_name = MagicMock(return_value=None)

        with pytest.raises(HTTPException) as exc_info:
            MultiRowEditService(mock_service).edit_multiple_rows(
                session_id="sess_1",
                sheet_name="WrongSheet",
                row_uids=["row_0"],
                field_name="name",
                new_value="Bob",
            )

        assert exc_info.value.status_code == 404
        assert "לא נמצא" in exc_info.value.detail

    def test_edit_multiple_rows_can_create_displayed_institution_field(self):
        rows_data = [{"name": "Alice"}, {"name": "Bob"}]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        sheet = mock_record.workbook_dataset.get_sheet_by_name("TestSheet")

        MultiRowEditService(mock_service).edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=["row_0", "row_1"],
            field_name="SugMosad",
            new_value="222",
        )

        assert [row["SugMosad"] for row in sheet.rows] == ["222", "222"]

    def test_multirow_edit_values_are_visible_to_export_rows(self):
        from webapp.services.export_rows import visible_rows

        rows_data = [{"name": "Alice"}, {"name": "Bob"}]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        sheet = mock_record.workbook_dataset.get_sheet_by_name("TestSheet")
        sheet.field_names = ["name"]

        MultiRowEditService(mock_service).edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=["row_0", "row_1"],
            field_name="SugMosad",
            new_value="333",
        )

        rows, _columns = visible_rows(sheet)
        assert [row["SugMosad"] for row in rows] == ["333", "333"]

    def test_edit_multiple_rows_type_coercion(self):
        """Test that new value is coerced to correct type."""
        rows_data = [
            {"name": "Alice", "count": 10},
            {"name": "Bob", "count": 20},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        
        service = MultiRowEditService(mock_service)
        
        # Pass string "99" but expect it to be coerced to int 99
        result = service.edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=["row_0", "row_1"],
            field_name="count",
            new_value="99",
        )
        
        assert isinstance(result.updated_rows["row_0"]["count"], int)
        assert result.updated_rows["row_0"]["count"] == 99

    def test_edit_multiple_rows_marks_dataset_dirty(self):
        """Test that edit marks working dataset as dirty."""
        rows_data = [
            {"name": "Alice"},
            {"name": "Bob"},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        
        service = MultiRowEditService(mock_service)
        
        service.edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=["row_0"],
            field_name="name",
            new_value="Charlie",
        )
        
        # Verify update was called with working_dataset_dirty=True
        assert mock_service.update.called
        call_kwargs = mock_service.update.call_args[1]
        assert call_kwargs["working_dataset_dirty"] is True

    def test_edit_multiple_rows_records_edits(self):
        """Test that edits are recorded in session edits dict."""
        rows_data = [
            {"name": "Alice"},
            {"name": "Bob"},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        
        service = MultiRowEditService(mock_service)
        
        service.edit_multiple_rows(
            session_id="sess_1",
            sheet_name="TestSheet",
            row_uids=["row_0", "row_1"],
            field_name="name",
            new_value="David",
        )
        
        # Verify edit keys were added
        assert ("TestSheet", "row_0", "name") in mock_record.edits
        assert ("TestSheet", "row_1", "name") in mock_record.edits
        assert mock_record.edits[("TestSheet", "row_0", "name")] == "David"
        assert mock_record.edits[("TestSheet", "row_1", "name")] == "David"

    def test_edit_multiple_rows_invalid_field_raises_error(self):
        """Test that editing internal field raises HTTPException."""
        rows_data = [
            {"name": "Alice", "_row_uid": "row_0"},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        
        service = MultiRowEditService(mock_service)
        
        with pytest.raises(HTTPException) as exc_info:
            service.edit_multiple_rows(
                session_id="sess_1",
                sheet_name="TestSheet",
                row_uids=["row_0"],
                field_name="_row_uid",
                new_value="invalid",
            )
        
        assert exc_info.value.status_code == 400
        assert "internal field" in exc_info.value.detail.lower()

    def test_edit_multiple_rows_missing_row_uid_raises_error(self):
        """Test that non-existent row UID raises HTTPException."""
        rows_data = [
            {"name": "Alice"},
        ]
        mock_service, mock_record = self.create_mock_session_service_with_sheet(rows_data)
        
        service = MultiRowEditService(mock_service)
        
        with pytest.raises(HTTPException) as exc_info:
            service.edit_multiple_rows(
                session_id="sess_1",
                sheet_name="TestSheet",
                row_uids=["row_999"],
                field_name="name",
                new_value="Bob",
        )
    
        assert exc_info.value.status_code == 404
        assert exc_info.value.detail["message"] == "הבחירה בגריד אינה מעודכנת. נא לבחור את השורות מחדש ולנסות שוב."
        assert exc_info.value.detail["missing_rows"] == 1

    def test_edit_multiple_rows_sheet_not_found_raises_error(self):
        """Test that missing sheet raises HTTPException."""
        mock_service = MagicMock()
        mock_dataset = MagicMock()
        mock_dataset.get_sheet_by_name = MagicMock(return_value=None)
        
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_service.get = MagicMock(return_value=mock_record)
        
        service = MultiRowEditService(mock_service)
        
        with pytest.raises(HTTPException) as exc_info:
            service.edit_multiple_rows(
                session_id="sess_1",
                sheet_name="NonExistent",
                row_uids=["row_0"],
                field_name="name",
                new_value="Bob",
        )
    
        assert exc_info.value.status_code == 404
        assert "לא נמצא" in exc_info.value.detail


class TestUndoStack:
    """Test suite for undo stack functionality."""

    def test_undo_stack_can_undo(self):
        """Test that undo stack tracks undoable state."""
        stack = UndoStack()
        assert not stack.can_undo()
        
        stack.push_edit(
            sheet_name="Sheet1",
            row_uids=["row_1"],
            field_name="name",
            new_value="New",
            old_values={"row_1": "Old"},
        )
        
        assert stack.can_undo()

    def test_undo_stack_undo_and_redo(self):
        """Test undo and redo functionality."""
        stack = UndoStack()
        
        stack.push_edit(
            sheet_name="Sheet1",
            row_uids=["row_1", "row_2"],
            field_name="status",
            new_value="Active",
            old_values={"row_1": "Inactive", "row_2": "Inactive"},
        )
        
        assert stack.can_undo()
        assert not stack.can_redo()
        
        # Undo
        undo_result = stack.undo()
        assert undo_result is not None
        sheet_name, new_value, row_uids, field_name, old_values = undo_result
        assert new_value == "Active"
        assert len(row_uids) == 2
        
        assert not stack.can_undo()
        assert stack.can_redo()
        
        # Redo
        redo_result = stack.redo()
        assert redo_result is not None
        assert stack.can_undo()
        assert not stack.can_redo()

    def test_undo_stack_clears_redo_on_new_edit(self):
        """Test that new edit clears redo stack."""
        stack = UndoStack()
        
        # Push, undo, then push again
        stack.push_edit("Sheet1", ["row_1"], "field", "val1", {"row_1": "old1"})
        stack.undo()
        assert stack.can_redo()
        
        # New edit should clear redo
        stack.push_edit("Sheet1", ["row_2"], "field", "val2", {"row_2": "old2"})
        assert not stack.can_redo()

    def test_undo_stack_multiple_edits(self):
        """Test multiple edits on undo stack."""
        stack = UndoStack()
        
        stack.push_edit("Sheet1", ["row_1"], "name", "Alice", {"row_1": "A"})
        stack.push_edit("Sheet1", ["row_2"], "age", "30", {"row_2": "25"})
        stack.push_edit("Sheet1", ["row_3"], "status", "active", {"row_3": "inactive"})
        
        # Should undo in reverse order
        result1 = stack.undo()
        assert result1[3] == "status"
        
        result2 = stack.undo()
        assert result2[3] == "age"
        
        result3 = stack.undo()
        assert result3[3] == "name"
        
        # No more to undo
        assert stack.undo() is None

    def test_undo_returns_all_edit_info(self):
        """Test that undo returns complete edit information."""
        stack = UndoStack()
        
        old_vals = {"row_1": "OldValue1", "row_2": "OldValue2"}
        stack.push_edit(
            sheet_name="MySheet",
            row_uids=["row_1", "row_2"],
            field_name="type",
            new_value="NewType",
            old_values=old_vals,
        )
        
        sheet_name, new_value, row_uids, field_name, old_values = stack.undo()
        
        assert sheet_name == "MySheet"
        assert new_value == "NewType"
        assert row_uids == ["row_1", "row_2"]
        assert field_name == "type"
        assert old_values == old_vals

    def test_undo_stack_empty_undo_returns_none(self):
        """Test that undoing empty stack returns None."""
        stack = UndoStack()
        assert stack.undo() is None
        assert not stack.can_undo()

    def test_undo_stack_empty_redo_returns_none(self):
        """Test that redoing empty stack returns None."""
        stack = UndoStack()
        assert stack.redo() is None
        assert not stack.can_redo()
