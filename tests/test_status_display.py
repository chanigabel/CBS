"""Tests for processing report status display."""

import pytest
from webapp.models.processing_report import ProcessingReport
from webapp.services.report_status_builder import refresh_status, status_reason


class TestStatusDisplay:
    """Test suite for status display in processing reports."""

    def test_status_success_when_no_warnings(self):
        """Test that status is 'success' when no warnings or errors."""
        report = ProcessingReport(session_id="sess_1")
        report.completed_stages = ["upload", "extract", "standardize"]
        report.warnings = []
        report.errors = []
        report.per_sheet_warnings = []
        
        refresh_status(report)
        
        assert report.status == "success"

    def test_status_partial_success_with_warnings(self):
        """Test that status is 'partial_success' when warnings exist."""
        report = ProcessingReport(session_id="sess_1")
        report.completed_stages = ["upload", "extract", "standardize"]
        report.warnings = ["Warning 1"]
        report.errors = []
        
        refresh_status(report)
        
        assert report.status == "partial_success"

    def test_status_partial_success_with_missing_columns(self):
        """Test that status is 'partial_success' with missing input columns."""
        from webapp.models.processing_report import MissingInputColumnsBySheet
        
        report = ProcessingReport(session_id="sess_1")
        report.missing_input_columns = [
            MissingInputColumnsBySheet(sheet_name="Sheet1", columns=["col1", "col2"])
        ]
        report.errors = []
        report.warnings = []
        
        refresh_status(report)
        
        assert report.status == "partial_success"

    def test_status_failed_with_errors(self):
        """Test that status is 'failed' when errors exist."""
        report = ProcessingReport(session_id="sess_1")
        report.errors = ["Error 1"]
        report.warnings = []
        
        refresh_status(report)
        
        assert report.status == "failed"

    def test_status_reason_for_success(self):
        """Test status reason text for success."""
        report = ProcessingReport(session_id="sess_1")
        report.status = "success"
        report.completed_stages = ["upload", "extract", "standardize"]
        
        reason = status_reason(report)
        
        assert "success" in reason.lower()

    def test_status_reason_for_partial_success(self):
        """Test status reason text for partial success."""
        report = ProcessingReport(session_id="sess_1")
        report.status = "partial_success"
        report.warnings = ["Warning 1"]
        
        reason = status_reason(report)
        
        assert "partial_success" in reason.lower()

    def test_status_reason_for_failure(self):
        """Test status reason text for failure."""
        report = ProcessingReport(session_id="sess_1")
        report.status = "failed"
        report.errors = ["Error 1", "Error 2"]
        
        reason = status_reason(report)
        
        assert "failed" in reason.lower()
        assert "2" in reason  # Should mention error count

    def test_hebrew_status_mapping_for_display(self):
        """Test conversion of status to Hebrew display text.
        
        This tests that "בוצע" (completed) text can be returned appropriately.
        """
        # Simulate the conversion that should happen in the UI/API
        status_to_hebrew = {
            "success": "בוצע",
            "partial_success": "בוצע עם אזהרות",
            "failed": "נכשל",
        }
        
        report = ProcessingReport(session_id="sess_1")
        report.status = "success"
        refresh_status(report)
        
        hebrew_status = status_to_hebrew.get(report.status, "unknown")
        assert hebrew_status == "בוצע"

    def test_hebrew_status_with_warnings(self):
        """Test Hebrew status when warnings exist."""
        status_to_hebrew = {
            "success": "בוצע",
            "partial_success": "בוצע עם אזהרות",
            "failed": "נכשל",
        }
        
        report = ProcessingReport(session_id="sess_1")
        report.status = "partial_success"
        
        hebrew_status = status_to_hebrew.get(report.status, "unknown")
        assert hebrew_status == "בוצע עם אזהרות"

    def test_status_reason_lists_warning_types(self):
        """Test that status reason includes warning type counts."""
        from webapp.models.processing_report import SummaryCount
        
        report = ProcessingReport(session_id="sess_1")
        report.status = "partial_success"
        report.date_summary = [
            SummaryCount(message="Invalid dates", count=5)
        ]
        
        reason = status_reason(report)
        
        assert "5" in reason  # Should mention count

    def test_refresh_status_updates_reason(self):
        """Test that refresh_status updates both status and reason."""
        report = ProcessingReport(session_id="sess_1")
        report.warnings = ["Test warning"]
        
        refresh_status(report)
        
        assert report.status == "partial_success"
        assert report.status_reason is not None
        assert len(report.status_reason) > 0

    def test_status_with_per_sheet_warnings(self):
        """Test that per-sheet warnings trigger partial_success status."""
        from webapp.models.processing_report import PerSheetProcessingReport
        
        report = ProcessingReport(session_id="sess_1")
        sheet_report = PerSheetProcessingReport(sheet_name="Sheet1")
        sheet_report.warnings = ["Sheet warning"]
        report.per_sheet_warnings = [sheet_report]
        
        refresh_status(report)
        
        assert report.status == "partial_success"

    def test_status_with_date_errors(self):
        """Test that date summary triggers appropriate status."""
        from webapp.models.processing_report import SummaryCount
        
        report = ProcessingReport(session_id="sess_1")
        report.date_summary = [
            SummaryCount(message="Invalid dates", count=3)
        ]
        
        refresh_status(report)
        
        assert report.status == "partial_success"

    def test_status_with_identifier_issues(self):
        """Test that identifier summary triggers appropriate status."""
        from webapp.models.processing_report import SummaryCount
        
        report = ProcessingReport(session_id="sess_1")
        report.identifier_summary = [
            SummaryCount(message="חסר מזהים", count=2)
        ]
        
        refresh_status(report)
        
        assert report.status == "partial_success"
