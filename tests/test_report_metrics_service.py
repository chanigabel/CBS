"""Tests for processing report metrics service."""

import pytest
from unittest.mock import MagicMock

from webapp.services.report_metrics_service import ReportMetricsService
from webapp.models.processing_report import ProcessingReport, PerSheetProcessingReport, ReportSummary


class TestReportMetricsService:
    """Test suite for report metrics calculation."""

    def create_mock_session_with_dataset(self, sheets_data):
        """Create mock session and workbook dataset."""
        mock_session_service = MagicMock()
        
        # Create mock sheets
        mock_sheets = []
        for sheet_name, rows_data in sheets_data.items():
            mock_rows = [
                {
                    "_row_uid": f"{sheet_name}_row_{i}",
                    **row_data
                }
                for i, row_data in enumerate(rows_data)
            ]
            
            mock_sheet = MagicMock()
            mock_sheet.name = sheet_name
            mock_sheet.rows = mock_rows
            mock_sheets.append(mock_sheet)
        
        # Create mock dataset
        mock_dataset = MagicMock()
        mock_dataset.sheets = mock_sheets
        
        # Create mock record
        mock_record = MagicMock()
        mock_record.workbook_dataset = mock_dataset
        mock_record.edits = {}
        
        mock_session_service.get = MagicMock(return_value=mock_record)
        
        return mock_session_service, mock_record

    def test_collect_manual_edits_by_sheet(self):
        """Test collecting manual edits grouped by sheet."""
        mock_service = MagicMock()
        service = ReportMetricsService(mock_service)
        
        edits = {
            ("Sheet1", "row_1", "name"): "Alice",
            ("Sheet1", "row_2", "name"): "Bob",
            ("Sheet2", "row_1", "age"): "30",
        }
        
        result = service._collect_manual_edits_by_sheet(edits)
        
        assert "Sheet1" in result
        assert "Sheet2" in result
        assert result["Sheet1"] == {"row_1", "row_2"}
        assert result["Sheet2"] == {"row_1"}

    def test_estimate_auto_changes_by_sheet(self):
        """Test estimating auto-changed rows (rows with corrections)."""
        sheets_data = {
            "Sheet1": [
                {"first_name": "John", "first_name_corrected": "Jon"},
                {"first_name": "Jane", "first_name_corrected": "Jane"},
            ]
        }
        mock_service, mock_record = self.create_mock_session_with_dataset(sheets_data)
        
        service = ReportMetricsService(mock_service)
        manual_edits = {}
        
        result = service._estimate_auto_changes_by_sheet(
            mock_record.workbook_dataset, manual_edits
        )
        
        assert "Sheet1" in result
        assert result["Sheet1"] == 1  # Only first row has actual change

    def test_determine_sheet_status_completed(self):
        """Test sheet status determination for completed sheet."""
        per_sheet = PerSheetProcessingReport(sheet_name="Sheet1")
        per_sheet.total_rows = 10
        per_sheet.errors = []
        per_sheet.warnings = []
        
        service = ReportMetricsService(MagicMock())
        status = service._determine_sheet_status(per_sheet)
        
        assert status == "בוצע"

    def test_determine_sheet_status_with_warnings(self):
        """Test sheet status determination for sheet with warnings."""
        per_sheet = PerSheetProcessingReport(sheet_name="Sheet1")
        per_sheet.total_rows = 10
        per_sheet.errors = []
        per_sheet.warnings = ["Some warning"]
        
        service = ReportMetricsService(MagicMock())
        status = service._determine_sheet_status(per_sheet)
        
        assert status == "בוצע עם אזהרות"

    def test_determine_sheet_status_failed(self):
        """Test sheet status determination for failed sheet."""
        per_sheet = PerSheetProcessingReport(sheet_name="Sheet1")
        per_sheet.total_rows = 10
        per_sheet.errors = ["Error 1", "Error 2"]
        per_sheet.warnings = []
        
        service = ReportMetricsService(MagicMock())
        status = service._determine_sheet_status(per_sheet)
        
        assert status == "נכשל"

    def test_calculate_summary_aggregates_counts(self):
        """Test that summary aggregates per-sheet counts."""
        service = ReportMetricsService(MagicMock())
        
        report = ProcessingReport(session_id="sess_1")
        
        # Create per-sheet reports
        sheet1 = PerSheetProcessingReport(sheet_name="Sheet1")
        sheet1.total_rows = 100
        sheet1.rows_processed = 100
        sheet1.rows_exported = 100
        sheet1.rows_changed_automatically = 20
        sheet1.rows_changed_manually = 5
        sheet1.sheet_status = "בוצע"
        
        sheet2 = PerSheetProcessingReport(sheet_name="Sheet2")
        sheet2.total_rows = 50
        sheet2.rows_processed = 50
        sheet2.rows_exported = 50
        sheet2.rows_changed_automatically = 10
        sheet2.rows_changed_manually = 2
        sheet2.sheet_status = "בוצע עם אזהרות"
        sheet2.warnings = ["Warning 1"]
        
        report.per_sheet_warnings = [sheet1, sheet2]
        
        summary = service._calculate_summary(report)
        
        assert summary.total_rows == 150
        assert summary.rows_processed == 150
        assert summary.rows_exported == 150
        assert summary.rows_changed_automatically == 30
        assert summary.rows_changed_manually == 7
        assert summary.sheets_completed == 1
        assert summary.sheets_completed_with_warnings == 1

    def test_enrich_report_adds_summary(self):
        """Test that enrich_report adds summary to report."""
        sheets_data = {
            "Sheet1": [
                {"name": "Alice"},
                {"name": "Bob"},
            ]
        }
        mock_service, mock_record = self.create_mock_session_with_dataset(sheets_data)
        
        service = ReportMetricsService(mock_service)
        
        report = ProcessingReport(session_id="sess_1")
        report.per_sheet_warnings = [
            PerSheetProcessingReport(sheet_name="Sheet1", total_rows=2)
        ]
        
        enriched = service.enrich_report("sess_1", report)
        
        assert enriched.summary is not None
        assert isinstance(enriched.summary, ReportSummary)

    def test_enrich_report_without_dataset(self):
        """Test that enrich_report handles missing dataset gracefully."""
        mock_service = MagicMock()
        mock_record = MagicMock()
        mock_record.workbook_dataset = None
        mock_service.get = MagicMock(return_value=mock_record)
        
        service = ReportMetricsService(mock_service)
        
        report = ProcessingReport(session_id="sess_1")
        enriched = service.enrich_report("sess_1", report)
        
        # Should return the same report without modification
        assert enriched is report

    def test_row_has_corrections_detects_corrected_fields(self):
        """Test detection of corrected fields in rows."""
        service = ReportMetricsService(MagicMock())
        
        row_with_corrections = {
            "first_name": "John",
            "first_name_corrected": "Jon",
        }
        
        assert service._row_has_corrections(row_with_corrections) is True

    def test_row_has_corrections_ignores_unchanged(self):
        """Test that unchanged corrected fields are not reported as corrections."""
        service = ReportMetricsService(MagicMock())
        
        row_unchanged = {
            "first_name": "John",
            "first_name_corrected": "John",
        }
        
        assert service._row_has_corrections(row_unchanged) is False

    def test_row_has_corrections_empty_corrected_field(self):
        """Test that empty corrected fields are not reported as corrections."""
        service = ReportMetricsService(MagicMock())
        
        row_empty = {
            "first_name": "John",
            "first_name_corrected": "",
        }
        
        assert service._row_has_corrections(row_empty) is False

    def test_row_has_corrections_with_status(self):
        """Test detection of corrections via validation status."""
        service = ReportMetricsService(MagicMock())
        
        row_with_status = {
            "date_field": "2024-01-01",
            "_validation_status": "ערך תאריך לא תקין",
        }
        
        assert service._row_has_corrections(row_with_status) is True

    def test_enrich_report_tracks_manual_changes(self):
        """Test that manual edits are tracked in enriched report."""
        sheets_data = {
            "Sheet1": [
                {"name": "Alice"},
                {"name": "Bob"},
            ]
        }
        mock_service, mock_record = self.create_mock_session_with_dataset(sheets_data)
        
        # Add manual edits
        mock_record.edits = {
            ("Sheet1", "Sheet1_row_0", "name"): "Alicia",
            ("Sheet1", "Sheet1_row_1", "name"): "Robert",
        }
        
        service = ReportMetricsService(mock_service)
        
        report = ProcessingReport(session_id="sess_1")
        report.per_sheet_warnings = [
            PerSheetProcessingReport(sheet_name="Sheet1", total_rows=2)
        ]
        
        enriched = service.enrich_report("sess_1", report)
        
        assert enriched.rows_changed_manually == 2
