"""Service for enriching and calculating processing report metrics."""

from __future__ import annotations

import logging
from typing import Dict, List, Optional, Set
from collections import defaultdict

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from webapp.models.processing_report import ProcessingReport, PerSheetProcessingReport, ReportSummary
from webapp.services.session_service import SessionService

logger = logging.getLogger(__name__)


class ReportMetricsService:
    """Calculate and enrich processing report metrics.
    
    Responsibility: Generate accurate counts and statistics from session data,
    including manually edited rows, automatically changed rows, and per-sheet summaries.
    """

    def __init__(self, session_service: SessionService) -> None:
        self.session_service = session_service

    def enrich_report(
        self, session_id: str, report: ProcessingReport
    ) -> ProcessingReport:
        """Enrich a processing report with calculated metrics.
        
        Args:
            session_id: The session ID
            report: The report to enrich
            
        Returns:
            Updated report with enriched metrics
        """
        record = self.session_service.get(session_id)
        
        if record.workbook_dataset is None:
            # Cannot calculate metrics without dataset
            logger.warning(
                "report_enrichment_no_dataset",
                extra={
                    "event": "report_enrichment_no_dataset",
                    "session_id": session_id,
                },
            )
            return report

        # Track manual changes and build per-sheet metrics
        manual_edits_by_sheet = self._collect_manual_edits_by_sheet(record.edits)
        auto_changes_by_sheet = self._estimate_auto_changes_by_sheet(
            record.workbook_dataset, manual_edits_by_sheet
        )

        # Update per-sheet reports with new metrics
        self._update_per_sheet_reports(
            report, record.workbook_dataset, manual_edits_by_sheet, auto_changes_by_sheet
        )

        # Calculate overall summary
        summary = self._calculate_summary(report)
        report.summary = summary

        # Update top-level counters
        report.rows_changed_manually = sum(
            len(edits) for edits in manual_edits_by_sheet.values()
        )
        report.rows_changed_automatically = sum(
            auto_changes_by_sheet.values()
        )

        logger.info(
            "report_enriched",
            extra={
                "event": "report_enriched",
                "session_id": session_id,
                "manual_changes": report.rows_changed_manually,
                "auto_changes": report.rows_changed_automatically,
            },
        )

        return report

    def _collect_manual_edits_by_sheet(self, edits: dict) -> Dict[str, Set[str]]:
        """Extract set of unique row_uids that were manually edited, per sheet.
        
        Edits dict is shaped as: {(sheet_name, row_uid, field): new_value}
        
        Returns:
            Dict mapping sheet_name -> set of row_uids that were edited
        """
        by_sheet: Dict[str, Set[str]] = defaultdict(set)
        
        for key in edits.keys():
            if isinstance(key, tuple) and len(key) >= 2:
                sheet_name = key[0]
                row_uid = key[1]
                by_sheet[sheet_name].add(row_uid)
        
        return dict(by_sheet)

    def _estimate_auto_changes_by_sheet(
        self,
        workbook_dataset: WorkbookDataset,
        manual_edits_by_sheet: Dict[str, Set[str]],
    ) -> Dict[str, int]:
        """Estimate number of rows changed automatically by standardization.
        
        This is a heuristic: look for rows with corrected fields that differ from
        original fields, excluding rows that were manually edited.
        
        Args:
            workbook_dataset: The processed workbook dataset
            manual_edits_by_sheet: Set of manually edited row_uids per sheet
            
        Returns:
            Dict mapping sheet_name -> count of automatically changed rows
        """
        auto_changes: Dict[str, int] = {}
        
        if not hasattr(workbook_dataset, "sheets"):
            return auto_changes

        for sheet in workbook_dataset.sheets:
            sheet_name = sheet.name
            manual_row_uids = manual_edits_by_sheet.get(sheet_name, set())
            auto_count = 0

            for row in sheet.rows:
                row_uid = row.get("_row_uid", "")
                
                # Skip manually edited rows
                if row_uid in manual_row_uids:
                    continue
                
                # Heuristic: if row has any corrected fields that differ from source,
                # consider it auto-changed. This is imperfect but reasonable.
                has_corrections = self._row_has_corrections(row)
                if has_corrections:
                    auto_count += 1

            auto_changes[sheet_name] = auto_count

        return auto_changes

    def _row_has_corrections(self, row: dict) -> bool:
        """Check if a row has any corrected/standardized fields that differ from source.
        
        Heuristic: look for fields that have a _corrected variant or look for
        standardization status markers.
        """
        # Look for common corrected field patterns
        corrected_patterns = [
            "first_name_corrected",
            "last_name_corrected",
            "father_name_corrected",
            "birth_date_corrected",
            "gender_corrected",
            "identifier_corrected",
            "date_entry_corrected",
        ]
        
        for pattern in corrected_patterns:
            if pattern in row and row[pattern]:
                # Found a non-empty corrected field
                source_pattern = pattern.replace("_corrected", "")
                source_value = row.get(source_pattern)
                corrected_value = row.get(pattern)
                
                # If corrected differs from source, row was changed
                if source_value != corrected_value:
                    return True
        
        # Alternative: check for validation status that indicates changes
        if "_validation_status" in row:
            status = row.get("_validation_status", "")
            if status and status.lower() not in ["ok", "", "none"]:
                return True
        
        return False

    def _update_per_sheet_reports(
        self,
        report: ProcessingReport,
        workbook_dataset: WorkbookDataset,
        manual_edits_by_sheet: Dict[str, Set[str]],
        auto_changes_by_sheet: Dict[str, int],
    ) -> None:
        """Update per-sheet reports with enriched metrics."""
        
        if not hasattr(workbook_dataset, "sheets"):
            return

        per_sheet_map = {p.sheet_name: p for p in report.per_sheet_warnings}

        for sheet in workbook_dataset.sheets:
            sheet_name = sheet.name
            
            # Get or create per-sheet report
            if sheet_name not in per_sheet_map:
                per_sheet = PerSheetProcessingReport(sheet_name=sheet_name)
                report.per_sheet_warnings.append(per_sheet)
            else:
                per_sheet = per_sheet_map[sheet_name]

            # Update counts
            per_sheet.total_rows = len(sheet.rows)
            per_sheet.rows_processed = len(sheet.rows)
            per_sheet.rows_exported = len(sheet.rows)  # TODO: refine with actual export logic
            per_sheet.rows_changed_manually = len(manual_edits_by_sheet.get(sheet_name, set()))
            per_sheet.rows_changed_automatically = auto_changes_by_sheet.get(sheet_name, 0)
            
            # Determine sheet status
            per_sheet.sheet_status = self._determine_sheet_status(per_sheet)

    def _determine_sheet_status(self, per_sheet: PerSheetProcessingReport) -> str:
        """Determine the status string for a sheet.
        
        Returns:
            "בוצע" (completed), "בוצע עם אזהרות" (completed with warnings),
            "נכשל" (failed), or "unknown"
        """
        if per_sheet.errors and len(per_sheet.errors) > 0:
            return "נכשל"
        elif per_sheet.warnings and len(per_sheet.warnings) > 0:
            return "בוצע עם אזהרות"
        elif per_sheet.total_rows > 0:
            return "בוצע"
        else:
            return "unknown"

    def _calculate_summary(self, report: ProcessingReport) -> ReportSummary:
        """Calculate overall summary statistics from the full report."""
        summary = ReportSummary()

        # Count totals from per-sheet reports
        for per_sheet in report.per_sheet_warnings:
            summary.total_rows += per_sheet.total_rows
            summary.rows_processed += per_sheet.rows_processed
            summary.rows_exported += per_sheet.rows_exported
            summary.rows_changed_automatically += per_sheet.rows_changed_automatically
            summary.rows_changed_manually += per_sheet.rows_changed_manually
            
            # Count sheet statuses
            if per_sheet.sheet_status == "בוצע":
                summary.sheets_completed += 1
            elif per_sheet.sheet_status == "בוצע עם אזהרות":
                summary.sheets_completed_with_warnings += 1
                summary.sheets_with_warnings += 1
            elif per_sheet.sheet_status == "נכשל":
                summary.sheets_failed += 1

        summary.sheets_with_warnings += report.per_sheet_warnings.__len__() - summary.sheets_completed - summary.sheets_completed_with_warnings - summary.sheets_failed

        return summary
