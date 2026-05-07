"""Excel reading operations for the standardization system.

This module provides the ExcelReader class which encapsulates all openpyxl
read operations. It isolates Excel I/O from business logic.
"""

import re

from typing import Any, List, Optional, Dict, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from ..data_types import ColumnHeaderInfo, TableRegion, DateGroup, DateFieldType
from . import column_detection, field_matching, merged_cells, sheet_access, table_detection


class ExcelReader:
    """Handles reading data from Excel worksheets.

    This class encapsulates all openpyxl read operations, providing a clean
    interface for the processing layer. It includes intelligent table detection
    to handle complex Excel forms with variable header positions.
    """

    # Field keywords for intelligent detection (normalized)
    FIELD_KEYWORDS = {
        'first_name': ['שם פרטי', 'first name', 'firstname', 'שם', 'name','first'],
        'last_name': ['שם משפחה', 'last name', 'lastname', 'משפחה', 'surname', 'family name', 'last'],
        'father_name': ['שם האב', 'father name', 'fathername', 'אב', 'father'],
        'gender': ['מין', 'gender', 'sex', 'זכר', 'נקבה'],
        'id_number': ['מספר זהות', 'תעודת זהות', 'id number', 'id', 'ת.ז', 'תז','תעודת_זהות','זהות_תעודת','זהות תעודת'],
        'passport': ['דרכון', 'passport', 'מספר דרכון'],
        'birth_date': ['תאריך לידה', 'birth date', 'date of birth', 'לידה', 'dob'],
        'entry_date': ['תאריך כניסה', 'entry date', 'admission date', 'כניסה למוסד', 'כניסה'],
        'year': ['שנה', 'year', 'yr'],
        'month': ['חודש', 'month', 'mon'],
        'day': ['יום', 'day'],
    }

    # Words to ignore in headers — columns whose headers contain these words
    # are skipped during field detection.  This prevents corrected/status
    # columns from a previous processing run from being re-imported as source
    # fields when the output file is uploaded again.
    IGNORE_KEYWORDS = [
        'מתוקן', 'corrected', 'fixed', 'updated',
        # Status column names written by the standardization pipeline.
        # These are UI-only and must never be treated as source input fields.
        'identifier_status', 'gender_status',
        'birth_date_status', 'entry_date_status',
        'validationstatus', 'validation_status',
        'סטטוס מזהה', 'סטטוס תאריך',
    ]

    def __init__(self) -> None:
        """Initialize the ExcelReader with caching for table detection."""
        self._table_region_cache: Dict[int, Optional[TableRegion]] = {}
        self._column_mapping_cache: Dict[int, Dict[str, ColumnHeaderInfo]] = {}

    def invalidate_cache(self, worksheet: Worksheet) -> None:
        """Invalidate cached column mapping for a worksheet.

        Must be called after columns are inserted or deleted so that
        subsequent find_header calls re-scan the current worksheet state.

        Args:
            worksheet: The worksheet whose cache entry should be cleared
        """
        ws_id = id(worksheet)
        self._column_mapping_cache.pop(ws_id, None)
        self._table_region_cache.pop(ws_id, None)

    def detect_table_region(self, worksheet: Worksheet, max_scan_rows: int = 30) -> Optional[TableRegion]:
        return table_detection.detect_table_region(
            worksheet,
            max_scan_rows,
            self._table_region_cache,
            self._normalize_text,
            self._score_header_row,
            self._score_subheader_row,
            self._is_merged_cell,
            self._get_merged_cell_range,
        )

    def _score_header_row(self, worksheet: Worksheet, row_idx: int, max_col: int) -> int:
        return table_detection.score_header_row(
            worksheet,
            row_idx,
            max_col,
            self._normalize_text,
            self._contains_field_keyword,
            self._is_merged_cell,
            self._get_merged_cell_range,
        )

    def _score_subheader_row(self, worksheet: Worksheet, row_idx: int, max_col: int) -> int:
        return table_detection.score_subheader_row(
            worksheet,
            row_idx,
            max_col,
            self._normalize_text,
            self._is_merged_cell,
            self._get_merged_cell_range,
        )

    def _is_column_index_row(self, worksheet: Worksheet, row_idx: int, start_col: int, end_col: int) -> bool:
        return table_detection.is_column_index_row(worksheet, row_idx, start_col, end_col)

    def _find_table_columns(self, worksheet: Worksheet, header_row: int, max_col: int) -> Tuple[int, int]:
        return table_detection.find_table_columns(
            worksheet,
            header_row,
            max_col,
            self._is_merged_cell,
            self._get_merged_cell_range,
        )

    def _find_table_end_row(
        self, worksheet: Worksheet, data_start_row: int, start_col: int, end_col: int
    ) -> int:
        return table_detection.find_table_end_row(
            worksheet,
            data_start_row,
            start_col,
            end_col,
            self._is_merged_cell,
            self._get_merged_cell_range,
        )

    def _normalize_text(self, text: str) -> str:
        return field_matching.normalize_text(text)

    def _contains_field_keyword(self, normalized_text: str) -> bool:
        return field_matching.contains_field_keyword(normalized_text, self.FIELD_KEYWORDS)

    def _should_ignore_column(self, cell_text: str) -> bool:
        return field_matching.should_ignore_column(cell_text, self.IGNORE_KEYWORDS)

    def _find_label_row(self, worksheet: Worksheet, col: int, header_area_rows: list) -> int:
        return field_matching.find_label_row(
            worksheet,
            col,
            header_area_rows,
            self._looks_like_data_value,
        )

    def _looks_like_data_value(self, cell_value) -> bool:
        return field_matching.looks_like_data_value(cell_value)

    def _is_merged_cell(self, worksheet: Worksheet, row: int, col: int) -> bool:
        return merged_cells.is_merged_cell(worksheet, row, col)

    def _get_merged_cell_range(self, worksheet: Worksheet, row: int, col: int) -> Optional[Tuple[int, int, int, int]]:
        return merged_cells.get_merged_cell_range(worksheet, row, col)

    def detect_columns(self, worksheet: Worksheet) -> Dict[str, ColumnHeaderInfo]:
        return column_detection.detect_columns(self, worksheet)

    def detect_date_groups(self, worksheet: Worksheet, table_region: TableRegion) -> Dict[DateFieldType, DateGroup]:
        return column_detection.detect_date_groups(self, worksheet, table_region)

    def _match_field(self, normalized_text: str) -> Optional[str]:
        return column_detection.match_field(self, normalized_text)

    def _detect_date_subcolumns(
        self, worksheet: Worksheet, start_col: int, subheader_row: int, max_col: int
    ) -> Dict[str, int]:
        return column_detection.detect_date_subcolumns(self, worksheet, start_col, subheader_row, max_col)

    def find_header(
        self, worksheet: Worksheet, search_terms: List[str], normalize_linebreaks: bool = False
    ) -> Optional[ColumnHeaderInfo]:
        return sheet_access.find_header(self, worksheet, search_terms, normalize_linebreaks)

    def read_column_array(self, worksheet: Worksheet, col: int, start_row: int, end_row: int) -> List[Any]:
        return sheet_access.read_column_array(self, worksheet, col, start_row, end_row)

    def read_cell_value(self, worksheet: Worksheet, row: int, col: int) -> Any:
        return sheet_access.read_cell_value(self, worksheet, row, col)

    def get_last_row(self, worksheet: Worksheet, col: int) -> int:
        return sheet_access.get_last_row(self, worksheet, col)

