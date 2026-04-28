"""Name field processor for first name, last name, and father's name.

This module provides the NameFieldProcessor class that processes name fields
in Excel worksheets. It handles first name, last name, and father's name
columns, applying text standardization and father name pattern detection.
"""

from typing import List, Optional, Dict
from openpyxl.worksheet.worksheet import Worksheet
from .field_processor import FieldProcessor
from ..io_layer.excel_reader import ExcelReader
from ..io_layer.excel_writer import ExcelWriter
from ..engines.name_engine import NameEngine
from ..data_types import ColumnHeaderInfo, FatherNamePattern


class NameFieldProcessor(FieldProcessor):
    """Process first name, last name, and father's name fields."""

    def __init__(self, reader: ExcelReader, writer: ExcelWriter, name_engine: NameEngine):
        super().__init__(reader, writer)
        self.name_engine = name_engine
        self.first_name_info: Optional[ColumnHeaderInfo] = None
        self.last_name_info: Optional[ColumnHeaderInfo] = None
        self.father_name_info: Optional[ColumnHeaderInfo] = None
        self.corrected_columns: Dict[str, int] = {}

    def _reset_state(self) -> None:
        """Reset per-sheet state so the processor can be reused across sheets."""
        self.first_name_info = None
        self.last_name_info = None
        self.father_name_info = None
        self.corrected_columns = {}

    def find_headers(self, worksheet: Worksheet) -> bool:
        """Find שם פרטי, שם משפחה, שם האב headers."""
        self._reset_state()

        first_name_terms = ["שם פרטי", "first name", "firstname"]
        self.first_name_info = self.reader.find_header(worksheet, first_name_terms)

        last_name_terms = ["שם משפחה", "last name", "lastname", "family name"]
        self.last_name_info = self.reader.find_header(worksheet, last_name_terms)

        father_name_terms = ["שם האב", "שם אב", "father's name", "father name"]
        self.father_name_info = self.reader.find_header(worksheet, father_name_terms)

        return any([self.first_name_info, self.last_name_info, self.father_name_info])

    def prepare_output_columns(self, worksheet: Worksheet) -> None:
        """Insert corrected columns with ' - מתוקן' suffix.

        Processes left-to-right (first → last → father) so that each insertion
        shifts the columns to the right of it, and we update tracked positions
        accordingly after each insertion.
        """
        # Prepare output column for first name
        if self.first_name_info:
            corrected_header = self.first_name_info.header_text + " - מתוקן"
            corrected_col = self.writer.prepare_output_column(
                worksheet, self.first_name_info.col, corrected_header, self.first_name_info.header_row
            )
            self.corrected_columns["first_name"] = corrected_col

            # Shift columns that come after the inserted column
            if self.last_name_info and self.last_name_info.col > self.first_name_info.col:
                self.last_name_info.col += 1
            if self.father_name_info and self.father_name_info.col > self.first_name_info.col:
                self.father_name_info.col += 1

        # Prepare output column for last name
        if self.last_name_info:
            corrected_header = self.last_name_info.header_text + " - מתוקן"
            corrected_col = self.writer.prepare_output_column(
                worksheet, self.last_name_info.col, corrected_header, self.last_name_info.header_row
            )
            self.corrected_columns["last_name"] = corrected_col

            # Shift father name if it comes after the inserted column
            if self.father_name_info and self.father_name_info.col > self.last_name_info.col:
                self.father_name_info.col += 1

        # Prepare output column for father's name
        if self.father_name_info:
            corrected_header = self.father_name_info.header_text + " - מתוקן"
            corrected_col = self.writer.prepare_output_column(
                worksheet, self.father_name_info.col, corrected_header, self.father_name_info.header_row
            )
            self.corrected_columns["father_name"] = corrected_col

    def detect_father_name_pattern(self, father_names: List[str], last_names: List[str]) -> FatherNamePattern:
        """Detect if last name should be removed from father name."""
        sample_size = min(5, len(father_names), len(last_names))

        if sample_size == 0:
            return FatherNamePattern.NONE

        contains_count = 0
        first_position_count = 0
        last_position_count = 0

        for i in range(sample_size):
            father_name = father_names[i]
            last_name = last_names[i]

            if not father_name or not last_name:
                continue

            if last_name in father_name:
                contains_count += 1
                words = father_name.split()

                if len(words) > 0:
                    if words[0] == last_name or last_name in words[0]:
                        first_position_count += 1
                    if words[-1] == last_name or last_name in words[-1]:
                        last_position_count += 1

        if contains_count < 3:
            return FatherNamePattern.NONE

        if first_position_count >= 3:
            return FatherNamePattern.REMOVE_FIRST

        if last_position_count >= 3:
            return FatherNamePattern.REMOVE_LAST

        return FatherNamePattern.NONE

    def process_data(self, worksheet: Worksheet) -> None:
        """Normalize names using NameEngine and highlight changes."""
        if self.first_name_info:
            self._process_simple_name_field(
                worksheet, self.first_name_info, self.corrected_columns["first_name"], self.first_name_info.col
            )

        if self.last_name_info:
            self._process_simple_name_field(
                worksheet, self.last_name_info, self.corrected_columns["last_name"], self.last_name_info.col
            )

        if self.father_name_info:
            self._process_father_name_field(worksheet, self.father_name_info.col)

    def _process_simple_name_field(
        self, worksheet: Worksheet, header_info: ColumnHeaderInfo, corrected_col: int, adjusted_original_col: int
    ) -> None:
        """Process a simple name field (first name or last name)."""
        start_row = header_info.header_row + 1
        end_row = header_info.last_row

        original_values = self.reader.read_column_array(worksheet, adjusted_original_col, start_row, end_row)

        corrected_values = []
        for value in original_values:
            if value is None:
                corrected_values.append("")
            else:
                normalized = self.name_engine.normalize_name(str(value))
                corrected_values.append(normalized)

        self.writer.write_column_array(worksheet, corrected_col, start_row, corrected_values)
        self.writer.highlight_changed_cells(worksheet, corrected_col, start_row, original_values, corrected_values)

    def _process_father_name_field(self, worksheet: Worksheet, adjusted_original_col: int) -> None:
        """Process father's name field with pattern detection."""
        if self.father_name_info is None:
            return

        start_row = self.father_name_info.header_row + 1
        end_row = self.father_name_info.last_row

        original_father_names = self.reader.read_column_array(worksheet, adjusted_original_col, start_row, end_row)

        normalized_father_names = []
        for value in original_father_names:
            if value is None:
                normalized_father_names.append("")
            else:
                normalized = self.name_engine.normalize_name(str(value))
                normalized_father_names.append(normalized)

        if self.last_name_info:
            # self.last_name_info.col already reflects any column shifts from prepare_output_columns
            last_name_start = self.last_name_info.header_row + 1
            last_name_end = max(self.last_name_info.last_row, end_row)

            original_last_names = self.reader.read_column_array(
                worksheet, self.last_name_info.col, last_name_start, last_name_end
            )

            normalized_last_names = []
            for value in original_last_names:
                if value is None:
                    normalized_last_names.append("")
                else:
                    normalized = self.name_engine.normalize_name(str(value))
                    normalized_last_names.append(normalized)

            pattern = self.detect_father_name_pattern(normalized_father_names, normalized_last_names)

            corrected_father_names = []
            for i, father_name in enumerate(normalized_father_names):
                if i < len(normalized_last_names):
                    last_name = normalized_last_names[i]
                    corrected = self.name_engine.remove_last_name_from_father(father_name, last_name, pattern)
                    corrected_father_names.append(corrected)
                else:
                    corrected_father_names.append(father_name)
        else:
            corrected_father_names = normalized_father_names

        corrected_col = self.corrected_columns["father_name"]
        self.writer.write_column_array(worksheet, corrected_col, start_row, corrected_father_names)
        self.writer.highlight_changed_cells(
            worksheet, corrected_col, start_row, original_father_names, corrected_father_names
        )
