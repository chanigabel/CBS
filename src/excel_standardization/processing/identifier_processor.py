"""Identifier field processor for Israeli ID and passport fields."""

from typing import Optional
from openpyxl.worksheet.worksheet import Worksheet
from .field_processor import FieldProcessor
from ..io_layer.excel_reader import ExcelReader
from ..io_layer.excel_writer import ExcelWriter
from ..engines.identifier_engine import IdentifierEngine
from ..data_types import ColumnHeaderInfo


class IdentifierFieldProcessor(FieldProcessor):

    def __init__(self, reader: ExcelReader, writer: ExcelWriter, identifier_engine: IdentifierEngine):
        super().__init__(reader, writer)

        self.identifier_engine = identifier_engine

        self.id_header_info: Optional[ColumnHeaderInfo] = None
        self.passport_header_info: Optional[ColumnHeaderInfo] = None

        self.corrected_id_col: Optional[int] = None
        self.corrected_passport_col: Optional[int] = None
        self.corrected_status_col: Optional[int] = None

    def find_headers(self, worksheet: Worksheet) -> bool:
        # Reset per-sheet state
        self.id_header_info = None
        self.passport_header_info = None
        self.corrected_id_col = None
        self.corrected_passport_col = None
        self.corrected_status_col = None

        id_terms = ["מספר זהות", "תעודת זהות", "ת.ז"]
        passport_terms = ["מספר דרכון", "דרכון"]

        self.id_header_info = self.reader.find_header(worksheet, id_terms)
        self.passport_header_info = self.reader.find_header(worksheet, passport_terms)

        # VBA parity: ProcessIdentifiers exits unless BOTH headers are found.
        if self.id_header_info is None or self.passport_header_info is None:
            self.id_header_info = None
            self.passport_header_info = None
            return False

        return True

    def prepare_output_columns(self, worksheet: Worksheet) -> None:
        # VBA parity: output columns are inserted immediately after the passport column.
        if not self.id_header_info or not self.passport_header_info:
            return

        header_row = self.passport_header_info.header_row
        base_col = self.passport_header_info.col

        # VBA parity: insert three columns immediately to the right of the
        # anchor column ONLY if the first corrected header is not already
        # present at that position.
        existing = worksheet.cell(row=header_row, column=base_col + 1).value
        if isinstance(existing, str) and existing.strip() == "ת.ז. - מתוקן":
            # Idempotent: columns already exist, just record their indices.
            self.corrected_id_col = base_col + 1
            self.corrected_passport_col = base_col + 2
            self.corrected_status_col = base_col + 3
            return

        headers = ["ת.ז. - מתוקן", "דרכון - מתוקן", "סטטוס מזהה"]

        self.writer.insert_output_columns(
            worksheet,
            after_col=base_col,
            count=3,
            header_row=header_row,
            headers=headers,
        )

        # Defensive VBA-parity: ensure the three headers exist exactly where VBA writes them.
        # Some openpyxl edge cases around insertion/style propagation can leave blanks.
        for i, h in enumerate(headers, start=1):
            cell = worksheet.cell(row=header_row, column=base_col + i)
            if (cell.value or "").strip() != h:
                cell.value = h

        self.corrected_id_col = base_col + 1
        self.corrected_passport_col = base_col + 2
        self.corrected_status_col = base_col + 3

    def process_data(self, worksheet: Worksheet) -> None:

        if not self.id_header_info and not self.passport_header_info:
            return

        if (
            self.corrected_id_col is None
            or self.corrected_passport_col is None
            or self.corrected_status_col is None
        ):
            return

        # Use whichever header is available to determine row range
        ref_header = self.id_header_info or self.passport_header_info
        start_row = ref_header.header_row + 1

        max_last_row = max(
            self.id_header_info.last_row if self.id_header_info else 0,
            self.passport_header_info.last_row if self.passport_header_info else 0,
        )

        if max_last_row < start_row:
            return

        if self.id_header_info:
            id_values = self.reader.read_column_array(
                worksheet,
                self.id_header_info.col,
                start_row,
                max_last_row,
            )
        else:
            id_values = [None] * (max_last_row - start_row + 1)

        if self.passport_header_info:
            passport_values = self.reader.read_column_array(
                worksheet,
                self.passport_header_info.col,
                start_row,
                max_last_row,
            )
        else:
            passport_values = [None] * (max_last_row - start_row + 1)

        corrected_ids = []
        corrected_passports = []
        status_texts = []

        for i in range(len(id_values)):
            id_val = id_values[i] if i < len(id_values) else None
            passport_val = passport_values[i] if i < len(passport_values) else None

            result = self.identifier_engine.normalize_identifiers(id_val, passport_val)

            corrected_ids.append(result.corrected_id)
            corrected_passports.append(result.corrected_passport)
            status_texts.append(result.status_text)

        self.writer.write_column_array(
            worksheet,
            self.corrected_id_col,
            start_row,
            corrected_ids,
        )

        self.writer.write_column_array(
            worksheet,
            self.corrected_passport_col,
            start_row,
            corrected_passports,
        )

        self.writer.write_column_array(
            worksheet,
            self.corrected_status_col,
            start_row,
            status_texts,
        )

        # Highlight changed cells (VBA parity: pink where corrected != Trim(original))
        self.writer.highlight_changed_cells(
            worksheet, self.corrected_id_col, start_row, id_values, corrected_ids
        )
        self.writer.highlight_changed_cells(
            worksheet, self.corrected_passport_col, start_row, passport_values, corrected_passports
        )