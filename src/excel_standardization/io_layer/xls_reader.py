"""XlsReader — reads legacy .xls files and converts them to SheetDataset/WorkbookDataset.

This module is the ONLY place in the codebase that imports xlrd.
It is used exclusively for .xls files.  All .xlsx/.xlsm files continue
to use the existing openpyxl-based ExcelToJsonExtractor.

The conversion strategy:
1. Open the .xls file with xlrd.
2. For each sheet, build a minimal openpyxl-like worksheet shim so that
   the existing ExcelReader.detect_columns / detect_table_region logic
   can be reused without modification.
3. Extract rows using the detected column mapping.
4. Return SheetDataset / WorkbookDataset objects identical in structure
   to those produced by ExcelToJsonExtractor.

Hebrew error message for unreadable files:
    "לא ניתן לקרוא את קובץ ה־XLS. יש לוודא שהקובץ תקין ואינו מוגן בסיסמה."
"""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional

from ..data_types import (
    ColumnHeaderInfo,
    JsonRow,
    SheetDataset,
    WorkbookDataset,
)

logger = logging.getLogger(__name__)

XLS_ERROR_HE = (
    "לא ניתן לקרוא את קובץ ה־XLS. "
    "יש לוודא שהקובץ תקין ואינו מוגן בסיסמה."
)


# ---------------------------------------------------------------------------
# Minimal openpyxl-like shim so ExcelReader can work on .xls data
# ---------------------------------------------------------------------------

class _XlsCell:
    """Minimal cell shim that ExcelReader can read via .value."""

    __slots__ = ("value", "coordinate", "row", "column")

    def __init__(self, value: Any, row: int, col: int) -> None:
        self.value = value
        self.row = row
        self.column = col
        # openpyxl coordinate format e.g. "A1"
        self.coordinate = _col_letter(col) + str(row)


class _XlsWorksheet:
    """Minimal worksheet shim that ExcelReader can operate on.

    ExcelReader accesses:
        - worksheet.title
        - worksheet.max_row
        - worksheet.max_column
        - worksheet.cell(row, column) -> cell with .value
        - worksheet.merged_cells  (we expose an empty set)
        - iteration via worksheet.iter_rows() (not used by ExcelReader directly)
    """

    def __init__(self, title: str, data: List[List[Any]]) -> None:
        """
        Args:
            title: Sheet name.
            data:  2-D list [row_idx][col_idx] of Python values (0-based).
        """
        self.title = title
        self._data = data
        self.max_row = len(data)
        self.max_column = max((len(r) for r in data), default=0)
        # openpyxl exposes merged_cells as a MergedCellRange container;
        # we expose a plain set — ExcelReader only checks `coord in merged_cells`.
        self.merged_cells: set = set()

    def cell(self, row: int, column: int) -> _XlsCell:
        """Return a cell shim for the given 1-based row/column."""
        r = row - 1
        c = column - 1
        if 0 <= r < len(self._data) and 0 <= c < len(self._data[r]):
            value = self._data[r][c]
        else:
            value = None
        return _XlsCell(value, row, column)

    def iter_rows(self, min_row: int = 1, max_row: Optional[int] = None,
                  min_col: int = 1, max_col: Optional[int] = None,
                  values_only: bool = False):
        """Yield rows as tuples of cell shims (or values if values_only)."""
        end_row = (max_row or self.max_row)
        end_col = (max_col or self.max_column)
        for r in range(min_row, end_row + 1):
            row_cells = []
            for c in range(min_col, end_col + 1):
                cell = self.cell(r, c)
                row_cells.append(cell.value if values_only else cell)
            yield tuple(row_cells)


def _col_letter(col: int) -> str:
    """Convert 1-based column index to Excel letter(s), e.g. 1→'A', 27→'AA'."""
    result = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ---------------------------------------------------------------------------
# xlrd cell-type → Python value conversion
# ---------------------------------------------------------------------------

def _xlrd_cell_to_python(cell, datemode: int) -> Any:
    """Convert an xlrd Cell to a plain Python value.

    xlrd cell types:
        0 = XL_CELL_EMPTY
        1 = XL_CELL_TEXT
        2 = XL_CELL_NUMBER
        3 = XL_CELL_DATE
        4 = XL_CELL_BOOLEAN
        5 = XL_CELL_ERROR
        6 = XL_CELL_BLANK
    """
    import xlrd  # local import — only used for .xls files

    ctype = cell.ctype
    cvalue = cell.value

    if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
        return None
    if ctype == xlrd.XL_CELL_TEXT:
        return cvalue if cvalue != "" else None
    if ctype == xlrd.XL_CELL_NUMBER:
        # Return int when the float has no fractional part
        if cvalue == int(cvalue):
            return int(cvalue)
        return cvalue
    if ctype == xlrd.XL_CELL_DATE:
        try:
            dt = xlrd.xldate_as_datetime(cvalue, datemode)
            # Return date-only when time is midnight
            if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
                return dt.date()
            return dt
        except Exception:
            return cvalue
    if ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cvalue)
    if ctype == xlrd.XL_CELL_ERROR:
        return None
    return cvalue


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_xls_sheet_names(file_path: str) -> List[str]:
    """Return the list of sheet names from a .xls file.

    Raises:
        ValueError: with Hebrew message if the file cannot be opened.
    """
    try:
        import xlrd
        wb = xlrd.open_workbook(file_path, on_demand=True)
        names = wb.sheet_names()
        wb.release_resources()
        if not names:
            raise ValueError(XLS_ERROR_HE)
        return names
    except ValueError:
        raise
    except Exception as exc:
        logger.warning(f"xlrd cannot open '{file_path}': {exc}")
        raise ValueError(XLS_ERROR_HE) from exc


def extract_xls_to_workbook_dataset(file_path: str) -> WorkbookDataset:
    """Read a .xls file and return a WorkbookDataset.

    Uses the existing ExcelReader header-detection logic by wrapping each
    xlrd sheet in a lightweight openpyxl-compatible shim.

    Raises:
        ValueError: with Hebrew message if the file cannot be opened.
    """
    try:
        import xlrd
    except ImportError as exc:
        raise ImportError(
            "xlrd is required to read .xls files. "
            "Install it with: pip install xlrd>=2.0.1"
        ) from exc

    try:
        wb_xlrd = xlrd.open_workbook(file_path)
    except Exception as exc:
        logger.warning(f"xlrd cannot open '{file_path}': {exc}")
        raise ValueError(XLS_ERROR_HE) from exc

    datemode = wb_xlrd.datemode
    sheets: List[SheetDataset] = []
    skipped_sheets: List[str] = []

    for sheet_idx in range(wb_xlrd.nsheets):
        sheet_xlrd = wb_xlrd.sheet_by_index(sheet_idx)
        sheet_name = sheet_xlrd.name

        try:
            # Build 2-D list of Python values
            data: List[List[Any]] = []
            for r in range(sheet_xlrd.nrows):
                row_vals = []
                for c in range(sheet_xlrd.ncols):
                    cell = sheet_xlrd.cell(r, c)
                    row_vals.append(_xlrd_cell_to_python(cell, datemode))
                data.append(row_vals)

            # Wrap in the openpyxl-compatible shim
            ws_shim = _XlsWorksheet(title=sheet_name, data=data)

            # Reuse existing ExcelReader + ExcelToJsonExtractor logic
            from .excel_reader import ExcelReader
            from .excel_to_json_extractor import ExcelToJsonExtractor

            reader = ExcelReader()
            extractor = ExcelToJsonExtractor(
                excel_reader=reader,
                skip_empty_rows=False,
                handle_formulas=False,   # no formulas in .xls shim
                preserve_types=True,
            )

            dataset = extractor.extract_sheet_to_json(ws_shim)

            if dataset.get_metadata("skipped", False):
                skipped_sheets.append(sheet_name)
                logger.warning(
                    f"XLS sheet '{sheet_name}' skipped: "
                    f"{dataset.get_metadata('error', 'unknown')}"
                )
            else:
                sheets.append(dataset)
                logger.info(
                    f"XLS sheet '{sheet_name}' extracted: "
                    f"{len(dataset.rows)} rows, {len(dataset.field_names)} fields"
                )

        except Exception as exc:
            logger.error(
                f"Failed to extract XLS sheet '{sheet_name}': {exc}", exc_info=True
            )
            skipped_sheets.append(sheet_name)

    return WorkbookDataset(
        source_file=file_path,
        sheets=sheets,
        metadata={
            "total_sheets": wb_xlrd.nsheets,
            "processed_sheets": len(sheets),
            "skipped_sheets": skipped_sheets,
            "source_format": "xls",
        },
    )


def extract_xls_sheet_to_dataset(file_path: str, sheet_name: str) -> SheetDataset:
    """Read a single named sheet from a .xls file and return a SheetDataset.

    Raises:
        ValueError: with Hebrew message if the file cannot be opened.
        KeyError: if the sheet name does not exist.
    """
    try:
        import xlrd
    except ImportError as exc:
        raise ImportError(
            "xlrd is required to read .xls files. "
            "Install it with: pip install xlrd>=2.0.1"
        ) from exc

    try:
        wb_xlrd = xlrd.open_workbook(file_path)
    except Exception as exc:
        logger.warning(f"xlrd cannot open '{file_path}': {exc}")
        raise ValueError(XLS_ERROR_HE) from exc

    try:
        sheet_xlrd = wb_xlrd.sheet_by_name(sheet_name)
    except xlrd.biffh.XLRDError:
        raise KeyError(f"Sheet '{sheet_name}' not found in '{file_path}'")

    datemode = wb_xlrd.datemode

    data: List[List[Any]] = []
    for r in range(sheet_xlrd.nrows):
        row_vals = []
        for c in range(sheet_xlrd.ncols):
            cell = sheet_xlrd.cell(r, c)
            row_vals.append(_xlrd_cell_to_python(cell, datemode))
        data.append(row_vals)

    ws_shim = _XlsWorksheet(title=sheet_name, data=data)

    from .excel_reader import ExcelReader
    from .excel_to_json_extractor import ExcelToJsonExtractor

    reader = ExcelReader()
    extractor = ExcelToJsonExtractor(
        excel_reader=reader,
        skip_empty_rows=False,
        handle_formulas=False,
        preserve_types=True,
    )

    return extractor.extract_sheet_to_json(ws_shim)
