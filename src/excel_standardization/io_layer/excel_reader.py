"""Excel reading operations for the standardization system.

This module provides the ExcelReader class which encapsulates all openpyxl
read operations. It isolates Excel I/O from business logic.
"""

import re
from typing import Any, List, Optional, Dict, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from ..data_types import ColumnHeaderInfo, TableRegion, DateGroup, DateFieldType


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

    # Words to ignore in headers
    IGNORE_KEYWORDS = ['מתוקן', 'corrected', 'fixed', 'updated']

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
        """Detect the table region in a worksheet by analyzing data patterns.

        Scans the first rows to find where the actual data table starts,
        accounting for titles, logos, and metadata above the table.

        Args:
            worksheet: The worksheet to analyze
            max_scan_rows: Maximum number of rows to scan for table detection

        Returns:
            TableRegion if table detected, None otherwise
        """
        # Check cache
        ws_id = id(worksheet)
        if ws_id in self._table_region_cache:
            return self._table_region_cache[ws_id]

        max_row = min(max_scan_rows, worksheet.max_row)
        max_col = worksheet.max_column

        # Score each row based on likelihood of being a header row
        row_scores = []
        for row_idx in range(1, max_row + 1):
            score = self._score_header_row(worksheet, row_idx, max_col)
            row_scores.append((row_idx, score))

        # Find the row with the highest score (most likely header row)
        if not row_scores:
            self._table_region_cache[ws_id] = None
            return None

        row_scores.sort(key=lambda x: x[1], reverse=True)
        best_row, best_score = row_scores[0]

        # If score is too low, no table detected
        if best_score < 3:
            self._table_region_cache[ws_id] = None
            return None

        # ------------------------------------------------------------------
        # Determine if we have multi-row headers (e.g., grouped date columns)
        header_rows = 1
        header_start_row = best_row

        # Check if the selected row is actually a sub-header row (has year/month/day keywords)
        is_subheader = False
        subheader_keywords = ['שנה', 'חודש', 'יום', 'year', 'month', 'day']
        for col_idx in range(1, min(max_col + 1, 50)):
            cell_value = worksheet.cell(row=best_row, column=col_idx).value
            if cell_value:
                normalized = self._normalize_text(str(cell_value))
                if any(kw in normalized for kw in subheader_keywords):
                    is_subheader = True
                    break

        # If selected row is a sub-header, check for parent header above
        if is_subheader and best_row > 1:
            parent_row = best_row - 1
            parent_score = self._score_header_row(worksheet, parent_row, max_col)
            if parent_score >= 2:
                header_rows = 2
                header_start_row = parent_row
            else:
                # Check for merged-cell parent with date keywords
                for col_idx in range(1, min(max_col + 1, 50)):
                    if self._is_merged_cell(worksheet, parent_row, col_idx):
                        mr = self._get_merged_cell_range(worksheet, parent_row, col_idx)
                        if mr and mr[3] > mr[2]:
                            pv = worksheet.cell(row=mr[0], column=mr[2]).value
                            if pv and any(kw in self._normalize_text(str(pv))
                                          for kw in ['תאריך', 'לידה', 'כניסה', 'date', 'birth', 'entry']):
                                header_rows = 2
                                header_start_row = parent_row
                                break

        # Check if there's a sub-header row below
        if best_row < max_row and header_rows == 1:
            next_row_score = self._score_subheader_row(worksheet, best_row + 1, max_col)
            if next_row_score >= 2:
                header_rows = 2

        # Find data start row (first row after headers with actual data)
        data_start_row = header_start_row + header_rows

        # Find table boundaries using header_start_row as the anchor.
        start_col, end_col = self._find_table_columns(worksheet, header_start_row, max_col)

        # When the layout has two header rows, expand column boundaries to cover both.
        if header_rows == 2:
            sub_start, sub_end = self._find_table_columns(worksheet, header_start_row + 1, max_col)
            start_col = min(start_col, sub_start)
            end_col = max(end_col, sub_end)

        # Skip any column-index reference rows immediately after the headers.
        if self._is_column_index_row(worksheet, data_start_row, start_col, end_col):
            data_start_row += 1

        end_row = self._find_table_end_row(worksheet, data_start_row, start_col, end_col)

        table_region = TableRegion(
            start_row=header_start_row,
            end_row=end_row,
            start_col=start_col,
            end_col=end_col,
            header_rows=header_rows,
            data_start_row=data_start_row,
        )

        self._table_region_cache[ws_id] = table_region
        return table_region

    def _score_header_row(self, worksheet: Worksheet, row_idx: int, max_col: int) -> int:
        """Score a row based on likelihood of being a header row.

        Args:
            worksheet: The worksheet
            row_idx: Row index to score
            max_col: Maximum column to check

        Returns:
            Score (higher = more likely to be header row)
        """
        score = 0
        non_empty_count = 0
        text_count = 0
        keyword_matches = 0

        for col_idx in range(1, min(max_col + 1, 50)):  # Check up to 50 columns
            # Handle merged cells - get value from top-left cell of merged range
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            
            # If cell is empty but part of a merged range, get value from merge origin
            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, row_idx, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, row_idx, col_idx)
                if merge_range:
                    # Get value from top-left cell of merged range
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value

            if cell_value is None or str(cell_value).strip() == "":
                continue

            non_empty_count += 1
            cell_text = str(cell_value).strip()

            # Check if it's text (not just numbers)
            if not cell_text.replace(".", "").replace(",", "").isdigit():
                text_count += 1

            # Check for keyword matches
            normalized_text = self._normalize_text(cell_text)
            if self._contains_field_keyword(normalized_text):
                keyword_matches += 1

        # Scoring logic:
        # - Multiple non-empty adjacent cells
        # - Mostly text (not numbers)
        # - Contains field keywords
        if non_empty_count >= 3:
            score += 2
        if non_empty_count >= 5:
            score += 1

        if text_count >= non_empty_count * 0.7:  # At least 70% text
            score += 2

        score += keyword_matches * 2  # Each keyword match adds 2 points

        return score

    def _score_subheader_row(self, worksheet: Worksheet, row_idx: int, max_col: int) -> int:
        """Score a row as a potential sub-header row (e.g., year/month/day).

        This method detects parent-child relationships by checking if cells in the
        row above are merged and span multiple columns, indicating a parent header.
        It also validates that parent headers have date-related keywords.

        Args:
            worksheet: The worksheet
            row_idx: Row index to score
            max_col: Maximum column to check

        Returns:
            Score (higher = more likely to be sub-header row)
        """
        score = 0
        subheader_keywords = ['שנה', 'חודש', 'יום', 'year', 'month', 'day']
        date_keywords = ['תאריך', 'לידה', 'כניסה', 'date', 'birth', 'entry']
        
        parent_row = row_idx - 1
        if parent_row < 1:
            return 0  # No parent row to check
        
        matched_subheaders = 0
        valid_parent_child_pairs = 0

        for col_idx in range(1, min(max_col + 1, 50)):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value

            # Handle merged cells - get value from top-left cell of merged range
            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, row_idx, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, row_idx, col_idx)
                if merge_range:
                    # Get value from top-left cell of merged range
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value

            if cell_value is None:
                continue

            normalized_text = self._normalize_text(str(cell_value))

            # Check if this cell matches subheader keywords
            has_subheader_keyword = any(keyword in normalized_text for keyword in subheader_keywords)

            if has_subheader_keyword:
                matched_subheaders += 1
                
                # Check if there's a parent header above this cell
                # Check if the cell above is merged and spans multiple columns
                if self._is_merged_cell(worksheet, parent_row, col_idx):
                    merge_range = self._get_merged_cell_range(worksheet, parent_row, col_idx)
                    if merge_range:
                        # Parent header spans multiple columns - strong indicator of parent-child relationship
                        start_col, end_col = merge_range[2], merge_range[3]
                        if end_col > start_col:
                            # Get parent header text
                            parent_cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value
                            if parent_cell_value:
                                parent_normalized = self._normalize_text(str(parent_cell_value))
                                # Verify parent has date-related keywords
                                if any(kw in parent_normalized for kw in date_keywords):
                                    # Strong parent-child relationship with date context
                                    valid_parent_child_pairs += 1
                                    score += 3
                                else:
                                    # Parent spans multiple columns but no date keyword
                                    score += 1
                            else:
                                score += 1
                        else:
                            # Parent is merged but doesn't span multiple columns
                            score += 1
                    else:
                        score += 1
                else:
                    # Check if parent cell (non-merged) has date-related keywords
                    parent_cell_value = worksheet.cell(row=parent_row, column=col_idx).value
                    if parent_cell_value:
                        parent_normalized = self._normalize_text(str(parent_cell_value))
                        if any(kw in parent_normalized for kw in date_keywords):
                            # Parent has date keyword, child has subheader - good match
                            valid_parent_child_pairs += 1
                            score += 2
                        else:
                            score += 1
                    else:
                        score += 1

        # Bonus for finding multiple valid parent-child pairs
        if valid_parent_child_pairs >= 2:
            score += 2
        
        # Bonus for finding all three date components (year, month, day)
        if matched_subheaders >= 3:
            score += 1

        return score


    def _is_column_index_row(self, worksheet: Worksheet, row_idx: int, start_col: int, end_col: int) -> bool:
        """Detect whether a row is a column-index reference row.

        Some Excel forms include a row of sequential or near-sequential small
        integers immediately after the header rows to label column positions for
        the form filler.  These rows must not be treated as data.

        A row is considered a column-index row when ALL of the following hold:
        1. Every non-null cell value is a positive integer.
        2. All values are small (≤ max column count, i.e., ≤ end_col).
        3. There are at least 3 distinct values.
        4. The values are all distinct (no duplicates).

        Note: gaps between values are allowed — real forms often skip column
        numbers for merged or empty columns.

        Args:
            worksheet: The worksheet
            row_idx: Row index to test
            start_col: First column of the table
            end_col: Last column of the table

        Returns:
            True if the row looks like a column-index reference row
        """
        values = []
        for col_idx in range(start_col, end_col + 1):
            cell_value = worksheet.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            # Must be an integer (or float that is a whole number).
            # Non-integer, non-numeric values (e.g. datetime objects, strings)
            # are treated as absent — they don't disqualify the row, because
            # some column-index rows have date/text values in non-indexed columns.
            if isinstance(cell_value, float):
                if cell_value != int(cell_value):
                    continue  # Non-whole float — treat as absent
                cell_value = int(cell_value)
            if not isinstance(cell_value, int):
                continue  # Non-integer — treat as absent
            # Must be a small positive integer within the column range
            if cell_value < 1 or cell_value > end_col:
                return False
            values.append(cell_value)

        # Need at least 3 values to be confident
        if len(values) < 3:
            return False

        # Values must be distinct (no duplicates)
        if len(values) != len(set(values)):
            return False

        return True

    def _find_table_columns(self, worksheet: Worksheet, header_row: int, max_col: int) -> Tuple[int, int]:
        """Find the start and end columns of the table.

        Args:
            worksheet: The worksheet
            header_row: The header row index
            max_col: Maximum column to check

        Returns:
            Tuple of (start_col, end_col)
        """
        start_col = 1
        end_col = max_col

        # Find first non-empty column
        for col_idx in range(1, max_col + 1):
            cell_value = worksheet.cell(row=header_row, column=col_idx).value
            
            # Handle merged cells
            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, header_row, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, header_row, col_idx)
                if merge_range:
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value
            
            if cell_value is not None and str(cell_value).strip() != "":
                start_col = col_idx
                break

        # Find last non-empty column
        for col_idx in range(max_col, 0, -1):
            cell_value = worksheet.cell(row=header_row, column=col_idx).value
            
            # Handle merged cells
            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, header_row, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, header_row, col_idx)
                if merge_range:
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value
            
            if cell_value is not None and str(cell_value).strip() != "":
                end_col = col_idx
                break

        return start_col, end_col

    def _find_table_end_row(
        self, worksheet: Worksheet, data_start_row: int, start_col: int, end_col: int
    ) -> int:
        """Find the last row of the table with data.

        Args:
            worksheet: The worksheet
            data_start_row: First row of data
            start_col: First column of table
            end_col: Last column of table

        Returns:
            Last row index with data
        """
        max_row = worksheet.max_row
        last_data_row = data_start_row

        # Scan from data start to find last row with data
        for row_idx in range(data_start_row, max_row + 1):
            has_data = False
            for col_idx in range(start_col, end_col + 1):
                cell_value = worksheet.cell(row=row_idx, column=col_idx).value
                
                # Handle merged cells
                if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, row_idx, col_idx):
                    merge_range = self._get_merged_cell_range(worksheet, row_idx, col_idx)
                    if merge_range:
                        cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value
                
                if cell_value is not None and str(cell_value).strip() != "":
                    has_data = True
                    break

            if has_data:
                last_data_row = row_idx
            elif row_idx > last_data_row + 5:  # Stop if 5 consecutive empty rows
                break

        return last_data_row

    def _normalize_text(self, text: str) -> str:
        """Normalize text for comparison.

        Removes line breaks, parentheses, converts to lowercase, collapses spaces.

        Args:
            text: Text to normalize

        Returns:
            Normalized text
        """
        # Remove line breaks
        text = text.replace("\n", " ").replace("\r", " ")

        # Remove parentheses and brackets
        text = re.sub(r'[()[\]{}]', ' ', text)

        # Convert to lowercase
        text = text.lower()

        # Collapse multiple spaces
        text = re.sub(r'\s+', ' ', text)

        # Trim
        text = text.strip()

        return text

    def _contains_field_keyword(self, normalized_text: str) -> bool:
        """Check if normalized text contains any field keyword.

        Args:
            normalized_text: Normalized text to check

        Returns:
            True if contains a field keyword
        """
        for field_keywords in self.FIELD_KEYWORDS.values():
            for keyword in field_keywords:
                if keyword in normalized_text:
                    return True
        return False

    def _should_ignore_column(self, cell_text: str) -> bool:
        """Check if a column should be ignored (e.g., already corrected columns).

        Args:
            cell_text: Cell text to check

        Returns:
            True if column should be ignored
        """
        normalized = self._normalize_text(cell_text)
        for ignore_word in self.IGNORE_KEYWORDS:
            if ignore_word in normalized:
                return True
        return False

    def _find_label_row(self, worksheet: Worksheet, col: int, header_area_rows: list) -> int:
        """Return the row in header_area_rows where col has its label (first non-empty, non-data cell)."""
        for hr in header_area_rows:
            v = worksheet.cell(row=hr, column=col).value
            if v is not None and str(v).strip() != "" and not self._looks_like_data_value(v):
                return hr
        return header_area_rows[0] if header_area_rows else 1

    def _looks_like_data_value(self, cell_value) -> bool:
        """Return True if a cell value looks like a data value rather than a header.

        Used by the sub-header pass to reject sample/example values that forms
        sometimes place in the sub-header row (e.g. '11.06.1997', datetime objects,
        plain numbers).  Real sub-headers are text strings containing Hebrew or
        English words — not dates or bare numbers.

        Args:
            cell_value: Raw cell value from openpyxl

        Returns:
            True if the value looks like data (should be skipped as a header)
        """
        from datetime import datetime as _dt, date as _date

        # datetime / date objects are always data values
        if isinstance(cell_value, (_dt, _date)):
            return True

        # Pure numbers are data values
        if isinstance(cell_value, (int, float)):
            return True

        if cell_value is None:
            return False

        txt = str(cell_value).strip()
        if not txt:
            return False

        # ISO datetime strings (e.g. "1997-09-04T00:00:00")
        import re as _re
        if _re.match(r'^\d{4}-\d{2}-\d{2}', txt):
            return True

        # Date strings with separators (e.g. "11.06.1997", "04/09/1997", "25/10/1899")
        if _re.match(r'^\d{1,2}[./]\d{1,2}[./]\d{2,4}$', txt):
            return True

        # Pure numeric strings (e.g. "12022001", "36872", "1234567")
        if txt.isdigit():
            return True

        # Numeric-like with only digits and separators (e.g. "1,234")
        stripped = txt.replace(',', '').replace('.', '').replace('-', '')
        if stripped.isdigit() and len(stripped) > 0:
            return True

        return False

    def _is_merged_cell(self, worksheet: Worksheet, row: int, col: int) -> bool:
        """Check if a cell is part of a merged range.

        Args:
            worksheet: The worksheet
            row: Row number (1-based)
            col: Column number (1-based)

        Returns:
            True if cell is part of a merged range, False otherwise
        """
        try:
            cell = worksheet.cell(row=row, column=col)
            # Check if this cell coordinate is in any merged cell range
            for merged_range in worksheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    return True
            return False
        except Exception:
            # If any error occurs, treat as not merged
            return False

    def _get_merged_cell_range(self, worksheet: Worksheet, row: int, col: int) -> Optional[Tuple[int, int, int, int]]:
        """Get the boundaries of a merged cell range.

        Args:
            worksheet: The worksheet
            row: Row number (1-based)
            col: Column number (1-based)

        Returns:
            Tuple of (start_row, end_row, start_col, end_col) if cell is merged, None otherwise
        """
        try:
            cell = worksheet.cell(row=row, column=col)
            # Find the merged range containing this cell
            for merged_range in worksheet.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Extract boundaries from the range
                    return (
                        merged_range.min_row,
                        merged_range.max_row,
                        merged_range.min_col,
                        merged_range.max_col,
                    )
            return None
        except Exception:
            # If any error occurs, return None
            return None

    def detect_columns(self, worksheet: Worksheet) -> Dict[str, ColumnHeaderInfo]:
        """Detect all relevant columns in the worksheet using intelligent table detection.

        This method:
        1. Detects the table region
        2. Identifies column headers using keyword matching
        3. Handles multi-row headers (e.g., date groups)
        4. Returns a mapping of field names to column information

        Args:
            worksheet: The worksheet to analyze

        Returns:
            Dictionary mapping field names to ColumnHeaderInfo
        """
        # Check cache
        ws_id = id(worksheet)
        if ws_id in self._column_mapping_cache:
            return self._column_mapping_cache[ws_id]

        # Detect table region
        table_region = self.detect_table_region(worksheet)
        if table_region is None:
            self._column_mapping_cache[ws_id] = {}
            return {}

        column_mapping: Dict[str, ColumnHeaderInfo] = {}
        processed_merged_cols = set()  # Track columns already processed as part of merged cells
        # Track every col_idx that was handled by the keyword-matching loop
        # (including date group parent headers that produce sub-columns rather
        # than a direct mapping entry).  These must be excluded from the
        # passthrough pass so they don't appear as duplicate raw columns.
        keyword_handled_cols: Set[int] = set()

        # Scan header row(s) for columns
        header_row = table_region.start_row
        subheader_row = header_row + 1 if table_region.header_rows == 2 else None

        # Deterministic date grouping (birth/entry)
        date_groups = self.detect_date_groups(worksheet, table_region)

        for col_idx in range(table_region.start_col, table_region.end_col + 1):
            # Skip if this column was already processed as part of a merged cell
            if col_idx in processed_merged_cols:
                continue

            # Get header cell text
            header_cell = worksheet.cell(row=header_row, column=col_idx)
            cell_value = header_cell.value

            # Handle merged cells - get value from top-left cell and mark all spanned columns
            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, header_row, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, header_row, col_idx)
                if merge_range:
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value
                    for merged_col in range(merge_range[2], merge_range[3] + 1):
                        processed_merged_cols.add(merged_col)
            elif self._is_merged_cell(worksheet, header_row, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, header_row, col_idx)
                if merge_range:
                    for merged_col in range(merge_range[2], merge_range[3] + 1):
                        processed_merged_cols.add(merged_col)

            if cell_value is None:
                continue

            header_text = str(cell_value).strip()

            if self._should_ignore_column(header_text):
                continue

            normalized_header = self._normalize_text(header_text)
            matched_field = self._match_field(normalized_header)

            if matched_field:
                keyword_handled_cols.add(col_idx)

                if matched_field in ["birth_date", "entry_date"] and subheader_row:
                    group_type = DateFieldType.BIRTH_DATE if matched_field == "birth_date" else DateFieldType.ENTRY_DATE
                    group = date_groups.get(group_type)
                    if group:
                        prefix = "birth" if matched_field == "birth_date" else "entry"
                        column_mapping[f"{prefix}_year"] = ColumnHeaderInfo(
                            col=group.year_col,
                            header_row=subheader_row,
                            last_row=table_region.end_row,
                            header_text=str(worksheet.cell(row=subheader_row, column=group.year_col).value or ""),
                        )
                        column_mapping[f"{prefix}_month"] = ColumnHeaderInfo(
                            col=group.month_col,
                            header_row=subheader_row,
                            last_row=table_region.end_row,
                            header_text=str(worksheet.cell(row=subheader_row, column=group.month_col).value or ""),
                        )
                        column_mapping[f"{prefix}_day"] = ColumnHeaderInfo(
                            col=group.day_col,
                            header_row=subheader_row,
                            last_row=table_region.end_row,
                            header_text=str(worksheet.cell(row=subheader_row, column=group.day_col).value or ""),
                        )
                    else:
                        column_mapping[matched_field] = ColumnHeaderInfo(
                            col=col_idx,
                            header_row=header_row,
                            last_row=table_region.end_row,
                            header_text=header_text,
                        )
                else:
                    column_mapping[matched_field] = ColumnHeaderInfo(
                        col=col_idx,
                        header_row=header_row,
                        last_row=table_region.end_row,
                        header_text=header_text,
                    )

        # ---------------------------------------------------------------
        # Passthrough pass: add every column whose header did NOT match a
        # keyword so that no Excel column is silently dropped.
        # ---------------------------------------------------------------
        already_mapped_cols: Set[int] = {info.col for info in column_mapping.values()}

        for col_idx in range(table_region.start_col, table_region.end_col + 1):
            if col_idx in already_mapped_cols or col_idx in processed_merged_cols or col_idx in keyword_handled_cols:
                continue

            cell_value = worksheet.cell(row=header_row, column=col_idx).value

            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, header_row, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, header_row, col_idx)
                if merge_range:
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value

            if cell_value is None or str(cell_value).strip() == "":
                continue

            header_text = str(cell_value).strip()

            if self._should_ignore_column(header_text):
                continue

            safe_key = re.sub(r'[^\w\u0590-\u05FF]+', '_', header_text).strip('_') or f"col_{col_idx}"
            if safe_key in column_mapping:
                safe_key = f"{safe_key}_{col_idx}"

            column_mapping[safe_key] = ColumnHeaderInfo(
                col=col_idx,
                header_row=header_row,
                last_row=table_region.end_row,
                header_text=header_text,
            )

        # ---------------------------------------------------------------
        # Sub-header pass (two-row header layout only):
        # When header_rows == 2, some regular fields (e.g. שם פרטי, שם משפחה,
        # שם האב) may live exclusively on the sub-header row while the top
        # header row has empty cells in those columns.
        #
        # ALL שנה/חודש/יום cells on the sub-header row are excluded to prevent
        # phantom `year`/`month`/`day` columns in the mapping.
        # ---------------------------------------------------------------
        if subheader_row is not None:
            already_mapped_cols = {info.col for info in column_mapping.values()}

            # Exclude all date-component sub-header columns
            _date_component_cols: Set[int] = set()
            for dg in date_groups.values():
                _date_component_cols.update([dg.year_col, dg.month_col, dg.day_col])
            _eff_end = table_region.end_col
            for _c in range(worksheet.max_column or 0, 0, -1):
                _v = worksheet.cell(row=subheader_row, column=_c).value
                if _v is not None and str(_v).strip() != "":
                    _eff_end = max(_eff_end, _c)
                    break
            for _c in range(table_region.start_col, _eff_end + 1):
                _v = worksheet.cell(row=subheader_row, column=_c).value
                if _v is not None and str(_v).strip() in ("שנה", "חודש", "יום"):
                    _date_component_cols.add(_c)

            for col_idx in range(table_region.start_col, table_region.end_col + 1):
                if col_idx in already_mapped_cols or col_idx in _date_component_cols:
                    continue

                cell_value = worksheet.cell(row=subheader_row, column=col_idx).value

                if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, subheader_row, col_idx):
                    merge_range = self._get_merged_cell_range(worksheet, subheader_row, col_idx)
                    if merge_range:
                        cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value

                if cell_value is None or str(cell_value).strip() == "":
                    continue

                header_text = str(cell_value).strip()

                if self._should_ignore_column(header_text):
                    continue

                # Skip cells that look like data values rather than headers.
                if self._looks_like_data_value(cell_value):
                    continue

                normalized_header = self._normalize_text(header_text)
                matched_field = self._match_field(normalized_header)

                if matched_field:
                    if matched_field not in column_mapping:
                        column_mapping[matched_field] = ColumnHeaderInfo(
                            col=col_idx,
                            header_row=subheader_row,
                            last_row=table_region.end_row,
                            header_text=header_text,
                        )
                else:
                    safe_key = re.sub(r'[^\w\u0590-\u05FF]+', '_', header_text).strip('_') or f"col_{col_idx}"
                    if safe_key in column_mapping:
                        safe_key = f"{safe_key}_{col_idx}"
                    column_mapping[safe_key] = ColumnHeaderInfo(
                        col=col_idx,
                        header_row=subheader_row,
                        last_row=table_region.end_row,
                        header_text=header_text,
                    )

        # Sort the final mapping by physical Excel column number so that
        # field_names (built from list(column_mapping.keys())) reflects the
        # true left-to-right worksheet column order regardless of which pass
        # (keyword, passthrough, sub-header) inserted each entry.
        sorted_mapping: Dict[str, ColumnHeaderInfo] = dict(
            sorted(column_mapping.items(), key=lambda kv: kv[1].col)
        )

        self._column_mapping_cache[ws_id] = sorted_mapping
        return sorted_mapping

    def detect_date_groups(self, worksheet: Worksheet, table_region: TableRegion) -> Dict[DateFieldType, DateGroup]:
        """Detect split date groups for birth/entry independently per field.

        For each date field (birth_date, entry_date), this method:
        1. Finds the parent header column by scanning all header rows
        2. Scans subsequent header rows for שנה/חודש/יום sub-headers
        3. Returns a DateGroup only if all three components were found

        Each field is detected independently — one field's shape does not
        influence the other's detection.
        """
        groups: Dict[DateFieldType, DateGroup] = {}

        if table_region.header_rows < 2:
            return groups

        header_row = table_region.start_row
        subheader_row = header_row + 1

        # Expand effective_end_col to cover all sub-header content
        effective_end_col = table_region.end_col
        for c in range(worksheet.max_column or 0, 0, -1):
            v = worksheet.cell(row=subheader_row, column=c).value
            if v is not None and str(v).strip() != "":
                effective_end_col = max(effective_end_col, c)
                break

        # Find parent header columns for birth and entry on header_row only
        def find_parent_col(keyword_list: List[str]) -> Optional[int]:
            for c in range(table_region.start_col, effective_end_col + 1):
                v = worksheet.cell(row=header_row, column=c).value
                if v is None or str(v).strip() == "":
                    continue
                norm = self._normalize_text(str(v).strip())
                if any(self._normalize_text(k) in norm for k in keyword_list):
                    return c
            return None

        birth_parent_col = find_parent_col(self.FIELD_KEYWORDS.get("birth_date", []))
        entry_parent_col = find_parent_col(self.FIELD_KEYWORDS.get("entry_date", []))

        for field_type, parent_col, other_parent_col in [
            (DateFieldType.BIRTH_DATE, birth_parent_col, entry_parent_col),
            (DateFieldType.ENTRY_DATE, entry_parent_col, birth_parent_col),
        ]:
            if parent_col is None:
                continue

            # Hard stop for the primary scan (parent_col+1 to hard_stop_primary):
            # stop BEFORE the other field's parent column to prevent bleeding.
            # Hard stop for the fallback scan (parent_col to hard_stop_fallback):
            # include the other field's parent column to handle the case where
            # birth's יום sits at the same column as entry's parent header.
            if other_parent_col is not None and other_parent_col > parent_col:
                hard_stop_primary = other_parent_col - 1
                hard_stop_fallback = other_parent_col
            else:
                hard_stop_primary = effective_end_col
                hard_stop_fallback = effective_end_col

            # Scan subheader_row from parent_col+1 to hard_stop_primary for שנה/חודש/יום
            year_col = month_col = day_col = 0
            found_count = 0
            for c in range(parent_col + 1, hard_stop_primary + 1):
                v = worksheet.cell(row=subheader_row, column=c).value
                if v is None or str(v).strip() == "":
                    if found_count > 0 and c < parent_col + 10:
                        continue
                    elif found_count > 0:
                        break
                    else:
                        continue
                txt = str(v).strip()
                if txt == "שנה" and year_col == 0:
                    year_col = c; found_count += 1
                elif txt == "חודש" and month_col == 0:
                    month_col = c; found_count += 1
                elif txt == "יום" and day_col == 0:
                    day_col = c; found_count += 1
                elif found_count > 0:
                    break
                if found_count == 3:
                    break

            # If not found starting at parent_col+1, try including parent_col itself
            # (handles merged-cell layouts where the first sub-header is at parent_col,
            # and also the Case B overlap where birth's יום is at entry's parent column).
            if not (year_col and month_col and day_col):
                year_col = month_col = day_col = 0
                for c in range(parent_col, hard_stop_fallback + 1):
                    v = worksheet.cell(row=subheader_row, column=c).value
                    if v is None or str(v).strip() == "":
                        continue
                    txt = str(v).strip()
                    if txt == "שנה" and year_col == 0:
                        year_col = c
                    elif txt == "חודש" and month_col == 0:
                        month_col = c
                    elif txt == "יום" and day_col == 0:
                        day_col = c
                    if year_col and month_col and day_col:
                        break

            if year_col and month_col and day_col:
                groups[field_type] = DateGroup(
                    year_col=year_col,
                    month_col=month_col,
                    day_col=day_col,
                    main_col=parent_col,
                    field_type=field_type,
                )

        return groups

    def _match_field(self, normalized_text: str) -> Optional[str]:
        """Match normalized text to a field name.

        Args:
            normalized_text: Normalized header text

        Returns:
            Field name if matched, None otherwise
        """
        # Sort keywords by length (longest first) to prioritize more specific matches
        best_match = None
        best_match_length = 0
        
        for field_name, keywords in self.FIELD_KEYWORDS.items():
            for keyword in keywords:
                if keyword in normalized_text and len(keyword) > best_match_length:
                    best_match = field_name
                    best_match_length = len(keyword)
        
        return best_match

    def _detect_date_subcolumns(
        self, worksheet: Worksheet, start_col: int, subheader_row: int, max_col: int
    ) -> Dict[str, int]:
        """Detect year/month/day sub-columns for a date group.

        This method checks for merged parent headers that span multiple child columns.
        It detects split date fields even when year/month/day appear as child columns
        under a parent date header. It intelligently determines the search range based
        on parent header boundaries.

        Args:
            worksheet: The worksheet
            start_col: Starting column of the date group
            subheader_row: Row containing sub-headers
            max_col: Maximum column to check

        Returns:
            Dictionary mapping 'year', 'month', 'day' to column indices
        """
        date_columns = {}

        # Determine search range based on parent header structure
        parent_row = subheader_row - 1
        search_start_col = start_col
        search_end_col = min(start_col + 5, max_col + 1)

        if parent_row >= 1:
            # Check if start_col is part of a merged parent header
            if self._is_merged_cell(worksheet, parent_row, start_col):
                merge_range = self._get_merged_cell_range(worksheet, parent_row, start_col)
                if merge_range:
                    # Parent spans from merge_range[2] to merge_range[3]
                    search_start_col = merge_range[2]
                    search_end_col = min(merge_range[3] + 1, max_col + 1)

                    # Verify parent has date-related keywords
                    parent_cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value
                    if parent_cell_value:
                        parent_normalized = self._normalize_text(str(parent_cell_value))
                        date_keywords = ['תאריך', 'לידה', 'כניסה', 'date', 'birth', 'entry']
                        if not any(kw in parent_normalized for kw in date_keywords):
                            # Parent doesn't have date keywords, fall back to original range
                            search_start_col = start_col
                            search_end_col = min(start_col + 5, max_col + 1)
            else:
                # Parent is not merged, check if it has date keywords
                parent_cell_value = worksheet.cell(row=parent_row, column=start_col).value
                if parent_cell_value:
                    parent_normalized = self._normalize_text(str(parent_cell_value))
                    date_keywords = ['תאריך', 'לידה', 'כניסה', 'date', 'birth', 'entry']
                    if any(kw in parent_normalized for kw in date_keywords):
                        # Parent has date keywords, expand search range slightly
                        search_end_col = min(start_col + 5, max_col + 1)

        # Check columns in the determined range for year/month/day
        for col_idx in range(search_start_col, search_end_col):
            cell_value = worksheet.cell(row=subheader_row, column=col_idx).value

            # Handle merged cells in subheader row
            if (cell_value is None or str(cell_value).strip() == "") and self._is_merged_cell(worksheet, subheader_row, col_idx):
                merge_range = self._get_merged_cell_range(worksheet, subheader_row, col_idx)
                if merge_range:
                    cell_value = worksheet.cell(row=merge_range[0], column=merge_range[2]).value

            if cell_value is None:
                continue

            normalized = self._normalize_text(str(cell_value))

            if any(kw in normalized for kw in ['שנה', 'year']):
                if 'year' not in date_columns:
                    date_columns['year'] = col_idx
            elif any(kw in normalized for kw in ['חודש', 'month']):
                if 'month' not in date_columns:
                    date_columns['month'] = col_idx
            elif any(kw in normalized for kw in ['יום', 'day']):
                if 'day' not in date_columns:
                    date_columns['day'] = col_idx

        # Only return if we found all three
        if len(date_columns) == 3:
            return date_columns
        return {}


    def find_header(
        self, worksheet: Worksheet, search_terms: List[str], normalize_linebreaks: bool = False
    ) -> Optional[ColumnHeaderInfo]:
        """Find column by exact text matching (xlPart equivalent).

        This replicates the VBA ExcelReader.FindHeader behavior:
        - Uses xlPart semantics (substring match anywhere in the cell)
        - Scans the worksheet directly without table-region heuristics

        Args:
            worksheet: The worksheet to search
            search_terms: List of header variants to search for
            normalize_linebreaks: If True, normalize line break characters
                before matching (converts \\r\\n, \\r, \\n to \\n)

        Returns:
            ColumnHeaderInfo if header found, None otherwise
        """
        # Direct scan (VBA-style Find with xlPart)
        for row_idx in range(1, worksheet.max_row + 1):
            for col_idx in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=row_idx, column=col_idx)
                cell_value = cell.value

                if cell_value is None:
                    continue

                # Convert to string for comparison
                cell_text = str(cell_value)

                # Normalize line breaks if requested
                if normalize_linebreaks:
                    cell_text = cell_text.replace("\r\n", "\n").replace("\r", "\n")

                # Check if any search term is contained in the cell text (xlPart)
                for search_term in search_terms:
                    search_text = search_term
                    if normalize_linebreaks:
                        search_text = search_text.replace("\\r\\n", "\n").replace("\\r", "\n").replace("\\n", "\n")

                    if search_text in cell_text:
                        # Skip already-corrected columns (contain "מתוקן" or "corrected")
                        if "מתוקן" in cell_text or "corrected" in cell_text.lower():
                            continue

                        # Found the header, now find the last row with data
                        last_row = self.get_last_row(worksheet, col_idx)

                        return ColumnHeaderInfo(
                            col=col_idx, header_row=row_idx, last_row=last_row, header_text=cell_text
                        )

        return None

    def read_column_array(self, worksheet: Worksheet, col: int, start_row: int, end_row: int) -> List[Any]:
        """Read column data as array.

        Reads a range of cells from a single column and returns them as a list.
        This enables array-based processing instead of cell-by-cell operations.

        Args:
            worksheet: The worksheet to read from
            col: Column number (1-based)
            start_row: Starting row number (1-based, inclusive)
            end_row: Ending row number (1-based, inclusive)

        Returns:
            List of cell values from the specified range
        """
        values = []
        for row_idx in range(start_row, end_row + 1):
            cell = worksheet.cell(row=row_idx, column=col)
            values.append(cell.value)
        return values

    def read_cell_value(self, worksheet: Worksheet, row: int, col: int) -> Any:
        """Read single cell value.

        Args:
            worksheet: The worksheet to read from
            row: Row number (1-based)
            col: Column number (1-based)

        Returns:
            The cell value (can be str, int, float, datetime, None, etc.)
        """
        return worksheet.cell(row=row, column=col).value

    def get_last_row(self, worksheet: Worksheet, col: int) -> int:
        """Find last non-empty row in column.

        Searches from the bottom of the worksheet upward to find the last
        row that contains a non-empty value in the specified column.

        Args:
            worksheet: The worksheet to search
            col: Column number (1-based)

        Returns:
            Row number of last non-empty cell (1-based), or 0 if column is empty
        """
        # Start from the worksheet's max_row and search upward
        for row_idx in range(worksheet.max_row, 0, -1):
            cell = worksheet.cell(row=row_idx, column=col)
            if cell.value is not None and str(cell.value).strip() != "":
                return row_idx

        return 0
