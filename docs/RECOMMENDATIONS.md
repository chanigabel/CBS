# RECOMMENDATIONS — Tests, Functionality, and UI/UX

> Based on a full scan of the codebase as it exists today.
> Coverage baseline: 80% overall (date_engine 66%, name_engine 65%, orchestrator 56%).
> Every recommendation is grounded in specific code locations.

---

## PART 1 — MOST IMPORTANT TESTS TO ADD

### 1.1 Critical missing web-path tests

These are the highest-value gaps: behaviors that exist in production code but have zero test coverage.

---

#### T-01 · Edit is lost after re-normalize

**Why it matters:** `record.edits` is populated but never replayed. This is the most dangerous silent data loss in the system. No test currently catches it.

**File to add to:** `tests/webapp/test_normalization_service.py`

```python
def test_manual_edit_is_lost_after_renormalize(session_with_file):
    """Regression: edits stored in record.edits are silently discarded on re-normalize."""
    svc, norm_svc = session_with_file
    from webapp.services.edit_service import EditService
    from webapp.models.requests import CellEditRequest

    # First normalize to get corrected fields
    norm_svc.normalize("test-session")

    # Manually edit a corrected field
    edit_svc = EditService(svc)
    record = svc.get("test-session")
    sheet = record.workbook_dataset.sheets[0]
    field = sheet.field_names[0]
    req = CellEditRequest(row_index=0, field_name=field, new_value="MANUAL_EDIT")
    edit_svc.edit_cell("test-session", sheet.sheet_name, req)

    # Verify edit is in memory
    record = svc.get("test-session")
    assert record.workbook_dataset.sheets[0].rows[0][field] == "MANUAL_EDIT"
    assert len(record.edits) == 1

    # Re-normalize — edit should be lost (this documents the current broken behavior)
    norm_svc.normalize("test-session")
    record = svc.get("test-session")
    # This assertion documents the bug: edit is gone
    assert record.workbook_dataset.sheets[0].rows[0][field] != "MANUAL_EDIT"
    # edits dict is still populated but was never replayed
    assert len(record.edits) == 1  # still there, never used
```

---

#### T-02 · Deleted rows return after re-normalize

**File to add to:** `tests/webapp/test_normalization_service.py`

```python
def test_deleted_rows_return_after_renormalize(session_with_file):
    """Regression: rows deleted in-memory are restored when re-normalize re-extracts from disk."""
    svc, norm_svc = session_with_file
    from webapp.services.edit_service import EditService
    from webapp.models.requests import DeleteRowRequest

    norm_svc.normalize("test-session")
    record = svc.get("test-session")
    original_count = len(record.workbook_dataset.sheets[0].rows)
    assert original_count >= 1

    edit_svc = EditService(svc)
    sheet_name = record.workbook_dataset.sheets[0].sheet_name
    req = DeleteRowRequest(row_indices=[0])
    edit_svc.delete_rows("test-session", sheet_name, req)

    record = svc.get("test-session")
    assert len(record.workbook_dataset.sheets[0].rows) == original_count - 1

    # Re-normalize restores the deleted row
    norm_svc.normalize("test-session")
    record = svc.get("test-session")
    assert len(record.workbook_dataset.sheets[0].rows) == original_count
```

---

#### T-03 · Export before normalization produces blank personal-data columns

**File to add to:** `tests/webapp/test_export_service.py`

```python
def test_export_before_normalization_produces_blank_corrected_columns(tmp_path):
    """Export without prior normalization: all *_corrected fields absent → blank cells."""
    from openpyxl import load_workbook as lw
    svc, _ = make_session_with_workbook()
    # Remove corrected fields to simulate pre-normalization state
    record = svc.get("export-session")
    for row in record.workbook_dataset.sheets[0].rows:
        row.pop("first_name_corrected", None)
        row.pop("last_name_corrected", None)

    export_svc = ExportService(svc, tmp_path / "output")
    output_path = export_svc.export("export-session")
    wb = lw(str(output_path))
    ws = wb.active
    # ShemPrati maps to first_name_corrected — should be blank
    # Row 2 is first data row (row 1 is header)
    assert ws.cell(row=2, column=3).value is None  # ShemPrati column
```

---

#### T-04 · Entry-before-birth NOT checked in web path (regression guard)

**File to add to:** `tests/webapp/test_normalization_service.py`

```python
def test_entry_before_birth_not_flagged_in_web_path(tmp_path):
    """Documents that entry-before-birth is NOT checked in the web pipeline.
    This test should FAIL once the feature is implemented, serving as a reminder.
    """
    import io
    from openpyxl import Workbook
    from webapp.services.upload_service import UploadService
    from webapp.services.normalization_service import NormalizationService
    from webapp.services.session_service import SessionService

    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Sheet1")
    ws.append(["first_name", "birth_year", "birth_month", "birth_day",
               "entry_year", "entry_month", "entry_day"])
    # Entry date (1990) before birth date (2000) — logically impossible
    ws.append(["Alice", 2000, 1, 1, 1990, 1, 1])
    buf = io.BytesIO(); wb.save(buf)

    svc = SessionService(); svc.clear_all()
    upload_svc = UploadService(svc, tmp_path / "u", tmp_path / "w")
    resp = upload_svc.handle_upload("test.xlsx", buf.getvalue())
    norm_svc = NormalizationService(svc)
    norm_svc.normalize(resp.session_id)

    record = svc.get(resp.session_id)
    row = record.workbook_dataset.sheets[0].rows[0]
    # Currently: no warning written. This documents the gap.
    assert row.get("entry_date_status", "") == ""  # no cross-validation warning
```

---

#### T-05 · DDMM hardcoded — US-format date silently wrong

**File to add to:** `tests/test_normalization_pipeline.py`

```python
def test_us_format_date_fails_silently_with_hardcoded_ddmm():
    """Documents that DDMM is hardcoded: US-format '01/15/1990' (Jan 15) is
    parsed as day=01, month=15 → invalid month, not as month=01, day=15."""
    pipeline = make_pipeline()
    row = {"birth_year": "01/15/1990", "birth_month": None, "birth_day": None}
    pipeline.apply_date_normalization(row)
    # With DDMM: day=01, month=15 → invalid
    assert row["birth_date_status"] != ""  # some error status
    # The correct parse (MMDD) would give month=1, day=15 — not achieved
    assert row.get("birth_month_corrected") != 1 or row.get("birth_day_corrected") != 15
```

---

#### T-06 · Whitespace-only gender inconsistency

**File to add to:** `tests/test_normalization_pipeline.py`

```python
def test_whitespace_gender_reaches_engine_unlike_empty_string():
    """Documents inconsistency: '' is preserved as-is but '   ' reaches engine → returns 1."""
    pipeline = make_pipeline()

    row_empty = {"gender": ""}
    pipeline.apply_gender_normalization(row_empty)
    assert row_empty["gender_corrected"] == ""  # preserved

    row_whitespace = {"gender": "   "}
    pipeline.apply_gender_normalization(row_whitespace)
    # Inconsistency: whitespace reaches engine, engine strips → returns 1 (male)
    assert row_whitespace["gender_corrected"] == 1  # not preserved as "   "
```

---

#### T-07 · Export mapping: corrected-only, no fallback

**File to add to:** `tests/webapp/test_export_service.py`

```python
def test_export_uses_corrected_field_only_no_fallback(tmp_path):
    """If first_name_corrected is absent, ShemPrati is blank — no fallback to first_name."""
    from openpyxl import load_workbook as lw
    svc = SessionService(); svc.clear_all()
    sheet = SheetDataset(
        sheet_name="דיירים יחידים",
        header_row=1, header_rows_count=1,
        field_names=["first_name"],
        rows=[{"first_name": "Alice"}],  # no first_name_corrected
    )
    wb = WorkbookDataset(source_file="t.xlsx", sheets=[sheet])
    from webapp.models.session import SessionRecord
    record = SessionRecord(session_id="s", source_file_path="", working_copy_path="",
                           original_filename="t.xlsx", status="uploaded", workbook_dataset=wb)
    svc.create(record)
    export_svc = ExportService(svc, tmp_path / "out")
    path = export_svc.export("s")
    out_wb = lw(str(path))
    ws = out_wb.active
    # ShemPrati (col 3) should be blank — no fallback
    assert ws.cell(row=2, column=3).value is None
```

---

### 1.2 Missing date_engine coverage (currently 66%)

The 101 missing statements are concentrated in `validate_entry_before_birth` and edge cases of `_parse_mixed_month_numeric`.

**File to add to:** `tests/test_date_engine.py`

```python
class TestValidateEntryBeforeBirth:
    def test_entry_before_birth_returns_false(self):
        from src.excel_normalization.data_types import DateParseResult
        engine = DateEngine()
        birth = DateParseResult(year=2000, month=1, day=1, is_valid=True, status_text="")
        entry = DateParseResult(year=1990, month=1, day=1, is_valid=True, status_text="")
        assert engine.validate_entry_before_birth(birth, entry) is False

    def test_entry_after_birth_returns_true(self):
        from src.excel_normalization.data_types import DateParseResult
        engine = DateEngine()
        birth = DateParseResult(year=1990, month=1, day=1, is_valid=True, status_text="")
        entry = DateParseResult(year=2010, month=6, day=15, is_valid=True, status_text="")
        assert engine.validate_entry_before_birth(birth, entry) is True

    def test_invalid_birth_skips_check(self):
        from src.excel_normalization.data_types import DateParseResult
        engine = DateEngine()
        birth = DateParseResult(year=None, month=None, day=None, is_valid=False, status_text="error")
        entry = DateParseResult(year=1990, month=1, day=1, is_valid=True, status_text="")
        assert engine.validate_entry_before_birth(birth, entry) is True  # skipped

    def test_year_only_birth_skips_check(self):
        from src.excel_normalization.data_types import DateParseResult
        engine = DateEngine()
        birth = DateParseResult(year=1990, month=0, day=0, is_valid=False, status_text="חסר חודש ויום")
        entry = DateParseResult(year=1985, month=1, day=1, is_valid=True, status_text="")
        assert engine.validate_entry_before_birth(birth, entry) is True

class TestExcelSerialDate:
    def test_serial_36526_is_year_2000(self):
        engine = DateEngine()
        from src.excel_normalization.data_types import DateFormatPattern, DateFieldType
        result = engine.parse_date(None, None, None, 36526, DateFormatPattern.DDMM, DateFieldType.BIRTH_DATE)
        assert result.year == 2000
        assert result.month == 1
        assert result.day == 1

    def test_serial_zero_is_invalid(self):
        engine = DateEngine()
        from src.excel_normalization.data_types import DateFormatPattern, DateFieldType
        result = engine.parse_date(None, None, None, 0, DateFormatPattern.DDMM, DateFieldType.BIRTH_DATE)
        assert result.is_valid is False
```

---

### 1.3 Missing name_engine coverage (currently 65%)

**File to add to:** `tests/test_name_engine.py`

```python
class TestDetectFirstNamePattern:
    def test_returns_none_with_fewer_than_3_matches(self):
        from src.excel_normalization.engines.text_processor import TextProcessor
        engine = NameEngine(TextProcessor())
        # Only 2 rows where last name appears in first name
        first = [["כהן יוסי"], ["כהן שרה"], ["דוד"]]
        last  = [["כהן"],      ["כהן"],      ["לוי"]]
        result = engine.detect_first_name_pattern(first, last)
        from src.excel_normalization.data_types import FatherNamePattern
        assert result == FatherNamePattern.NONE

    def test_returns_remove_first_with_3_matches_at_start(self):
        from src.excel_normalization.engines.text_processor import TextProcessor
        engine = NameEngine(TextProcessor())
        first = [["כהן יוסי"], ["כהן שרה"], ["כהן דוד"]]
        last  = [["כהן"],      ["כהן"],      ["כהן"]]
        result = engine.detect_first_name_pattern(first, last)
        from src.excel_normalization.data_types import FatherNamePattern
        assert result == FatherNamePattern.REMOVE_FIRST

class TestRemoveLastNameFromFirstName:
    def test_single_token_never_modified(self):
        from src.excel_normalization.engines.text_processor import TextProcessor
        engine = NameEngine(TextProcessor())
        result = engine.remove_last_name_from_first_name("כהן", "כהן")
        assert result == "כהן"

    def test_stage_a_removes_embedded_last_name(self):
        from src.excel_normalization.engines.text_processor import TextProcessor
        engine = NameEngine(TextProcessor())
        result = engine.remove_last_name_from_first_name("כהן יוסי", "כהן")
        assert result == "יוסי"

    def test_stage_b_positional_remove_first(self):
        from src.excel_normalization.engines.text_processor import TextProcessor
        from src.excel_normalization.data_types import FatherNamePattern
        engine = NameEngine(TextProcessor())
        # Last name not a substring — Stage B fires
        result = engine.remove_last_name_from_first_name(
            "לוי יוסי", "כהן", FatherNamePattern.REMOVE_FIRST
        )
        assert result == "יוסי"
```

---

### 1.4 WorkbookService display-shaping tests (currently untested edge cases)

**File to add to:** `tests/webapp/test_workbook_service.py`

```python
def test_empty_rows_filtered_from_display(session_with_workbook):
    """Rows where all original columns are None/empty are hidden."""
    _, wb_svc = session_with_workbook
    svc, _ = session_with_workbook
    record = svc.get("wb-session")
    record.workbook_dataset.sheets[0].rows.append(
        {"first_name": None, "last_name": None}
    )
    response = wb_svc.get_sheet_data("wb-session", "Sheet1")
    # The empty row should not appear
    assert all(
        r.get("first_name") is not None or r.get("last_name") is not None
        for r in response.rows
    )

def test_leading_numeric_helper_row_hidden(session_with_workbook):
    """First row that is all-numeric is hidden from display."""
    svc, wb_svc = session_with_workbook
    record = svc.get("wb-session")
    # Prepend a numeric helper row
    record.workbook_dataset.sheets[0].rows.insert(0, {"first_name": 1, "last_name": 2})
    response = wb_svc.get_sheet_data("wb-session", "Sheet1")
    # Helper row should be gone
    assert response.rows[0]["first_name"] != 1

def test_corrected_field_placed_after_original(session_with_workbook):
    """After normalization, first_name_corrected appears immediately after first_name."""
    svc, wb_svc = session_with_workbook
    record = svc.get("wb-session")
    for row in record.workbook_dataset.sheets[0].rows:
        row["first_name_corrected"] = row["first_name"] + "_c"
    response = wb_svc.get_sheet_data("wb-session", "Sheet1")
    cols = response.field_names
    fn_idx = cols.index("first_name")
    fnc_idx = cols.index("first_name_corrected")
    assert fnc_idx == fn_idx + 1

def test_row_with_only_corrected_values_is_filtered_out():
    """Row where source columns are all empty but corrected fields exist is hidden."""
    from src.excel_normalization.data_types import SheetDataset, WorkbookDataset
    from webapp.models.session import SessionRecord
    from webapp.services.session_service import SessionService
    from webapp.services.workbook_service import WorkbookService

    svc = SessionService(); svc.clear_all()
    sheet = SheetDataset(
        sheet_name="S", header_row=1, header_rows_count=1,
        field_names=["first_name"],
        rows=[{"first_name": None, "first_name_corrected": "יוסי"}],
    )
    wb = WorkbookDataset(source_file="t.xlsx", sheets=[sheet])
    record = SessionRecord(session_id="s", source_file_path="", working_copy_path="",
                           original_filename="t.xlsx", status="uploaded", workbook_dataset=wb)
    svc.create(record)
    response = WorkbookService(svc).get_sheet_data("s", "S")
    assert len(response.rows) == 0  # filtered because source is empty
```

---

### 1.5 derived_columns tests

**File to add to:** `tests/webapp/` (new file `test_derived_columns.py`)

```python
from webapp.services.derived_columns import apply_derived_columns, detect_serial_field, SYNTHETIC_SERIAL_KEY

def test_synthetic_serial_injected_when_no_source_column():
    rows = [{"first_name": "Alice"}, {"first_name": "Bob"}]
    rows, cols = apply_derived_columns(rows, ["first_name"], ["first_name"])
    assert rows[0][SYNTHETIC_SERIAL_KEY] == 1
    assert rows[1][SYNTHETIC_SERIAL_KEY] == 2
    assert cols[0] == SYNTHETIC_SERIAL_KEY

def test_blank_serial_cells_auto_filled():
    rows = [{"מספר_סידורי": None}, {"מספר_סידורי": ""}]
    rows, _ = apply_derived_columns(rows, ["מספר_סידורי"], ["מספר_סידורי"])
    assert rows[0]["מספר_סידורי"] == 1
    assert rows[1]["מספר_סידורי"] == 2

def test_mosad_id_injected_from_metadata():
    rows = [{"first_name": "Alice"}]
    rows, cols = apply_derived_columns(rows, ["first_name"], ["first_name"], meta_mosad_id="999")
    assert rows[0]["MosadID"] == "999"
    assert "MosadID" in cols

def test_mosad_id_not_shown_when_absent():
    rows = [{"first_name": "Alice"}]
    rows, cols = apply_derived_columns(rows, ["first_name"], ["first_name"], meta_mosad_id=None)
    assert "MosadID" not in cols

def test_id_number_not_mistaken_for_serial():
    result = detect_serial_field(["id_number", "first_name"])
    assert result is None
```


---

## PART 2 — MISSING OR WEAK FUNCTIONALITY TO IMPROVE

These are concrete, targeted fixes — not rewrites.

---

### F-01 · Replay edits after re-normalize

**Problem:** `record.edits` is populated on every `PATCH /cell` but `NormalizationService.normalize` re-extracts from disk and discards all in-memory changes.

**Where:** `webapp/services/normalization_service.py`, end of `normalize()`, after the merge loop.

**Exact fix — add ~8 lines:**

```python
# After: record.workbook_dataset.sheets = updated_sheets
# Add:
if record.edits:
    for (sheet_name, row_idx, field_name), value in record.edits.items():
        sheet = record.workbook_dataset.get_sheet_by_name(sheet_name)
        if sheet and 0 <= row_idx < len(sheet.rows):
            if field_name in sheet.rows[row_idx]:
                sheet.rows[row_idx][field_name] = value
```

**Risk:** Low. The edits dict already has the right structure `{(sheet, row, field): value}`. No new data model needed.

---

### F-02 · Entry-before-birth cross-validation

**Problem:** `DateEngine.validate_entry_before_birth` is fully implemented but never called by `NormalizationPipeline`.

**Where:** `src/excel_normalization/processing/normalization_pipeline.py`, `apply_date_normalization()`.

**Exact fix — add ~15 lines after both date fields are normalized:**

```python
# After both _normalize_date_field calls, add:
# Cross-validate entry vs birth
birth_result = getattr(self, "_last_birth_result", None)
entry_result = getattr(self, "_last_entry_result", None)
if birth_result and entry_result:
    if not self.date_engine.validate_entry_before_birth(birth_result, entry_result):
        existing = json_row.get("entry_date_status", "")
        warning = "תאריך כניסה לפני תאריך לידה"
        json_row["entry_date_status"] = f"{existing} | {warning}".strip(" |") if existing else warning
```

Also store results in `_normalize_date_field`:
```python
# After: result = self.date_engine.parse_date(...)
if prefix == "birth":
    self._last_birth_result = result
elif prefix == "entry":
    self._last_entry_result = result
```

**Risk:** Low. The engine method already exists and is tested. Only the call site is missing.

---

### F-03 · Date format auto-detection (DDMM vs MMDD)

**Problem:** `NormalizationPipeline._normalize_date_field` always passes `DateFormatPattern.DDMM`. A sheet with US-format dates (`01/15/1990`) will silently produce wrong results.

**Where:** `src/excel_normalization/processing/normalization_pipeline.py`, `normalize_dataset()`.

**Exact fix — detect pattern once per dataset, cache on pipeline:**

```python
# In normalize_dataset(), after pattern detection for names, add:
if self.apply_date_normalization_enabled and self.date_engine:
    # Sample first 10 non-null separated date values to detect format
    date_samples = []
    for row in corrected_dataset.rows[:20]:
        for prefix in ("birth", "entry"):
            for field in (f"{prefix}_date", f"{prefix}_year"):
                v = row.get(field)
                if v and isinstance(v, str) and ("/" in v or "." in v):
                    date_samples.append(v)
    self._date_format_pattern = _detect_date_format(date_samples)
```

Add helper function (can be a module-level function):
```python
def _detect_date_format(samples):
    from src.excel_normalization.data_types import DateFormatPattern
    ddmm = mmdd = 0
    for s in samples:
        parts = s.replace(".", "/").split("/")
        if len(parts) >= 2:
            try:
                a, b = int(parts[0]), int(parts[1])
                if a > 12 and b <= 12: ddmm += 1
                elif b > 12 and a <= 12: mmdd += 1
            except ValueError:
                pass
    return DateFormatPattern.MMDD if mmdd > ddmm else DateFormatPattern.DDMM
```

Then in `_normalize_date_field`, replace the hardcoded `DateFormatPattern.DDMM` with:
```python
pattern = getattr(self, "_date_format_pattern", DateFormatPattern.DDMM)
```

**Risk:** Low-medium. The detection logic already exists in `DateFieldProcessor.detect_date_format_pattern`. This is a port, not new logic.

---

### F-04 · Whitespace-only gender preservation

**Problem:** `apply_gender_normalization` has an early-return for `None` and `""` but not for whitespace-only strings. `"   "` reaches the engine, which strips it and returns `1` (male default), inconsistent with how `None` and `""` are handled.

**Where:** `src/excel_normalization/processing/normalization_pipeline.py`, `apply_gender_normalization()`.

**Exact fix — one line change:**

```python
# Current:
if original is None or original == "":
# Change to:
if original is None or str(original).strip() == "":
```

**Risk:** Minimal. One character change. Aligns whitespace behavior with None/empty behavior.

---

### F-05 · File size limit on upload

**Problem:** `await file.read()` reads the entire file into memory with no size check. A 500MB file would be accepted.

**Where:** `webapp/api/upload.py`.

**Exact fix:**

```python
MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB

@router.post("/upload", response_model=UploadResponse)
async def upload_file(
    file: UploadFile = File(...),
    upload_service: UploadService = Depends(get_upload_service),
) -> UploadResponse:
    file_bytes = await file.read()
    if len(file_bytes) > MAX_UPLOAD_BYTES:
        from fastapi import HTTPException
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Maximum size is {MAX_UPLOAD_BYTES // (1024*1024)} MB.",
        )
    return upload_service.handle_upload(file.filename or "upload.xlsx", file_bytes)
```

**Risk:** None. Pure additive guard.

---

### F-06 · Export output file cleanup

**Problem:** `ExportService.export` creates a new timestamped file on every call. The `output/` directory grows indefinitely.

**Where:** `webapp/services/export_service.py`, `ExportService.export()`.

**Exact fix — add cleanup before saving:**

```python
# Before: wb.save(str(output_path))
# Add: delete previous exports for this session (keep only latest)
import glob
stem = Path(record.original_filename).stem
for old_file in self.output_dir.glob(f"{stem}_normalized_*.xlsx"):
    try:
        old_file.unlink()
    except Exception:
        pass
```

**Risk:** Low. Only deletes files matching the same stem pattern.

---

### F-07 · new_value type coercion in edit

**Problem:** `CellEditRequest.new_value: str` forces all edits to strings. Editing `birth_year` (originally `int`) stores `"1990"` (str), which can cause type inconsistencies in downstream export.

**Where:** `webapp/services/edit_service.py`, `edit_cell()`.

**Exact fix — coerce to original type:**

```python
# After: row = sheet.rows[req.row_index]
# Before: sheet.rows[req.row_index][req.field_name] = req.new_value
# Add:
original_value = row.get(req.field_name)
coerced_value: Any = req.new_value
if isinstance(original_value, int):
    try: coerced_value = int(req.new_value)
    except (ValueError, TypeError): pass
elif isinstance(original_value, float):
    try: coerced_value = float(req.new_value)
    except (ValueError, TypeError): pass
sheet.rows[req.row_index][req.field_name] = coerced_value
```

**Risk:** Low. Falls back to string if coercion fails.

---

### F-08 · Session cleanup endpoint

**Problem:** Sessions accumulate in memory for the process lifetime. `SessionService.delete` exists but no API endpoint calls it.

**Where:** `webapp/api/workbook.py` (or a new `webapp/api/session.py`).

**Exact fix — add one endpoint:**

```python
@router.delete("/workbook/{session_id}", status_code=204)
def close_session(
    session_id: str,
    session_service: SessionService = Depends(get_session_service),
) -> None:
    """Remove a session from memory. Does not delete uploaded files."""
    session_service.delete(session_id)
```

**Risk:** None. `SessionService.delete` already exists and works correctly.


---

## PART 3 — UI/UX ENHANCEMENTS

Based on reading `webapp/static/app.js`, `webapp/templates/index.html`, and the actual API behavior.

---

### U-01 · Warn user before re-normalizing when edits exist

**Problem:** The user edits cells, then clicks "Run Normalization" — all edits are silently lost. There is no warning.

**Current code:** `runNormalization()` in `app.js` calls `POST /normalize` immediately with no confirmation.

**Fix — add a guard in `runNormalization()`:**

```javascript
async function runNormalization() {
    if (!state.sessionId) return;
    dismissError();

    // Check if there are unsaved edits in the current session
    const session = sessions.get(state.sessionId);
    if (session && session.hasEdits) {
        const confirmed = confirm(
            'Running normalization will discard your manual edits.\n\nContinue?'
        );
        if (!confirmed) return;
    }
    // ... rest of function
}
```

Also set `session.hasEdits = true` in `commitEdit()` after a successful PATCH, and reset it after normalization.

**Impact:** Prevents the most common data-loss scenario with zero backend changes.

---

### U-02 · Show normalization status per-row in the grid

**Problem:** After normalization, the grid shows `birth_date_status` and `identifier_status` columns but they contain Hebrew text that is not visually distinguished from data. A row with `"חודש לא תקין"` looks the same as a row with `""`.

**Current code:** `renderGrid()` applies class `status-cell` to status columns but the CSS does not visually differentiate error vs. empty status.

**Fix — in `renderGrid()`, add conditional class to status cells:**

```javascript
} else if (cls === 'status') {
    const statusText = String(value || '').trim();
    if (statusText === '') {
        td.className = 'status-cell status-ok';
    } else {
        td.className = 'status-cell status-error';
    }
}
```

Add to `style.css`:
```css
.status-error { background-color: #fff0f0; color: #c0392b; font-weight: 500; }
.status-ok    { background-color: transparent; }
```

**Impact:** Users can immediately see which rows have normalization problems without reading every status cell.

---

### U-03 · Show diff highlighting in the grid (corrected vs original)

**Problem:** The grid already has `corrected-changed` class logic in `renderGrid()`:
```javascript
td.className = (value !== null && value !== undefined && value !== origVal)
    ? 'corrected-changed' : 'corrected-cell';
```
But this comparison uses `!==` which fails for type mismatches (e.g., original `1` vs corrected `"1"`). Gender corrected is always `1` or `2` (int) while original is a string — so every gender row shows as "changed" even when it isn't meaningful.

**Fix — normalize comparison in `renderGrid()`:**

```javascript
const origVal = row[col.replace(/_corrected$/, '')];
const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
const corrStr = (value !== null && value !== undefined) ? String(value).trim() : '';
td.className = (corrStr !== origStr && corrStr !== '')
    ? 'corrected-changed' : 'corrected-cell';
```

**Impact:** Eliminates false "changed" highlights on gender and numeric fields. Users see only genuine corrections.

---

### U-04 · Normalize single sheet, not all sheets

**Problem:** `runNormalization()` always calls `POST /normalize` (all sheets). If the user is working on one sheet of a 3-sheet workbook, they wait for all three to normalize.

**Current code:**
```javascript
const result = await apiCall('POST', `/api/workbook/${state.sessionId}/normalize`);
```

The API already supports `?sheet=name`. The UI just doesn't use it.

**Fix — normalize only the current sheet:**

```javascript
const sheetParam = state.currentSheet
    ? `?sheet=${encodeURIComponent(state.currentSheet)}`
    : '';
const result = await apiCall('POST',
    `/api/workbook/${state.sessionId}/normalize${sheetParam}`);
```

**Impact:** Normalization of a single sheet is significantly faster. No backend changes needed.

---

### U-05 · Show row count and normalization success rate in the grid header

**Problem:** After normalization, the API returns `per_sheet_stats` with `rows` and `success_rate`. This data is currently shown only in `grid-stats` as a one-line text and is overwritten on the next sheet load.

**Current code in `runNormalization()`:**
```javascript
document.getElementById('grid-stats').textContent =
    `Normalization complete (${result.sheets_processed} sheet(s)) — ${stats}`;
```

**Fix — persist stats on the session object and show them in the sheet tab:**

```javascript
// In runNormalization(), after success:
result.per_sheet_stats.forEach(s => {
    const session = sessions.get(state.sessionId);
    if (session) {
        if (!session.sheetStats) session.sheetStats = {};
        session.sheetStats[s.sheet_name] = s;
    }
});
// Re-render sheet tabs with stats
renderSheetSelectorWithStats(session.sheetNames, session.sheetStats);
```

In `renderSheetSelector`, add a badge:
```javascript
const stat = (session.sheetStats || {})[name];
if (stat && stat.success_rate < 1.0) {
    btn.textContent = `${name} ⚠ ${Math.round(stat.success_rate * 100)}%`;
    btn.title = `${stat.rows} rows, ${Math.round(stat.success_rate * 100)}% success`;
}
```

**Impact:** Users immediately see which sheets have normalization problems without clicking each one.

---

### U-06 · Confirm before deleting multiple rows

**Problem:** `deleteSelectedRows()` deletes immediately with no confirmation. Deleting 50 rows is irreversible (until re-normalize, which has its own problems).

**Current code:**
```javascript
async function deleteSelectedRows() {
    if (state.selectedRows.size === 0) return;
    await _deleteRows([...state.selectedRows]);
}
```

**Fix — add confirmation for bulk deletes:**

```javascript
async function deleteSelectedRows() {
    if (state.selectedRows.size === 0) return;
    const n = state.selectedRows.size;
    if (n > 1) {
        const confirmed = confirm(`Delete ${n} rows? This cannot be undone without re-normalizing.`);
        if (!confirmed) return;
    }
    await _deleteRows([...state.selectedRows]);
}
```

**Impact:** Prevents accidental bulk deletion. Single-row delete (via ✕ button) remains instant.

---

### U-07 · Show which fields were changed per row (summary badge)

**Problem:** After normalization, a row with 8 corrected fields looks identical to a row with 0 corrections. The user has to scan every corrected column to find changed rows.

**Fix — add a "changes" badge column in `renderGrid()`:**

```javascript
// In the tbody loop, after building all cells:
const changedFields = displayColumns.filter(col => {
    if (!col.endsWith('_corrected')) return false;
    const orig = row[col.replace(/_corrected$/, '')];
    const corr = row[col];
    return corr !== null && corr !== undefined &&
           String(corr).trim() !== String(orig ?? '').trim();
});
if (changedFields.length > 0) {
    tr.classList.add('row-has-changes');
    // Add a small badge in the delete column
    const badge = document.createElement('span');
    badge.className = 'change-badge';
    badge.textContent = changedFields.length;
    badge.title = 'Fields changed: ' + changedFields.map(f => f.replace('_corrected','')).join(', ');
    tdDel.appendChild(badge);
}
```

Add to CSS:
```css
.row-has-changes { background-color: #f0f7ff; }
.change-badge { background: #3182ce; color: white; border-radius: 10px;
                padding: 1px 6px; font-size: 11px; margin-left: 4px; }
```

**Impact:** Users can immediately identify which rows were modified by normalization.

---

### U-08 · Keyboard shortcut for normalization

**Problem:** The workflow is: upload → view → normalize → export. The normalize button requires a mouse click. Power users processing many files benefit from keyboard access.

**Fix — add keyboard shortcut in `DOMContentLoaded`:**

```javascript
document.addEventListener('keydown', e => {
    // Ctrl+Enter or Cmd+Enter = normalize
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        if (state.sessionId) runNormalization();
    }
    // Ctrl+S or Cmd+S = export
    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        if (state.sessionId) exportWorkbook();
    }
});
```

Update the button labels to show the shortcut:
```html
<button id="normalize-btn" ...>▶ Run Normalization <kbd>Ctrl+↵</kbd></button>
<button id="export-btn" ...>⬇ Export / Download <kbd>Ctrl+S</kbd></button>
```

**Impact:** Zero backend changes. Significant speed improvement for repeated use.

---

### U-09 · Freeze header row in the data grid

**Problem:** The data grid has no frozen header. When scrolling a sheet with 200 rows, column names disappear.

**Current code:** `renderGrid()` creates a plain `<table>` with no sticky positioning.

**Fix — add to `style.css`:**

```css
.grid-container {
    max-height: 70vh;
    overflow-y: auto;
    overflow-x: auto;
}

.data-grid thead tr {
    position: sticky;
    top: 0;
    z-index: 10;
    background: #2d3748;
}
```

**Impact:** One CSS change. No JavaScript changes. Dramatically improves usability for large sheets.

---

### U-10 · Show upload progress for large files

**Problem:** `handleUpload()` disables the button and shows "Uploading..." but gives no progress feedback. For a 10MB file on a slow connection, the UI appears frozen.

**Current code:**
```javascript
statusDiv.innerHTML = `Uploading ${files.length} file(s)... <span class="loading"></span>`;
```

**Fix — use `XMLHttpRequest` with progress events for the upload:**

```javascript
function uploadWithProgress(file, onProgress) {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        const fd = new FormData();
        fd.append('file', file);
        xhr.upload.addEventListener('progress', e => {
            if (e.lengthComputable) onProgress(Math.round(e.loaded / e.total * 100));
        });
        xhr.addEventListener('load', () => {
            if (xhr.status >= 200 && xhr.status < 300) resolve(JSON.parse(xhr.responseText));
            else reject(new Error(JSON.parse(xhr.responseText).detail || `HTTP ${xhr.status}`));
        });
        xhr.addEventListener('error', () => reject(new Error('Network error')));
        xhr.open('POST', '/api/upload');
        xhr.send(fd);
    });
}
```

Then in `handleUpload()`:
```javascript
const data = await uploadWithProgress(file, pct => {
    statusDiv.textContent = `Uploading ${file.name}: ${pct}%`;
});
```

**Impact:** Users know the upload is progressing. Prevents "is it frozen?" confusion.

---

## PRIORITY SUMMARY

| Priority | Item | Type | Effort | Impact |
|---|---|---|---|---|
| 🔴 P1 | F-01: Replay edits after re-normalize | Functionality | ~8 lines | Prevents silent data loss |
| 🔴 P1 | U-01: Warn before re-normalize with edits | UI/UX | ~5 lines JS | Prevents silent data loss |
| 🔴 P1 | T-01: Test edit loss after re-normalize | Test | ~20 lines | Documents the bug |
| 🟠 P2 | F-02: Entry-before-birth validation | Functionality | ~15 lines | Catches logical errors |
| 🟠 P2 | F-03: DDMM/MMDD auto-detection | Functionality | ~20 lines | Fixes silent wrong dates |
| 🟠 P2 | U-09: Freeze grid header row | UI/UX | 5 lines CSS | Major usability win |
| 🟠 P2 | U-03: Fix corrected-changed comparison | UI/UX | ~5 lines JS | Removes false highlights |
| 🟡 P3 | F-04: Whitespace gender preservation | Functionality | 1 line | Consistency fix |
| 🟡 P3 | F-05: File size limit | Functionality | ~5 lines | Safety guard |
| 🟡 P3 | U-04: Normalize single sheet | UI/UX | 3 lines JS | Performance |
| 🟡 P3 | U-02: Status cell visual distinction | UI/UX | ~10 lines | Clarity |
| 🟡 P3 | U-06: Confirm bulk delete | UI/UX | ~5 lines JS | Safety |
| 🟢 P4 | F-06: Export file cleanup | Functionality | ~8 lines | Disk hygiene |
| 🟢 P4 | F-07: Edit type coercion | Functionality | ~10 lines | Type safety |
| 🟢 P4 | F-08: Session cleanup endpoint | Functionality | ~5 lines | Memory hygiene |
| 🟢 P4 | U-05: Sheet stats in tabs | UI/UX | ~15 lines JS | Visibility |
| 🟢 P4 | U-07: Row change badge | UI/UX | ~15 lines JS | Visibility |
| 🟢 P4 | U-08: Keyboard shortcuts | UI/UX | ~10 lines JS | Power users |
| 🟢 P4 | U-10: Upload progress | UI/UX | ~20 lines JS | Large files |
