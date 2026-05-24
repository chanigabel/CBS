from pathlib import Path
import re


def test_upload_file_input_accepts_legacy_xls_files():
    html = Path("webapp/templates/index.html").read_text(encoding="utf-8")

    assert 'accept=".xlsx,.xlsm,.xls"' in html


def test_app_js_is_single_domcontentloaded_bootstrap():
    app_js = Path("webapp/static/app.js").read_text(encoding="utf-8")
    export_js = Path("webapp/static/js/export.js").read_text(encoding="utf-8")

    assert app_js.count("DOMContentLoaded") == 1
    assert "DOMContentLoaded" not in export_js


def test_grid_selection_uses_backend_row_uid_not_generated_client_side():
    grid_js = Path("webapp/static/js/grid.js").read_text(encoding="utf-8")

    assert "function getRowUid(row)" in grid_js
    assert "row._row_uid || row.row_uid" in grid_js
    assert "crypto.randomUUID" not in grid_js
    assert "Math.random" not in grid_js


def test_sheet_reload_reconciles_selected_rows_with_backend_payload():
    upload_js = Path("webapp/static/js/upload.js").read_text(encoding="utf-8")

    assert "validRowUids" in upload_js
    assert "state.selectedRows = new Set([...state.selectedRows].filter" in upload_js


def test_keyboard_shortcuts_are_centralized_and_reuse_button_actions():
    app_js = Path("webapp/static/app.js").read_text(encoding="utf-8")

    assert "const keyboardShortcutRegistry" in app_js
    assert "document.addEventListener('keydown', handleKeyboardShortcut)" in app_js
    assert "action: () => undoLastGridEdit()" in app_js
    assert "action: () => deleteSelectedRowsFromShortcut()" in app_js
    assert "isKeyboardEditableTarget(event.target)" in app_js


def test_export_shortcuts_are_mapped_to_correct_actions():
    app_js = Path("webapp/static/app.js").read_text(encoding="utf-8")

    full_export = re.search(
        r"id: 'export-workbook'.*?event\.key\.toLowerCase\(\) === 's'.*?action: \(\) => exportWorkbook\(\)",
        app_js,
        re.DOTALL,
    )
    sheet_export = re.search(
        r"id: 'export-current-sheet'.*?event\.shiftKey.*?event\.key\.toLowerCase\(\) === 'e'.*?action: \(\) => exportCurrentSheetFromShortcut\(\)",
        app_js,
        re.DOTALL,
    )

    assert full_export
    assert sheet_export
    assert "id: 'export-current-sheet'" in app_js
    assert "action: () => exportCurrentSheet()" not in app_js
    assert "אין גיליון נבחר לייצוא" in app_js


def test_shortcuts_are_ignored_while_typing_in_editors():
    app_js = Path("webapp/static/app.js").read_text(encoding="utf-8")

    assert "textarea, select, [role=\"textbox\"]" in app_js
    assert "target.matches && target.matches('input')" in app_js
    assert "[contenteditable=\"true\"]" in app_js
    assert "if (isKeyboardEditableTarget(event.target)) return;" in app_js


def test_shortcut_hints_are_shown_on_action_buttons():
    html = Path("webapp/templates/index.html").read_text(encoding="utf-8")

    assert 'title="ייצוא קובץ (Ctrl+S)"' in html
    assert 'title="ייצוא גיליון (Ctrl+Shift+E)"' in html
    assert 'title="הרצת סטנדרטיזציה (Ctrl+Enter)"' in html
    assert 'title="בחר שורות כדי למחוק (Delete)"' in html
    assert 'title="אין שינוי לביטול (Ctrl+Z)"' in html

    assert '<span class="shortcut-hint">Ctrl+S</span>' in html
    assert '<span class="shortcut-hint">Ctrl+Shift+E</span>' in html
    assert '<span class="shortcut-hint">Ctrl+Enter</span>' in html
    assert '<span class="shortcut-hint">Ctrl+Z</span>' in html
    assert '<span class="shortcut-hint">Delete</span>' in html
    assert '<span class="shortcut-hint">Esc</span>' in html
    assert '<kbd class="kbd-hint">' not in html


def test_full_export_button_does_not_show_single_sheet_shortcut():
    html = Path("webapp/templates/index.html").read_text(encoding="utf-8")

    full_export_button = re.search(r'<button id="export-btn".*?</button>', html, re.DOTALL)
    sheet_export_button = re.search(r'<button id="export-sheet-btn".*?</button>', html, re.DOTALL)

    assert full_export_button
    assert sheet_export_button
    assert "Ctrl+S" in full_export_button.group(0)
    assert "Ctrl+Shift+E" not in full_export_button.group(0)
    assert "Ctrl+Shift+E" in sheet_export_button.group(0)
