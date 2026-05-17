from pathlib import Path


def test_upload_file_input_accepts_legacy_xls_files():
    html = Path("webapp/templates/index.html").read_text(encoding="utf-8")

    assert 'accept=".xlsx,.xlsm,.xls"' in html


def test_app_js_is_single_domcontentloaded_bootstrap():
    app_js = Path("webapp/static/app.js").read_text(encoding="utf-8")
    export_js = Path("webapp/static/js/export.js").read_text(encoding="utf-8")

    assert app_js.count("DOMContentLoaded") == 1
    assert "DOMContentLoaded" not in export_js
