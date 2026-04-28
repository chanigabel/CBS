# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for Excel standardization Web App
#
# Build with:
#   pyinstaller Excelstandardization.spec
#
# Output: dist/Excelstandardization/Excelstandardization.exe

import sys
from pathlib import Path

ROOT = Path(SPECPATH)

block_cipher = None

# ---------------------------------------------------------------------------
# Data files to bundle
# ---------------------------------------------------------------------------
# Format: (source_glob_or_path, destination_folder_inside_bundle)

datas = [
    # Web UI assets
    (str(ROOT / "webapp" / "templates"), "webapp/templates"),
    (str(ROOT / "webapp" / "static"),    "webapp/static"),
]

# ---------------------------------------------------------------------------
# Hidden imports that PyInstaller's static analysis misses
# ---------------------------------------------------------------------------
hidden_imports = [
    # FastAPI / Starlette internals
    "starlette.routing",
    "starlette.middleware",
    "starlette.middleware.base",
    "starlette.staticfiles",
    "starlette.templating",
    "starlette.responses",
    "starlette.requests",
    "starlette.datastructures",
    "starlette.background",
    "starlette.concurrency",
    "starlette.exceptions",
    "starlette.types",
    # Uvicorn
    "uvicorn.logging",
    "uvicorn.loops",
    "uvicorn.loops.auto",
    "uvicorn.loops.asyncio",
    "uvicorn.protocols",
    "uvicorn.protocols.http",
    "uvicorn.protocols.http.auto",
    "uvicorn.protocols.http.h11_impl",
    "uvicorn.protocols.websockets",
    "uvicorn.protocols.websockets.auto",
    "uvicorn.lifespan",
    "uvicorn.lifespan.on",
    # Pydantic
    "pydantic",
    "pydantic.v1",
    "pydantic_core",
    # Jinja2
    "jinja2",
    "jinja2.ext",
    # openpyxl
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.utils.exceptions",
    # python-multipart
    "multipart",
    # anyio / h11
    "anyio",
    "anyio._backends._asyncio",
    "h11",
    # Our own packages
    "webapp",
    "webapp.app",
    "webapp.dependencies",
    "webapp.api",
    "webapp.api.upload",
    "webapp.api.workbook",
    "webapp.api.normalize",
    "webapp.api.edit",
    "webapp.api.export",
    "webapp.services",
    "webapp.services.session_service",
    "webapp.services.upload_service",
    "webapp.services.workbook_service",
    "webapp.services.standardization_service",
    "webapp.services.edit_service",
    "webapp.services.export_service",
    "webapp.services.mosad_id_scanner",
    "webapp.services.derived_columns",
    "webapp.models",
    "webapp.models.session",
    "webapp.models.requests",
    "webapp.models.responses",
    "src.excel_standardization",
    "src.excel_standardization.orchestrator",
    "src.excel_standardization.data_types",
    "src.excel_standardization.io_layer",
    "src.excel_standardization.io_layer.excel_reader",
    "src.excel_standardization.io_layer.excel_to_json_extractor",
    "src.excel_standardization.io_layer.excel_writer",
    "src.excel_standardization.processing",
    "src.excel_standardization.processing.standardization_pipeline",
    "src.excel_standardization.processing.name_processor",
    "src.excel_standardization.processing.gender_processor",
    "src.excel_standardization.processing.date_processor",
    "src.excel_standardization.processing.identifier_processor",
    "src.excel_standardization.processing.field_processor",
    "src.excel_standardization.engines",
    "src.excel_standardization.engines.name_engine",
    "src.excel_standardization.engines.gender_engine",
    "src.excel_standardization.engines.date_engine",
    "src.excel_standardization.engines.identifier_engine",
    "src.excel_standardization.engines.text_processor",
    "src.excel_standardization.export",
    "src.excel_standardization.export.export_engine",
]

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------
a = Analysis(
    ["launcher.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Dev/test tools — not needed at runtime
        "pytest",
        "hypothesis",
        "black",
        "mypy",
        "flake8",
        "IPython",
        "jupyter",
        "matplotlib",
        "numpy",
        "pandas",
        "PIL",
        "tkinter",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------------------------------------------------------------------------
# One-folder build (recommended — faster startup than onefile)
# ---------------------------------------------------------------------------
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Excelstandardization",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,          # Keep console so users can see errors / Ctrl+C
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,             # Set to "installer/icon.ico" if you add an icon
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Excelstandardization",
)
