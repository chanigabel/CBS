"""Shared FastAPI dependency injection for service instances.

Runtime directories are resolved to a writable location that works both
during development (relative to CWD) and inside a PyInstaller-packaged exe
(under %LOCALAPPDATA%\\ExcelNormalization).
"""

import os
import sys
from pathlib import Path

from webapp.services.session_service import SessionService
from webapp.services.upload_service import UploadService
from webapp.services.workbook_service import WorkbookService
from webapp.services.normalization_service import NormalizationService
from webapp.services.edit_service import EditService
from webapp.services.export_service import ExportService


def _get_data_dir() -> Path:
    """Return the writable data directory for runtime files.

    - Packaged (PyInstaller): %LOCALAPPDATA%\\ExcelNormalization
    - Development: project root (current working directory)
    """
    if getattr(sys, "frozen", False):
        # Running inside a PyInstaller bundle
        base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "ExcelNormalization"
    else:
        base = Path.cwd()
    return base


_data_dir = _get_data_dir()

# Runtime directories
UPLOADS_DIR = _data_dir / "uploads"
WORK_DIR = _data_dir / "work"
OUTPUT_DIR = _data_dir / "output"

# Ensure they exist at import time so services can use them immediately
for _d in (UPLOADS_DIR, WORK_DIR, OUTPUT_DIR):
    _d.mkdir(parents=True, exist_ok=True)

# Shared service instances (singletons for the process lifetime)
_session_service = SessionService()
_upload_service = UploadService(_session_service, UPLOADS_DIR, WORK_DIR)
_workbook_service = WorkbookService(_session_service)
_normalization_service = NormalizationService(_session_service)
_edit_service = EditService(_session_service)
_export_service = ExportService(_session_service, OUTPUT_DIR)


def get_session_service() -> SessionService:
    return _session_service


def get_upload_service() -> UploadService:
    return _upload_service


def get_workbook_service() -> WorkbookService:
    return _workbook_service


def get_normalization_service() -> NormalizationService:
    return _normalization_service


def get_edit_service() -> EditService:
    return _edit_service


def get_export_service() -> ExportService:
    return _export_service
