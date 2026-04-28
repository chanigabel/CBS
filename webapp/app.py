"""FastAPI application entry point for the Excel standardization Web App.

Development:
    uvicorn webapp.app:app --reload

Packaged (via launcher):
    Excelstandardization.exe
"""

import logging
import hashlib
import sys
from pathlib import Path

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, FileResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.base import BaseHTTPMiddleware
from contextlib import asynccontextmanager

from webapp.api import upload, workbook, standardize, edit, export, institution

# ---------------------------------------------------------------------------
# Asset path resolution — works both in development and inside a PyInstaller
# one-folder bundle.  PyInstaller sets sys._MEIPASS to the temp extraction
# directory; in development __file__ gives us the webapp package directory.
# ---------------------------------------------------------------------------

def _asset_base() -> Path:
    """Return the directory that contains the webapp package assets."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        # PyInstaller bundle: assets are extracted to _MEIPASS/webapp/
        return Path(sys._MEIPASS) / "webapp"  # type: ignore[attr-defined]
    # Development: this file lives at webapp/app.py
    return Path(__file__).parent


_ASSET_BASE = _asset_base()
_STATIC_DIR = _ASSET_BASE / "static"
_TEMPLATES_DIR = _ASSET_BASE / "templates"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# App lifecycle
# ---------------------------------------------------------------------------


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Ensure runtime directories exist on startup."""
    # Dependencies module already creates them; this is a belt-and-suspenders guard.
    from webapp.dependencies import UPLOADS_DIR, WORK_DIR, OUTPUT_DIR
    for directory in (UPLOADS_DIR, WORK_DIR, OUTPUT_DIR):
        directory.mkdir(parents=True, exist_ok=True)
    logger.info("Excel standardization Web App started.")
    logger.info(f"Static assets: {_STATIC_DIR}")
    logger.info(f"Templates:     {_TEMPLATES_DIR}")
    yield


# ---------------------------------------------------------------------------
# FastAPI application
# ---------------------------------------------------------------------------

app = FastAPI(
    title="Excel standardization Web App",
    description="Local web application for standardizing Excel workbooks",
    version="1.0.0",
    lifespan=lifespan,
)

# ---------------------------------------------------------------------------
# Cache-busting middleware — prevents the browser from serving stale static
# files after the app is updated.  All /static/* responses get:
#   Cache-Control: no-cache, no-store, must-revalidate
# ---------------------------------------------------------------------------

class NoCacheStaticMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        response = await call_next(request)
        if request.url.path.startswith("/static/"):
            response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
            response.headers["Pragma"] = "no-cache"
            response.headers["Expires"] = "0"
        return response

app.add_middleware(NoCacheStaticMiddleware)

app.mount("/static", StaticFiles(directory=str(_STATIC_DIR)), name="static")
templates = Jinja2Templates(directory=str(_TEMPLATES_DIR))

# Register API routers
app.include_router(upload.router, prefix="/api")
app.include_router(workbook.router, prefix="/api")
app.include_router(standardize.router, prefix="/api")
app.include_router(edit.router, prefix="/api")
app.include_router(export.router, prefix="/api")
app.include_router(institution.router, prefix="/api")


def _file_hash(path: Path, length: int = 8) -> str:
    try:
        return hashlib.md5(path.read_bytes()).hexdigest()[:length]
    except Exception:
        return "0"


@app.get("/favicon.ico", include_in_schema=False)
def favicon() -> FileResponse:
    """Serve favicon.ico directly so browsers don't get a 404."""
    return FileResponse(str(_STATIC_DIR / "favicon.ico"))


@app.get("/", response_class=HTMLResponse)
def index(request: Request) -> HTMLResponse:
    """Serve the single-page UI with cache-busting version strings.

    Hashes are computed per-request so that reloading the server with
    --reload always serves the current file version to the browser.
    """
    return templates.TemplateResponse(
        request,
        "index.html",
        {
            "v_js":  _file_hash(_STATIC_DIR / "app.js"),
            "v_css": _file_hash(_STATIC_DIR / "style.css"),
        },
    )
