"""Export router: single-file and bulk ZIP export."""

import io
import unicodedata
import zipfile
from pathlib import Path
from typing import List
from urllib.parse import quote

from fastapi import APIRouter, Depends
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import BaseModel

from webapp.dependencies import get_export_service
from webapp.services.export_service import ExportService

router = APIRouter(tags=["export"])


def _content_disposition(filename: str) -> str:
    """Build a safe Content-Disposition header for *filename*.

    Uses the RFC 5987 / RFC 6266 dual-value form so that:
    - Modern browsers use ``filename*`` (UTF-8, percent-encoded) and
      display the real filename including Hebrew/non-ASCII characters.
    - Legacy clients fall back to the ASCII ``filename`` fallback, which
      strips non-ASCII characters but never raises a codec error.

    Example output:
        attachment; filename="normalized.xlsx"; filename*=UTF-8''%D7%A7%D7%95%D7%91%D7%A5_normalized.xlsx
    """
    # ASCII fallback: keep only printable ASCII, replace everything else with '_'
    ascii_fallback = "".join(
        c if c.isascii() and c.isprintable() and c not in ('"', "\\") else "_"
        for c in unicodedata.normalize("NFC", filename)
    ).strip("_") or "export.xlsx"

    # UTF-8 percent-encoded value for filename*
    encoded = quote(filename, safe="-_.~")  # RFC 3986 unreserved chars kept as-is

    return f'attachment; filename="{ascii_fallback}"; filename*=UTF-8\'\'{encoded}'


@router.post("/workbook/{session_id}/export")
def export_workbook(
    session_id: str,
    export_service: ExportService = Depends(get_export_service),
) -> FileResponse:
    """Export a single session's workbook as a downloadable Excel file."""
    output_path = export_service.export(session_id)
    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=output_path.name,
        headers={"Content-Disposition": _content_disposition(output_path.name)},
    )


class BulkExportRequest(BaseModel):
    session_ids: List[str]


@router.post("/export/bulk")
def export_bulk(
    req: BulkExportRequest,
    export_service: ExportService = Depends(get_export_service),
) -> StreamingResponse:
    """Export multiple sessions as a single ZIP archive."""
    if not req.session_ids:
        from fastapi import HTTPException
        raise HTTPException(status_code=400, detail="session_ids must not be empty.")

    buf = io.BytesIO()
    exported = 0

    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for sid in req.session_ids:
            try:
                output_path = export_service.export(sid)
                zf.write(output_path, arcname=output_path.name)
                exported += 1
            except Exception as exc:
                import logging
                logging.getLogger(__name__).warning(
                    f"Bulk export: skipping session {sid}: {exc}"
                )

    if exported == 0:
        from fastapi import HTTPException
        raise HTTPException(status_code=500, detail="All exports failed.")

    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="export_bulk.zip"'},
    )
