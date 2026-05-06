"""SessionRecord dataclass for tracking in-memory session state."""

from dataclasses import dataclass, field
from typing import List, Optional
from src.excel_standardization.data_types import WorkbookDataset
from webapp.models.processing_report import ProcessingReport


@dataclass
class SelectedRowsConfig:
    """A SugMosad value applied to a specific set of rows identified by _row_uid.

    Attributes:
        sug_mosad:  The institution type value to apply.
        row_uids:   List of _row_uid strings identifying the target rows.
    """
    sug_mosad: str
    row_uids: List[str] = field(default_factory=list)


@dataclass
class SugMosadConfig:
    """Scoped SugMosad (institution type) configuration.

    Scope hierarchy (mutually exclusive):
        scope == "workbook"      — apply to all sheets (legacy behaviour).
        scope == "sheet"         — apply to one named sheet only.
        scope == "selected_rows" — apply to specific rows (by _row_uid) inside one sheet.

    Attributes:
        scope:         One of "workbook" | "sheet" | "selected_rows".
        sug_mosad:     The value to apply (used for "workbook" and "sheet" scopes).
        sheet_name:    Required for "sheet" and "selected_rows" scopes.
        selected_rows: List of SelectedRowsConfig; used only for "selected_rows" scope.
                       Up to 3 entries, each with its own sug_mosad and row_uids.
    """
    scope: str                                    # "workbook" | "sheet" | "selected_rows"
    sug_mosad: str = ""                           # used for workbook / sheet scope
    sheet_name: Optional[str] = None              # required for sheet / selected_rows scope
    selected_rows: List[SelectedRowsConfig] = field(default_factory=list)


@dataclass
class SessionRecord:
    """In-memory record for a user's working session.

    Attributes:
        session_id: UUID string identifying this session
        source_file_path: Path to the original uploaded file (never modified)
        working_copy_path: Path to the working copy used for processing
        original_filename: The original filename as uploaded by the user
        status: Current session status: "uploaded" | "standardized"
        workbook_dataset: Extracted/normalized WorkbookDataset (None until extracted)
        edits: Manual cell edits recorded as {(sheet_name, row_idx, field): new_value}
        mosad_id: Institution identifier (MosadID) — workbook-level
        mosad_name: Institution name in Hebrew/free text — used for export filename
        mosad_types: Up to 3 user-entered institution type values (SugMosad).
                     The first entry is the active/default value used in export.
                     Never auto-filled — always user-entered.
        sug_mosad_configs: Scoped SugMosad configurations.  When non-empty these
                     take precedence over the legacy mosad_types[0] workbook-level
                     default during export.
        processing_report: Latest non-sensitive ProcessingReport for this session.
    """

    session_id: str
    source_file_path: str
    working_copy_path: str
    original_filename: str
    status: str
    workbook_dataset: Optional[WorkbookDataset] = None
    edits: dict = field(default_factory=dict)
    mosad_id: str = ""
    mosad_name: str = ""
    mosad_types: List[str] = field(default_factory=list)
    sug_mosad_configs: List[SugMosadConfig] = field(default_factory=list)
    processing_report: Optional[ProcessingReport] = None
