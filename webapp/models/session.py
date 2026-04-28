"""SessionRecord dataclass for tracking in-memory session state."""

from dataclasses import dataclass, field
from typing import List, Optional
from src.excel_standardization.data_types import WorkbookDataset


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
