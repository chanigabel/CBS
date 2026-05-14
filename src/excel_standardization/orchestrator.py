"""Dataset-pipeline orchestrator for active workbook processing.

The legacy direct Excel/VBA parity implementation was moved to
``archive_legacy/``. The active runtime path is:

Upload/Web/API -> WorkbookDataset/SheetDataset -> StandardizationPipeline ->
ExportService/ExportEngine.
"""

import logging
from typing import Optional

from .engines.date_engine import DateEngine
from .engines.gender_engine import GenderEngine
from .engines.identifier_engine import IdentifierEngine
from .engines.name_engine import NameEngine
from .engines.text_processor import TextProcessor
from .engine_management import get_default_engine_manager
from .io_layer.excel_reader import ExcelReader
from .workbook_json_flow import (
    export_normalized_json as export_normalized_json_flow,
    export_raw_json as export_raw_json_flow,
    export_vba_parity_workbook_from_json as export_vba_parity_workbook_from_json_flow,
)


LEGACY_DISABLED_MESSAGE = (
    "Disabled legacy direct Excel path. Use the Web/Dataset pipeline instead."
)


# המחלקה משמשת Facade לנתיב הפעיל: קריאות JSON/export וחסימת נתיבי legacy ישירים.
class StandardizationOrchestrator:
    """Coordinates workbook extraction, dataset normalization, and export.

    This active orchestrator intentionally does not import or instantiate the old
    worksheet field processors. Those direct Excel processors are archived for
    historical reference only.
    """

    # הפונקציה מאתחלת את קורא ה־Excel והמנועים העסקיים המשמשים את ה־Dataset pipeline.
    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)
        self.reader = ExcelReader()

        text_processor = TextProcessor()
        self.name_engine = NameEngine(text_processor)
        self.gender_engine = GenderEngine()
        self.date_engine = DateEngine()
        self.identifier_engine = IdentifierEngine()
        self.engine_manager = get_default_engine_manager()

    # הפונקציה חוסמת שימוש בנתיב legacy ישיר כדי לשמור על נתיב runtime יחיד.
    def normalize_workbook(self, file_path: str) -> None:
        """Disabled legacy direct Excel path.

        Disabled legacy direct Excel path. Use the Web/Dataset pipeline instead:
        Upload/Web/API -> WorkbookDataset/SheetDataset ->
        StandardizationPipeline -> ExportService/ExportEngine.
        """
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    # הפונקציה חוסמת entry point ישן שעיבד workbook ישירות דרך worksheet processors.
    def process_workbook_json(self, input_excel_path: str, output_excel_path: str) -> None:
        """Disabled legacy direct Excel path.

        This public entry point historically ran direct worksheet processors
        against openpyxl worksheets. It is disabled so active processing remains
        on the Web/Dataset pipeline.
        """
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    # הפונקציה חוסמת יצוא ישן מתוך processors ומכוונת לנתיב JSON הפעיל.
    def export_vba_parity_workbook_from_processors(
        self, input_excel_path: str, output_excel_path: Optional[str] = None
    ) -> str:
        """Disabled legacy direct Excel path.

        The processor-based Excel export is archived. Use
        ``export_vba_parity_workbook_from_json`` or the Web/API flow.
        """
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    # הפונקציה שומרת שם helper ישן לתאימות, אך אינה חלק מהזרימה הפעילה.
    def process_worksheet(self, worksheet: object) -> None:
        """Retained legacy helper name; no active public entry point calls it."""
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    # הפונקציה מפעילה חילוץ גולמי ל־JSON דרך ה־facade לצורכי בדיקה.
    def export_raw_json(self, input_excel_path: str, output_json_path: str) -> None:
        """Export a raw WorkbookDataset JSON representation from an Excel file."""
        self.logger.info("Exporting raw JSON: %s -> %s", input_excel_path, output_json_path)
        export_raw_json_flow(self.reader, input_excel_path, output_json_path)

    # הפונקציה מפעילה חילוץ וסטנדרטיזציה ומייצאת את ה־Dataset המנורמל ל־JSON.
    def export_normalized_json(self, input_excel_path: str, output_json_path: str) -> None:
        """Extract, normalize, and export a WorkbookDataset JSON file."""
        self.logger.info(
            "Exporting normalized JSON: %s -> %s",
            input_excel_path,
            output_json_path,
        )
        export_normalized_json_flow(
            self.reader,
            self.name_engine,
            self.gender_engine,
            self.date_engine,
            self.identifier_engine,
            input_excel_path,
            output_json_path,
            self.engine_manager,
        )

    # הפונקציה היא entry point פעיל ליצירת קובץ Excel מתוקנן מנתיב ה־Dataset.
    def export_vba_parity_workbook_from_json(
        self,
        input_excel_path: str,
        output_excel_path: Optional[str] = None,
    ) -> str:
        """Run the active Dataset pipeline and export an Excel workbook.

        Pipeline:
            ExcelToJsonExtractor -> StandardizationPipeline -> ExportEngine
        """
        return export_vba_parity_workbook_from_json_flow(
            self.reader,
            self.name_engine,
            self.gender_engine,
            self.date_engine,
            self.identifier_engine,
            input_excel_path,
            output_excel_path,
            self.engine_manager,
        )

# Backward-compatible alias for callers that still import the legacy name.
standardizationOrchestrator = StandardizationOrchestrator
