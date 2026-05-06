"""Dataset-pipeline orchestrator for active workbook processing.

The legacy direct Excel/VBA parity implementation was moved to
``archive_legacy/``. The active runtime path is:

Upload/Web/API -> WorkbookDataset/SheetDataset -> standardizationPipeline ->
ExportService/ExportEngine.
"""

import logging
from pathlib import Path
from typing import Optional

from .engines.date_engine import DateEngine
from .engines.gender_engine import GenderEngine
from .engines.identifier_engine import IdentifierEngine
from .engines.name_engine import NameEngine
from .engines.text_processor import TextProcessor
from .export.export_engine import ExportEngine
from .io_layer.excel_reader import ExcelReader
from .io_layer.excel_to_json_extractor import ExcelToJsonExtractor
from .json_exporter import JsonExporter
from .processing.standardization_pipeline import standardizationPipeline


LEGACY_DISABLED_MESSAGE = (
    "Disabled legacy direct Excel path. Use the Web/Dataset pipeline instead."
)


class standardizationOrchestrator:
    """Coordinates workbook extraction, dataset normalization, and export.

    This active orchestrator intentionally does not import or instantiate the old
    worksheet field processors. Those direct Excel processors are archived for
    historical reference only.
    """

    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)
        self.reader = ExcelReader()

        text_processor = TextProcessor()
        self.name_engine = NameEngine(text_processor)
        self.gender_engine = GenderEngine()
        self.date_engine = DateEngine()
        self.identifier_engine = IdentifierEngine()

    def normalize_workbook(self, file_path: str) -> None:
        """Disabled legacy direct Excel path.

        Disabled legacy direct Excel path. Use the Web/Dataset pipeline instead:
        Upload/Web/API -> WorkbookDataset/SheetDataset ->
        standardizationPipeline -> ExportService/ExportEngine.
        """
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    def process_workbook_json(self, input_excel_path: str, output_excel_path: str) -> None:
        """Disabled legacy direct Excel path.

        This public entry point historically ran direct worksheet processors
        against openpyxl worksheets. It is disabled so active processing remains
        on the Web/Dataset pipeline.
        """
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    def export_vba_parity_workbook_from_processors(
        self, input_excel_path: str, output_excel_path: Optional[str] = None
    ) -> str:
        """Disabled legacy direct Excel path.

        The processor-based Excel export is archived. Use
        ``export_vba_parity_workbook_from_json`` or the Web/API flow.
        """
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    def process_worksheet(self, worksheet: object) -> None:
        """Retained legacy helper name; no active public entry point calls it."""
        raise RuntimeError(LEGACY_DISABLED_MESSAGE)

    def export_raw_json(self, input_excel_path: str, output_json_path: str) -> None:
        """Export a raw WorkbookDataset JSON representation from an Excel file."""
        self.logger.info("Exporting raw JSON: %s -> %s", input_excel_path, output_json_path)

        workbook_dataset = self._extract_workbook(input_excel_path)
        if not workbook_dataset.sheets:
            raise ValueError(f"No valid sheets found in workbook '{input_excel_path}'")

        JsonExporter(indent=2, ensure_ascii=False).export_workbook_to_json(
            workbook_dataset,
            output_json_path,
        )

    def export_normalized_json(self, input_excel_path: str, output_json_path: str) -> None:
        """Extract, normalize, and export a WorkbookDataset JSON file."""
        self.logger.info(
            "Exporting normalized JSON: %s -> %s",
            input_excel_path,
            output_json_path,
        )

        workbook_dataset = self._extract_workbook(input_excel_path)
        workbook_dataset.sheets = self._normalize_sheets(workbook_dataset.sheets)

        JsonExporter(indent=2, ensure_ascii=False).export_workbook_to_json(
            workbook_dataset,
            output_json_path,
        )

    def export_vba_parity_workbook_from_json(
        self,
        input_excel_path: str,
        output_excel_path: Optional[str] = None,
    ) -> str:
        """Run the active Dataset pipeline and export an Excel workbook.

        Pipeline:
            ExcelToJsonExtractor -> standardizationPipeline -> ExportEngine
        """
        workbook_dataset = self._extract_workbook(input_excel_path)
        workbook_dataset.sheets = self._normalize_sheets(workbook_dataset.sheets)

        if output_excel_path is None:
            output_excel_path = self._default_export_path(input_excel_path)

        engine = ExportEngine()
        return engine.export_from_normalized_dataset(
            workbook_dataset,
            output_excel_path,
            corrected_columns_by_sheet=None,
        )

    def _extract_workbook(self, input_excel_path: str):
        extractor = ExcelToJsonExtractor(
            excel_reader=self.reader,
            skip_empty_rows=False,
            handle_formulas=True,
            preserve_types=True,
        )
        workbook_dataset = extractor.extract_workbook_to_json(input_excel_path)
        if not workbook_dataset.sheets:
            raise ValueError(f"No valid sheets found in workbook '{input_excel_path}'")
        return workbook_dataset

    def _build_pipeline(self) -> standardizationPipeline:
        return standardizationPipeline(
            name_engine=self.name_engine,
            gender_engine=self.gender_engine,
            date_engine=self.date_engine,
            identifier_engine=self.identifier_engine,
            apply_name_standardization_enabled=True,
            apply_gender_standardization_enabled=True,
            apply_date_standardization_enabled=True,
            apply_identifier_standardization_enabled=True,
        )

    def _normalize_sheets(self, sheets: list) -> list:
        pipeline = self._build_pipeline()
        normalized_sheets = []

        for sheet in sheets:
            normalized_sheets.append(pipeline.normalize_dataset(sheet))

        if not normalized_sheets:
            raise ValueError("No sheets were successfully normalized")

        return normalized_sheets

    def _default_export_path(self, input_excel_path: str) -> str:
        src = Path(input_excel_path)
        desktop = Path.home() / "Desktop"
        ext = src.suffix.lower()
        if ext not in [".xlsx", ".xlsm"]:
            ext = ".xlsx"

        candidate = desktop / f"{src.stem}_Export{ext}"
        suffix = 1
        while candidate.exists():
            candidate = desktop / f"{src.stem}_Export ({suffix}){ext}"
            suffix += 1

        return str(candidate)
