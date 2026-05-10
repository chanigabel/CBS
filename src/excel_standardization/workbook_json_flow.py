"""Active workbook JSON flow helpers for the orchestrator facade."""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from .engines.date_engine import DateEngine
from .engines.gender_engine import GenderEngine
from .engines.identifier_engine import IdentifierEngine
from .engines.name_engine import NameEngine
from .export.export_engine import ExportEngine
from .io_layer.excel_reader import ExcelReader
from .io_layer.excel_to_json_extractor import ExcelToJsonExtractor
from .processing.standardization_pipeline import StandardizationPipeline
from .json_exporter import JsonExporter


# הפונקציה מחלצת WorkbookDataset מקובץ Excel כשלב הראשון בנתיב ה־Dataset הפעיל.
def extract_workbook(reader: ExcelReader, input_excel_path: str):
    extractor = ExcelToJsonExtractor(
        excel_reader=reader,
        skip_empty_rows=False,
        handle_formulas=True,
        preserve_types=True,
    )
    workbook_dataset = extractor.extract_workbook_to_json(input_excel_path)
    if not workbook_dataset.sheets:
        raise ValueError(f"No valid sheets found in workbook '{input_excel_path}'")
    return workbook_dataset


# הפונקציה בונה את ה־pipeline עם כל מנועי הסטנדרטיזציה הפעילים.
def build_pipeline(
    name_engine: NameEngine,
    gender_engine: GenderEngine,
    date_engine: DateEngine,
    identifier_engine: IdentifierEngine,
) -> StandardizationPipeline:
    return StandardizationPipeline(
        name_engine=name_engine,
        gender_engine=gender_engine,
        date_engine=date_engine,
        identifier_engine=identifier_engine,
        apply_name_standardization_enabled=True,
        apply_gender_standardization_enabled=True,
        apply_date_standardization_enabled=True,
        apply_identifier_standardization_enabled=True,
    )


# הפונקציה מפעילה סטנדרטיזציה על כל גיליון שחולץ לפני שלב היצוא.
def normalize_sheets(
    sheets: list,
    name_engine: NameEngine,
    gender_engine: GenderEngine,
    date_engine: DateEngine,
    identifier_engine: IdentifierEngine,
) -> list:
    pipeline = build_pipeline(name_engine, gender_engine, date_engine, identifier_engine)
    normalized_sheets = []

    for sheet in sheets:
        normalized_sheets.append(pipeline.normalize_dataset(sheet))

    if not normalized_sheets:
        raise ValueError("No sheets were successfully normalized")

    return normalized_sheets


# הפונקציה מחשבת נתיב יצוא ברירת מחדל בלי לדרוס קבצים קיימים.
def default_export_path(input_excel_path: str) -> str:
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


# הפונקציה מריצה את נתיב Excel -> Dataset -> Pipeline -> Export לקובץ תוצאה.
def export_vba_parity_workbook_from_json(
    reader: ExcelReader,
    name_engine: NameEngine,
    gender_engine: GenderEngine,
    date_engine: DateEngine,
    identifier_engine: IdentifierEngine,
    input_excel_path: str,
    output_excel_path: Optional[str] = None,
) -> str:
    workbook_dataset = extract_workbook(reader, input_excel_path)
    workbook_dataset.sheets = normalize_sheets(
        workbook_dataset.sheets,
        name_engine,
        gender_engine,
        date_engine,
        identifier_engine,
    )

    if output_excel_path is None:
        output_excel_path = default_export_path(input_excel_path)

    engine = ExportEngine()
    return engine.export_from_normalized_dataset(
        workbook_dataset,
        output_excel_path,
        corrected_columns_by_sheet=None,
    )


# הפונקציה מייצאת את תוצאת החילוץ הגולמית ל־JSON לצורכי בדיקה או דיבוג.
def export_raw_json(reader: ExcelReader, input_excel_path: str, output_json_path: str) -> None:
    workbook_dataset = extract_workbook(reader, input_excel_path)
    JsonExporter(indent=2, ensure_ascii=False).export_workbook_to_json(
        workbook_dataset,
        output_json_path,
    )


# הפונקציה מייצאת Dataset לאחר סטנדרטיזציה ל־JSON בלי לכתוב Excel.
def export_normalized_json(
    reader: ExcelReader,
    name_engine: NameEngine,
    gender_engine: GenderEngine,
    date_engine: DateEngine,
    identifier_engine: IdentifierEngine,
    input_excel_path: str,
    output_json_path: str,
) -> None:
    workbook_dataset = extract_workbook(reader, input_excel_path)
    workbook_dataset.sheets = normalize_sheets(
        workbook_dataset.sheets,
        name_engine,
        gender_engine,
        date_engine,
        identifier_engine,
    )

    JsonExporter(indent=2, ensure_ascii=False).export_workbook_to_json(
        workbook_dataset,
        output_json_path,
    )
