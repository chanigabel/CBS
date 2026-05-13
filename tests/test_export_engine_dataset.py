from openpyxl import load_workbook

from src.excel_standardization.data_types import SheetDataset, WorkbookDataset
from src.excel_standardization.export.export_engine import ExportEngine


def test_export_engine_keeps_moved_passport_value_without_source_passport_column(tmp_path):
    engine = ExportEngine()
    sheet = SheetDataset(
        sheet_name=engine.SOURCE_SHEET_SPECS[0].source_sheet_name,
        header_row=1,
        header_rows_count=1,
        field_names=["id_number", "passport_corrected"],
        rows=[
            {
                "id_number": "ABC123",
                "id_number_corrected": "",
                "passport_corrected": "ABC123",
            }
        ],
    )
    workbook = WorkbookDataset(source_file="input.xlsx", sheets=[sheet])
    output_path = tmp_path / "export.xlsx"

    engine.export_from_normalized_dataset(workbook, str(output_path))

    wb = load_workbook(output_path)
    ws = wb["DayarimYahidim"]
    headers = [cell.value for cell in ws[1]]
    darkon_col = headers.index("Darkon") + 1
    assert ws.cell(row=2, column=darkon_col).value == "ABC123"
    wb.close()


def test_export_engine_uses_corrected_standardized_fields_only(tmp_path):
    engine = ExportEngine()
    sheet = SheetDataset(
        sheet_name=engine.SOURCE_SHEET_SPECS[0].source_sheet_name,
        header_row=1,
        header_rows_count=1,
        field_names=["first_name", "last_name", "father_name", "id_number", "gender"],
        rows=[
            {
                "first_name": "Original First",
                "first_name_corrected": "Corrected First",
                "last_name": "Original Last",
                "last_name_corrected": "Corrected Last",
                "father_name": "Original Father",
                "father_name_corrected": "Corrected Father",
                "id_number": "123",
                "id_number_corrected": "",
                "gender": "male",
                "gender_corrected": "",
            }
        ],
    )
    workbook = WorkbookDataset(source_file="input.xlsx", sheets=[sheet])
    output_path = tmp_path / "export.xlsx"

    engine.export_from_normalized_dataset(workbook, str(output_path))

    wb = load_workbook(output_path)
    ws = wb["DayarimYahidim"]
    headers = [cell.value for cell in ws[1]]
    first_col = headers.index("ShemPrati") + 1
    last_col = headers.index("ShemMishpaha") + 1
    father_col = headers.index("ShemHaAv") + 1
    id_col = headers.index("MisparZehut") + 1
    gender_col = headers.index("Min") + 1
    assert ws.cell(row=2, column=first_col).value == "Corrected First"
    assert ws.cell(row=2, column=last_col).value == "Corrected Last"
    assert ws.cell(row=2, column=father_col).value == "Corrected Father"
    assert ws.cell(row=2, column=id_col).value is None
    assert ws.cell(row=2, column=gender_col).value is None
    wb.close()
