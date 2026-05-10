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
