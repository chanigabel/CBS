"""Verify the column-index row is no longer included in extracted data."""
import json
import os
import sys

sys.path.insert(0, ".")

from src.excel_standardization.orchestrator import standardizationOrchestrator

input_path = r"C:\Users\ch058\OneDrive\שולחן העבודה\python_automation\Automations_DEV.xlsx"
output_dir = r"C:\Users\ch058\OneDrive\שולחן העבודה\python_automation\pipeline_output"
output_excel = os.path.join(output_dir, "output.xlsx")

orch = standardizationOrchestrator()
orch.process_workbook_json(input_path, output_excel)

with open(os.path.join(output_dir, "raw_dataset.json"), encoding="utf-8") as f:
    data = json.load(f)

for sheet in data["sheets"]:
    name = sheet["sheet_name"]
    rows = sheet["rows"]
    total = sheet["metadata"]["total_rows"]
    print(f"Sheet: {name!r}  total_rows={total}")
    if rows:
        r0 = rows[0]
        items = list(r0.items())[:6]
        print(f"  First row: {dict(items)}")
