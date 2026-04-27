from pathlib import Path

from openpyxl import load_workbook

from src.excel_normalization.orchestrator import NormalizationOrchestrator


def main() -> None:
    base = Path("B.xlsx").resolve()
    if not base.exists():
        raise SystemExit(f"B.xlsx not found at {base}")

    out_path = Path("python_auto.xlsx").resolve()

    orch = NormalizationOrchestrator()

    # Load workbook in the same way as export_vba_parity_workbook_from_processors.
    src = base
    file_ext = src.suffix.lower()
    is_macro_enabled = file_ext in [".xlsm", ".xltm", ".xlam"]

    wb = load_workbook(
        str(base),
        data_only=False,
        keep_vba=is_macro_enabled,
        keep_links=False,
    )

    # Run processors in-place on all worksheets (names, gender, dates, identifiers).
    for ws in wb.worksheets:
        orch.process_worksheet(ws)

    wb.save(str(out_path))
    print(f"Wrote processor-only normalized workbook to {out_path}")


if __name__ == "__main__":
    main()

