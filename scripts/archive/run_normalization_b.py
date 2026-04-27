from pathlib import Path

from openpyxl import load_workbook

from src.excel_normalization.orchestrator import NormalizationOrchestrator


def main() -> None:
    base = Path("B.xlsx").resolve()
    if not base.exists():
        raise SystemExit(f"B.xlsx not found at {base}")

    out_path = Path("B_normalized.xlsx").resolve()

    # Use the worksheet-based processors path to mirror VBA normalization.
    orch = NormalizationOrchestrator()
    orch.export_vba_parity_workbook_from_processors(str(base), str(out_path))
    print(f"Wrote normalized workbook to {out_path}")


if __name__ == "__main__":
    main()

