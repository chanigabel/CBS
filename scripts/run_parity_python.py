from pathlib import Path

from src.excel_standardization.orchestrator import standardizationOrchestrator


def main() -> None:
    src = Path("A.xlsx").resolve()
    out = Path("python_vs_vba_python.xlsx").resolve()
    if not src.exists():
        raise SystemExit(f"Missing source workbook: {src}")

    standardizationOrchestrator().process_workbook_json(str(src), str(out))
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()

