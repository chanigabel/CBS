"""Command-line interface for the Excel standardization pipeline.

Accepts an Excel workbook path, runs the full JSON-based standardization
pipeline, and writes the corrected output to a new ``_normalized.xlsx`` file.
The original input file is never modified.
"""

import argparse
import logging
import sys
from pathlib import Path
from datetime import datetime
from typing import NoReturn

from .orchestrator import standardizationOrchestrator


def setup_logging(file_path: str) -> None:
    """Configure logging with console INFO+ and file DEBUG+."""

    input_path = Path(file_path)
    log_dir = input_path.parent
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"standardization_{timestamp}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(file_formatter)
    logger.addHandler(file_handler)

    logger.info(f"Logging initialized. Log file: {log_file}")


def parse_arguments() -> argparse.Namespace:
    """Parse CLI arguments."""

    parser = argparse.ArgumentParser(
        description="Normalize Excel data files with person records"
    )

    parser.add_argument(
        "file_path",
        type=str,
        help="Path to the Excel workbook file to process"
    )

    return parser.parse_args()


def validate_file_path(file_path: str) -> None:
    """Validate input Excel file."""

    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if not path.is_file():
        raise ValueError(f"Path is not a file: {file_path}")

    if path.suffix.lower() not in [".xlsx", ".xlsm"]:
        raise ValueError(
            f"Invalid file format: {path.suffix}. Expected .xlsx or .xlsm"
        )

    if not path.stat().st_mode & 0o400:
        raise PermissionError(f"File is not readable: {file_path}")


def build_output_path(input_file: str) -> str:
    """Create output Excel path without modifying original file."""

    input_path = Path(input_file)

    output_path = input_path.with_name(
        input_path.stem + "_normalized.xlsx"
    )

    return str(output_path)


def main() -> int:
    """Main CLI entry point."""

    try:

        args = parse_arguments()
        file_path = args.file_path

        validate_file_path(file_path)

        setup_logging(file_path)
        logger = logging.getLogger(__name__)

        logger.info("=" * 80)
        logger.info("Excel Data standardization - Starting")
        logger.info("=" * 80)
        logger.info(f"Input file: {file_path}")

        # Build output file path
        output_excel_path = build_output_path(file_path)

        logger.info(f"Output file will be written to: {output_excel_path}")

        orchestrator = standardizationOrchestrator()

        logger.info("Starting workbook standardization...")

        orchestrator.process_workbook_json(
            file_path,
            output_excel_path
        )

        logger.info("=" * 80)
        logger.info("Excel Data standardization - Completed Successfully")
        logger.info("=" * 80)

        return 0

    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    except PermissionError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    except ValueError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1

    except Exception as e:
        try:
            logger = logging.getLogger(__name__)
            logger.error(
                f"Unexpected error during standardization: {e}",
                exc_info=True,
            )
        except Exception:
            print(f"ERROR: Unexpected error: {e}", file=sys.stderr)
        return 1


def cli_entry_point() -> NoReturn:
    """Console script entry point."""
    sys.exit(main())


if __name__ == "__main__":
    sys.exit(main())
