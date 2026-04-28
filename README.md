# Excel Data standardization System

A Python-based Excel data standardization system that replicates the exact behavior of a legacy VBA implementation. The system processes Excel workbooks containing person records (residents, staff, etc.) from various sources, standardizing inconsistent data into a standardized format.

## Overview

This system transforms person records with inconsistent structures, mixed languages (Hebrew/English), varying date formats, and invalid identifiers. It reads the original workbook without modifying it, runs the standardization pipeline, and writes a clean output file with standardized corrected values in a fixed column schema.

### Key Features

- **Rule-based and deterministic**: No machine learning or probabilistic algorithms
- **Non-destructive**: The original input file is never modified; output is written to a new file
- **Fixed output schema**: Exports a clean workbook with standardized column names (Hebrew field names)
- **Hebrew status messages**: Uses exact same status text as VBA system
- **Comprehensive logging**: Console and file logging for troubleshooting

### Processing Scope

The system normalizes four categories of data:

1. **Names**: First name, last name, father's name
   - Text standardization with language detection
   - Diacritic removal
   - Father name pattern removal

2. **Gender**: Various representations normalized to 1 (male) or 2 (female)

3. **Dates**: Birth date and entry date
   - Split columns (year, month, day)
   - Format detection (DDMM vs MMDD)
   - Business rule validation (age, future dates)

4. **Identifiers**: Israeli ID and passport numbers
   - Checksum validation for Israeli IDs
   - Character cleaning for passports

## Installation

### Requirements

- Python 3.9 or higher
- pip (Python package installer)

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Development Installation

For development with testing and linting tools:

```bash
pip install -e ".[dev]"
```

## Web Application

The easiest way to use this system is through the local web application, which provides a browser-based UI for uploading, standardizing, editing, and downloading Excel workbooks — no command line required.

### Start the Web App

```bash
uvicorn webapp.app:app --reload
```

Then open your browser and navigate to:

```
http://localhost:8000
```

### Web App Workflow

1. **Upload** — Select and upload an `.xlsx` or `.xlsm` file
2. **View** — Browse sheet data in the grid
3. **Normalize** — Click "Run standardization" to process all sheets
4. **Edit** — Click any cell to edit its value inline
5. **Export** — Click "Export / Download" to download the corrected workbook

The web app runs entirely locally — no internet connection required, no data leaves your machine.

## Usage

### Command Line Interface

Process an Excel workbook:

```bash
python -m excel_standardization.cli path/to/workbook.xlsx
```

Or if installed as a package:

```bash
excel-normalize path/to/workbook.xlsx
```

### Examples

```bash
# Process a workbook in the current directory
python -m excel_standardization.cli data.xlsx

# Process a workbook with full path
python -m excel_standardization.cli /path/to/workbook.xlsx
```

### Output

The system:
- **Never modifies the original input file** — it is read-only throughout
- Writes a new output file named `{input_stem}_normalized.xlsx` in the same directory
- The output uses a fixed column schema with standardized Hebrew field names
- Creates a log file in the same directory as the input file
- Log file format: `standardization_YYYYMMDD_HHMMSS.log`

## Architecture

The system follows a layered architecture with strict separation of concerns:

```
┌─────────────────────────────────────────────────────────────┐
│                      CLI Interface                          │
│                  (Argument parsing, logging)                │
└─────────────────────────────────────────────────────────────┘
                            │
┌─────────────────────────────────────────────────────────────┐
│                 standardizationOrchestrator                   │
│           (Coordinates processing across worksheets)        │
└─────────────────────────────────────────────────────────────┘
                            │
┌─────────────────────────────────────────────────────────────┐
│                    Field Processors                         │
│  (NameFieldProcessor, GenderFieldProcessor, etc.)          │
└─────────────────────────────────────────────────────────────┘
                            │
┌──────────────────────┬──────────────────────────────────────┐
│    I/O Layer         │        Business Logic Layer          │
│  (ExcelReader,       │  (NameEngine, DateEngine, etc.)     │
│   ExcelWriter)       │  (Pure functions, no Excel deps)    │
└──────────────────────┴──────────────────────────────────────┘
```

### Key Design Principles

1. **Strict Separation of Concerns**: Excel I/O operations are completely isolated from business logic
2. **Pure Functions**: Engine classes contain pure business logic with zero Excel dependencies
3. **Data Structure Abstraction**: Processors operate on plain Python data structures
4. **Template Method Pattern**: FieldProcessor provides a template for processing different field types
5. **Array-Based Operations**: Data is read as arrays, processed in memory, and written back

## Logging

The system provides comprehensive logging:

- **Console**: INFO level and above
- **File**: DEBUG level and above
- **Format**: `YYYY-MM-DD HH:MM:SS - module - LEVEL - message`

### Log Levels

- **ERROR**: File/worksheet failures that prevent processing
- **WARNING**: Unexpected conditions that don't prevent processing
- **INFO**: Processing milestones (worksheet start/complete, summary stats)
- **DEBUG**: Detailed processing information (header detection, pattern detection)

## Error Handling

The system handles errors gracefully:

### File-Level Errors (Fatal)
- File not found
- File not readable (permissions, corruption)
- Invalid Excel format

### Worksheet-Level Errors (Log and continue)
- Worksheet structure unexpected
- Header detection failures
- Column insertion failures

### Row-Level Errors (Log and continue)
- Invalid data values
- Parsing failures
- Validation failures

### Cell-Level Errors (Set status text)
- Date parsing errors → Hebrew status text
- ID validation errors → Hebrew status text

## Testing

Run the test suite:

```bash
# Run all tests
pytest

# Run with coverage report
pytest --cov=src/excel_standardization --cov-report=html

# Run specific test file
pytest tests/test_cli.py -v
```

### Test Coverage

The project includes:
- Unit tests for all engine classes
- Unit tests for I/O layer
- Integration tests for field processors
- Property-based tests for correctness properties
- CLI tests for error handling

## VBA Compatibility

This Python system replicates the exact behavior of the legacy VBA implementation. All standardization, validation, and transformation logic follows the same rules observed in the VBA code.

### Key Compatibility Points

- Same header matching logic (exact text with partial match)
- Same text standardization rules (diacritics, language detection, spacing)
- Same date parsing logic (format detection, two-digit year expansion)
- Same Israeli ID checksum algorithm
- Same status messages (in Hebrew)
- Same visual feedback (pink highlights for changes)

## Data Privacy and Security

⚠️ **Important**: This system processes personal data. Please ensure compliance with applicable privacy regulations:

- The system does not transmit data over networks
- All processing is done locally
- Original data is preserved in original columns
- Log files may contain personal information - handle appropriately
- Do not include sample data with real personal information in the repository

## Project Structure

```
excel-data-standardization/
├── src/
│   └── excel_standardization/
│       ├── __init__.py
│       ├── cli.py                    # CLI entry point
│       ├── orchestrator.py           # Orchestration layer
│       ├── data_types.py             # Data models and enums
│       ├── io_layer/
│       │   ├── excel_reader.py       # Excel reading operations
│       │   └── excel_writer.py       # Excel writing operations
│       ├── processing/
│       │   ├── field_processor.py    # Base class for processors
│       │   ├── name_processor.py     # Name field processing
│       │   ├── gender_processor.py   # Gender field processing
│       │   ├── date_processor.py     # Date field processing
│       │   └── identifier_processor.py  # ID/passport processing
│       └── engines/
│           ├── name_engine.py        # Name standardization logic
│           ├── text_processor.py     # Text manipulation utilities
│           ├── date_engine.py        # Date parsing and validation
│           ├── gender_engine.py      # Gender standardization logic
│           └── identifier_engine.py  # ID/passport validation logic
├── tests/
│   ├── test_cli.py
│   ├── test_date_engine.py
│   ├── test_gender_processor.py
│   ├── test_identifier_engine.py
│   └── ...
├── pyproject.toml
├── requirements.txt
└── README.md
```

## Development

### Code Style

The project uses:
- **black**: Code formatting
- **mypy**: Static type checking
- **flake8**: Linting

Run code quality checks:

```bash
# Format code
black src/ tests/

# Type checking
mypy src/

# Linting
flake8 src/ tests/
```

### Contributing

When contributing:
1. Write tests for new functionality
2. Ensure all tests pass
3. Maintain type hints
4. Follow existing code style
5. Update documentation

## License

[Add your license information here]

## Support

For issues or questions, please [add contact information or issue tracker link].
