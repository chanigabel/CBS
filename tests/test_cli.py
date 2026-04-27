"""Tests for CLI module."""

import pytest
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock
from io import StringIO

from src.excel_normalization.cli import (
    parse_arguments,
    validate_file_path,
    setup_logging,
    main
)


class TestValidateFilePath:
    """Tests for validate_file_path function."""
    
    def test_file_not_found(self):
        """Test that FileNotFoundError is raised for non-existent file."""
        with pytest.raises(FileNotFoundError, match="File not found"):
            validate_file_path("nonexistent_file.xlsx")
    
    def test_invalid_extension(self, tmp_path):
        """Test that ValueError is raised for non-Excel file."""
        # Create a text file
        test_file = tmp_path / "test.txt"
        test_file.write_text("test")
        
        with pytest.raises(ValueError, match="Invalid file format"):
            validate_file_path(str(test_file))
    
    def test_directory_instead_of_file(self, tmp_path):
        """Test that ValueError is raised when path is a directory."""
        with pytest.raises(ValueError, match="Path is not a file"):
            validate_file_path(str(tmp_path))
    
    def test_valid_xlsx_file(self, tmp_path):
        """Test that valid .xlsx file passes validation."""
        test_file = tmp_path / "test.xlsx"
        test_file.write_text("test")
        
        # Should not raise any exception
        validate_file_path(str(test_file))
    
    def test_valid_xlsm_file(self, tmp_path):
        """Test that valid .xlsm file passes validation."""
        test_file = tmp_path / "test.xlsm"
        test_file.write_text("test")
        
        # Should not raise any exception
        validate_file_path(str(test_file))


class TestSetupLogging:
    """Tests for setup_logging function."""
    
    def test_logging_setup(self, tmp_path):
        """Test that logging is configured correctly."""
        import logging
        
        # Create a test file path
        test_file = tmp_path / "test.xlsx"
        test_file.write_text("test")
        
        # Setup logging
        setup_logging(str(test_file))
        
        # Get root logger
        logger = logging.getLogger()
        
        # Check that logger has handlers
        assert len(logger.handlers) >= 2
        
        # Check that there's a console handler and file handler
        handler_types = [type(h).__name__ for h in logger.handlers]
        assert 'StreamHandler' in handler_types
        assert 'FileHandler' in handler_types
        
        # Check log level
        assert logger.level == logging.DEBUG


class TestMain:
    """Tests for main function."""
    
    @patch('src.excel_normalization.cli.NormalizationOrchestrator')
    def test_main_success(self, mock_orchestrator, tmp_path):
        """Test successful execution of main."""
        # Create a test Excel file
        test_file = tmp_path / "test.xlsx"
        test_file.write_text("test")
        
        expected_output = str(test_file.with_name("test_normalized.xlsx"))
        
        # Mock sys.argv
        with patch.object(sys, 'argv', ['cli.py', str(test_file)]):
            exit_code = main()
        
        # Check that orchestrator was called with the JSON-based pipeline method
        mock_orchestrator.assert_called_once()
        mock_orchestrator.return_value.process_workbook_json.assert_called_once_with(
            str(test_file), expected_output
        )
        
        # Check exit code
        assert exit_code == 0
    
    def test_main_file_not_found(self):
        """Test main with non-existent file."""
        with patch.object(sys, 'argv', ['cli.py', 'nonexistent.xlsx']):
            exit_code = main()
        
        # Check exit code
        assert exit_code == 1
    
    def test_main_invalid_format(self, tmp_path):
        """Test main with invalid file format."""
        test_file = tmp_path / "test.txt"
        test_file.write_text("test")
        
        with patch.object(sys, 'argv', ['cli.py', str(test_file)]):
            exit_code = main()
        
        # Check exit code
        assert exit_code == 1
    
    @patch('src.excel_normalization.cli.NormalizationOrchestrator')
    def test_main_unexpected_error(self, mock_orchestrator, tmp_path):
        """Test main with unexpected error during processing."""
        # Create a test Excel file
        test_file = tmp_path / "test.xlsx"
        test_file.write_text("test")
        
        # Make orchestrator raise an exception on the JSON-based pipeline method
        mock_orchestrator.return_value.process_workbook_json.side_effect = Exception("Test error")
        
        with patch.object(sys, 'argv', ['cli.py', str(test_file)]):
            exit_code = main()
        
        # Check exit code
        assert exit_code == 1


class TestParseArguments:
    """Tests for parse_arguments function."""
    
    def test_parse_file_path(self):
        """Test parsing file path argument."""
        with patch.object(sys, 'argv', ['cli.py', 'test.xlsx']):
            args = parse_arguments()
            assert args.file_path == 'test.xlsx'
    
    def test_parse_missing_argument(self):
        """Test that missing argument raises SystemExit."""
        with patch.object(sys, 'argv', ['cli.py']):
            with pytest.raises(SystemExit):
                parse_arguments()
