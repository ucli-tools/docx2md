"""
Tests for the command-line interface.
"""

import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import click
import pytest
from click.testing import CliRunner

from docx2md.cli import main


class TestCLI(unittest.TestCase):
    """Test cases for the command-line interface."""
    
    def setUp(self):
        """Set up test fixtures."""
        # Create a CLI runner
        self.runner = CliRunner()
        
        # Create a temporary directory for test files
        self.temp_dir = tempfile.TemporaryDirectory()
        self.test_dir = Path(self.temp_dir.name)
    
    def tearDown(self):
        """Tear down test fixtures."""
        # Clean up temporary directory
        self.temp_dir.cleanup()
    
    def test_help(self):
        """Test the help command."""
        result = self.runner.invoke(main, ["--help"])
        
        # Check that the command succeeded
        self.assertEqual(result.exit_code, 0)
        
        # Check that the help text contains expected sections
        self.assertIn("docx2md: Convert Microsoft Word documents", result.output)
        self.assertIn("--help", result.output)
        self.assertIn("--version", result.output)
    
    def test_version(self):
        """Test the version command."""
        result = self.runner.invoke(main, ["--version"])
        
        # Check that the command succeeded
        self.assertEqual(result.exit_code, 0)
        
        # Check that the version is displayed
        self.assertIn("docx2md, version", result.output)
    
    @patch("docx2md.cli.convert_docx_to_markdown")
    def test_convert_command(self, mock_convert):
        """Test the convert command."""
        # Create a dummy DOCX file
        input_file = self.test_dir / "test.docx"
        with open(input_file, "w") as f:
            f.write("Dummy DOCX content")
        
        # Set up the mock to return a success result
        mock_convert.return_value = {
            "input_file": str(input_file),
            "output_file": str(self.test_dir / "test.md"),
            "equations_count": 5,
            "images_count": 2,
            "tables_count": 1,
        }
        
        # Run the convert command
        result = self.runner.invoke(main, ["convert", str(input_file)])
        
        # Check that the command succeeded
        self.assertEqual(result.exit_code, 0)
        
        # Check that the mock was called with the correct arguments
        mock_convert.assert_called_once()
        args, kwargs = mock_convert.call_args
        self.assertEqual(str(args[0]), str(input_file))
    
    @patch("docx2md.cli.batch_convert")
    def test_batch_command(self, mock_batch):
        """Test the batch command."""
        # Create a dummy input directory
        input_dir = self.test_dir / "input"
        input_dir.mkdir()
        
        # Create a dummy output directory
        output_dir = self.test_dir / "output"
        
        # Set up the mock to return a success result
        mock_batch.return_value = {
            "files_processed": 3,
            "files_succeeded": 3,
            "files_failed": 0,
            "results": [],
        }
        
        # Run the batch command
        result = self.runner.invoke(main, ["batch", str(input_dir), str(output_dir)])
        
        # Check that the command succeeded
        self.assertEqual(result.exit_code, 0)
        
        # Check that the mock was called with the correct arguments
        mock_batch.assert_called_once()
        args, kwargs = mock_batch.call_args
        self.assertEqual(str(args[0]), str(input_dir))
        self.assertEqual(str(args[1]), str(output_dir))
    
    @patch("docx2md.cli.fix_delimiters")
    def test_fix_delimiters_command(self, mock_fix):
        """Test the fix-delimiters command."""
        # Create a dummy Markdown file
        input_file = self.test_dir / "test.md"
        with open(input_file, "w") as f:
            f.write("Dummy Markdown content")
        
        # Set up the mock to return a success result
        mock_fix.return_value = {
            "inline_original": 2,
            "display_original": 1,
            "inline_fixed": 2,
            "display_fixed": 1,
        }
        
        # Run the fix-delimiters command
        result = self.runner.invoke(main, ["fix-delimiters", str(input_file)])
        
        # Check that the command succeeded
        self.assertEqual(result.exit_code, 0)
        
        # Check that the mock was called with the correct arguments
        mock_fix.assert_called_once()
        args, kwargs = mock_fix.call_args
        self.assertEqual(str(args[0]), str(input_file))
    
    @patch("docx2md.processors.equations.validate_equations")
    def test_validate_command(self, mock_validate):
        """Test the validate command."""
        # Create a dummy Markdown file
        input_file = self.test_dir / "test.md"
        with open(input_file, "w") as f:
            f.write("Dummy Markdown content")
        
        # Set up the mock to return a success result
        mock_validate.return_value = {
            "is_valid": True,
            "inline_count": 2,
            "display_count": 1,
            "invalid_count": 0,
            "issues": [],
        }
        
        # Run the validate command
        result = self.runner.invoke(main, ["validate", str(input_file)])
        
        # Check that the command succeeded
        self.assertEqual(result.exit_code, 0)
        
        # Check that the mock was called with the correct arguments
        mock_validate.assert_called_once()
        args, kwargs = mock_validate.call_args
        self.assertEqual(str(args[0]), str(input_file))
    
    @patch("docx2md.processors.equations.validate_equations")
    def test_validate_command_failure(self, mock_validate):
        """Test the validate command with validation failure."""
        # Create a dummy Markdown file
        input_file = self.test_dir / "test.md"
        with open(input_file, "w") as f:
            f.write("Dummy Markdown content")
        
        # Set up the mock to return a failure result
        mock_validate.return_value = {
            "is_valid": False,
            "inline_count": 2,
            "display_count": 1,
            "invalid_count": 1,
            "issues": ["Invalid equation: $E = mc^2 + {$"],
        }
        
        # Run the validate command
        result = self.runner.invoke(main, ["validate", str(input_file)])
        
        # Check that the command failed
        self.assertEqual(result.exit_code, 1)
        
        # Check that the mock was called with the correct arguments
        mock_validate.assert_called_once()
        args, kwargs = mock_validate.call_args
        self.assertEqual(str(args[0]), str(input_file))


if __name__ == "__main__":
    unittest.main()