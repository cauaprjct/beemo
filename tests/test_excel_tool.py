"""Unit tests for ExcelTool class.

Tests cover basic operations, edge cases, and error handling for Excel file
manipulation.
"""

import os
import tempfile
from pathlib import Path

import pytest
from openpyxl import Workbook

from src.excel_tool import ExcelTool
from src.exceptions import CorruptedFileError


class TestExcelTool:
    """Test suite for ExcelTool class."""

    @pytest.fixture
    def excel_tool(self):
        """Fixture providing an ExcelTool instance."""
        return ExcelTool()

    @pytest.fixture
    def temp_dir(self):
        """Fixture providing a temporary directory for test files."""
        with tempfile.TemporaryDirectory() as tmpdir:
            yield tmpdir

    def test_create_and_read_excel(self, excel_tool, temp_dir):
        """Test creating an Excel file and reading it back."""
        file_path = os.path.join(temp_dir, "test.xlsx")
        data = {
            "Sheet1": [
                ["Name", "Age", "City"],
                ["Alice", 30, "New York"],
                ["Bob", 25, "London"]
            ]
        }

        # Create file
        excel_tool.create_excel(file_path, data)
        assert os.path.exists(file_path)

        # Read file
        result = excel_tool.read_excel(file_path)
        assert "sheets" in result
        assert "metadata" in result
        assert "Sheet1" in result["sheets"]
        assert result["sheets"]["Sheet1"] == data["Sheet1"]

    def test_create_excel_multiple_sheets(self, excel_tool, temp_dir):
        """Test creating an Excel file with multiple sheets."""
        file_path = os.path.join(temp_dir, "multi_sheet.xlsx")
        data = {
            "Sales": [["Product", "Revenue"], ["A", 100], ["B", 200]],
            "Expenses": [["Category", "Amount"], ["Rent", 500], ["Utilities", 100]]
        }

        excel_tool.create_excel(file_path, data)
        result = excel_tool.read_excel(file_path)

        assert len(result["sheets"]) == 2
        assert "Sales" in result["sheets"]
        assert "Expenses" in result["sheets"]
        assert result["sheets"]["Sales"] == data["Sales"]
        assert result["sheets"]["Expenses"] == data["Expenses"]

    def test_update_cell(self, excel_tool, temp_dir):
        """Test updating a specific cell in an Excel file."""
        file_path = os.path.join(temp_dir, "update_test.xlsx")
        data = {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3],
                [4, 5, 6]
            ]
        }

        # Create file
        excel_tool.create_excel(file_path, data)

        # Update cell
        excel_tool.update_cell(file_path, "Sheet1", 2, 2, 999)

        # Verify update
        result = excel_tool.read_excel(file_path)
        assert result["sheets"]["Sheet1"][1][1] == 999

    def test_update_cell_with_string(self, excel_tool, temp_dir):
        """Test updating a cell with a string value."""
        file_path = os.path.join(temp_dir, "update_string.xlsx")
        data = {"Sheet1": [["Old Value"]]}

        excel_tool.create_excel(file_path, data)
        excel_tool.update_cell(file_path, "Sheet1", 1, 1, "New Value")

        result = excel_tool.read_excel(file_path)
        assert result["sheets"]["Sheet1"][0][0] == "New Value"

    def test_add_sheet(self, excel_tool, temp_dir):
        """Test adding a new sheet to an existing Excel file."""
        file_path = os.path.join(temp_dir, "add_sheet_test.xlsx")
        initial_data = {"Sheet1": [["Initial", "Data"]]}

        # Create file with one sheet
        excel_tool.create_excel(file_path, initial_data)

        # Add new sheet
        new_sheet_data = [["New", "Sheet"], ["With", "Data"]]
        excel_tool.add_sheet(file_path, "Sheet2", new_sheet_data)

        # Verify both sheets exist
        result = excel_tool.read_excel(file_path)
        assert len(result["sheets"]) == 2
        assert "Sheet1" in result["sheets"]
        assert "Sheet2" in result["sheets"]
        assert result["sheets"]["Sheet2"] == new_sheet_data

    def test_read_excel_file_not_found(self, excel_tool):
        """Test reading a non-existent Excel file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError, match="Excel file not found"):
            excel_tool.read_excel("nonexistent.xlsx")

    def test_update_cell_file_not_found(self, excel_tool):
        """Test updating a cell in a non-existent file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError, match="Excel file not found"):
            excel_tool.update_cell("nonexistent.xlsx", "Sheet1", 1, 1, "value")

    def test_add_sheet_file_not_found(self, excel_tool):
        """Test adding a sheet to a non-existent file raises FileNotFoundError."""
        with pytest.raises(FileNotFoundError, match="Excel file not found"):
            excel_tool.add_sheet("nonexistent.xlsx", "NewSheet", [])

    def test_update_cell_invalid_sheet(self, excel_tool, temp_dir):
        """Test updating a cell in a non-existent sheet raises ValueError."""
        file_path = os.path.join(temp_dir, "test.xlsx")
        excel_tool.create_excel(file_path, {"Sheet1": [["Data"]]})

        with pytest.raises(ValueError, match="Sheet 'InvalidSheet' not found"):
            excel_tool.update_cell(file_path, "InvalidSheet", 1, 1, "value")

    def test_add_sheet_duplicate_name(self, excel_tool, temp_dir):
        """Test adding a sheet with a duplicate name raises ValueError."""
        file_path = os.path.join(temp_dir, "test.xlsx")
        excel_tool.create_excel(file_path, {"Sheet1": [["Data"]]})

        with pytest.raises(ValueError, match="Sheet 'Sheet1' already exists"):
            excel_tool.add_sheet(file_path, "Sheet1", [["New Data"]])

    def test_read_corrupted_file(self, excel_tool, temp_dir):
        """Test reading a corrupted Excel file raises CorruptedFileError."""
        file_path = os.path.join(temp_dir, "corrupted.xlsx")
        
        # Create a file with invalid content
        with open(file_path, "w") as f:
            f.write("This is not a valid Excel file")

        with pytest.raises(CorruptedFileError, match="corrupted or invalid"):
            excel_tool.read_excel(file_path)

    def test_create_excel_empty_data(self, excel_tool, temp_dir):
        """Test creating an Excel file with empty data."""
        file_path = os.path.join(temp_dir, "empty.xlsx")
        data = {"EmptySheet": []}

        excel_tool.create_excel(file_path, data)
        result = excel_tool.read_excel(file_path)

        assert "EmptySheet" in result["sheets"]
        assert result["sheets"]["EmptySheet"] == []

    def test_create_excel_with_various_data_types(self, excel_tool, temp_dir):
        """Test creating an Excel file with various data types."""
        file_path = os.path.join(temp_dir, "types.xlsx")
        data = {
            "Sheet1": [
                ["String", "Integer", "Float", "None"],
                ["text", 42, 3.14, None],
                ["another", -10, 0.0, None]
            ]
        }

        excel_tool.create_excel(file_path, data)
        result = excel_tool.read_excel(file_path)

        # Verify data types are preserved
        assert result["sheets"]["Sheet1"][1][0] == "text"
        assert result["sheets"]["Sheet1"][1][1] == 42
        assert abs(result["sheets"]["Sheet1"][1][2] - 3.14) < 0.01
        assert result["sheets"]["Sheet1"][1][3] is None

    def test_read_excel_metadata(self, excel_tool, temp_dir):
        """Test that read_excel returns metadata."""
        file_path = os.path.join(temp_dir, "metadata_test.xlsx")
        data = {"Sheet1": [["Data"]]}

        excel_tool.create_excel(file_path, data)
        result = excel_tool.read_excel(file_path)

        assert "metadata" in result
        assert "sheet_count" in result["metadata"]
        assert "sheet_names" in result["metadata"]
        assert result["metadata"]["sheet_count"] == 1
        assert result["metadata"]["sheet_names"] == ["Sheet1"]

    def test_update_cell_preserves_other_data(self, excel_tool, temp_dir):
        """Test that updating a cell preserves other data in the sheet."""
        file_path = os.path.join(temp_dir, "preserve_test.xlsx")
        data = {
            "Sheet1": [
                ["A", "B", "C"],
                ["D", "E", "F"],
                ["G", "H", "I"]
            ]
        }

        excel_tool.create_excel(file_path, data)
        excel_tool.update_cell(file_path, "Sheet1", 2, 2, "UPDATED")

        result = excel_tool.read_excel(file_path)
        
        # Check updated cell
        assert result["sheets"]["Sheet1"][1][1] == "UPDATED"
        
        # Check other cells are preserved
        assert result["sheets"]["Sheet1"][0][0] == "A"
        assert result["sheets"]["Sheet1"][2][2] == "I"

    def test_add_sheet_preserves_existing_sheets(self, excel_tool, temp_dir):
        """Test that adding a sheet preserves existing sheets."""
        file_path = os.path.join(temp_dir, "preserve_sheets.xlsx")
        initial_data = {
            "Sheet1": [["Original", "Data"]],
            "Sheet2": [["More", "Original"]]
        }

        excel_tool.create_excel(file_path, initial_data)
        excel_tool.add_sheet(file_path, "Sheet3", [["New", "Data"]])

        result = excel_tool.read_excel(file_path)
        
        # Check all sheets exist
        assert len(result["sheets"]) == 3
        assert result["sheets"]["Sheet1"] == initial_data["Sheet1"]
        assert result["sheets"]["Sheet2"] == initial_data["Sheet2"]
        assert result["sheets"]["Sheet3"] == [["New", "Data"]]
