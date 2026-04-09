"""Tests for enhanced error handling in ExcelTool."""

import pytest
import tempfile
import os
from pathlib import Path
from src.excel_tool import ExcelTool


class TestExcelToolErrorHandling:
    """Test enhanced error handling in ExcelTool."""
    
    def setup_method(self):
        """Setup test fixtures."""
        self.tool = ExcelTool()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test.xlsx")
    
    def teardown_method(self):
        """Cleanup test files."""
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        os.rmdir(self.temp_dir)
    
    def test_update_range_with_missing_value_field(self):
        """Test that update_range with missing 'value' field gives clear error."""
        # Create a test file
        self.tool.create_excel(self.test_file, {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3]
            ]
        })
        
        # Try to update with missing 'value' field
        invalid_updates = [
            {"row": 1, "col": 1}  # Missing 'value'
        ]
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.update_range(self.test_file, "Sheet1", invalid_updates)
        
        error_msg = str(exc_info.value).lower()
        assert "missing required field 'value'" in error_msg
        assert "expected format" in error_msg
        assert "{'row': int, 'col': int, 'value': any}" in str(exc_info.value)
    
    def test_update_range_with_missing_row_field(self):
        """Test that update_range with missing 'row' field gives clear error."""
        # Create a test file
        self.tool.create_excel(self.test_file, {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3]
            ]
        })
        
        # Try to update with missing 'row' field
        invalid_updates = [
            {"col": 1, "value": "test"}  # Missing 'row'
        ]
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.update_range(self.test_file, "Sheet1", invalid_updates)
        
        error_msg = str(exc_info.value).lower()
        assert "missing required field 'row'" in error_msg
    
    def test_update_range_with_missing_col_field(self):
        """Test that update_range with missing 'col' field gives clear error."""
        # Create a test file
        self.tool.create_excel(self.test_file, {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3]
            ]
        })
        
        # Try to update with missing 'col' field
        invalid_updates = [
            {"row": 1, "value": "test"}  # Missing 'col'
        ]
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.update_range(self.test_file, "Sheet1", invalid_updates)
        
        error_msg = str(exc_info.value).lower()
        assert "missing required field 'col'" in error_msg
    
    def test_update_range_with_valid_updates(self):
        """Test that update_range with valid updates works correctly."""
        # Create a test file
        self.tool.create_excel(self.test_file, {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3]
            ]
        })
        
        # Valid updates
        valid_updates = [
            {"row": 1, "col": 1, "value": "Updated A"},
            {"row": 1, "col": 2, "value": "Updated B"},
            {"row": 2, "col": 1, "value": 100}
        ]
        
        # Should not raise any exception
        self.tool.update_range(self.test_file, "Sheet1", valid_updates)
        
        # Verify updates were applied
        data = self.tool.read_excel(self.test_file)
        assert data['sheets']['Sheet1'][0][0] == "Updated A"
        assert data['sheets']['Sheet1'][0][1] == "Updated B"
        assert data['sheets']['Sheet1'][1][0] == 100
    
    def test_update_range_error_includes_position(self):
        """Test that error message includes the position of the invalid update."""
        # Create a test file
        self.tool.create_excel(self.test_file, {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3]
            ]
        })
        
        # Multiple updates, second one is invalid
        invalid_updates = [
            {"row": 1, "col": 1, "value": "Valid"},
            {"row": 2, "col": 2},  # Missing 'value' at position 1
            {"row": 3, "col": 3, "value": "Also Valid"}
        ]
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.update_range(self.test_file, "Sheet1", invalid_updates)
        
        error_msg = str(exc_info.value)
        assert "position 1" in error_msg.lower()
    
    def test_update_range_error_shows_received_data(self):
        """Test that error message shows what was actually received."""
        # Create a test file
        self.tool.create_excel(self.test_file, {
            "Sheet1": [
                ["A", "B", "C"],
                [1, 2, 3]
            ]
        })
        
        # Invalid update
        invalid_update = {"row": 1, "col": 1, "extra_field": "unexpected"}
        
        with pytest.raises(ValueError) as exc_info:
            self.tool.update_range(self.test_file, "Sheet1", [invalid_update])
        
        error_msg = str(exc_info.value)
        assert "received:" in error_msg.lower()
        assert str(invalid_update) in error_msg


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
