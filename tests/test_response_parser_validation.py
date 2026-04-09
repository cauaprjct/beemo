"""Tests for enhanced response parser validation."""

import pytest
from src.response_parser import ResponseParser
from src.exceptions import ValidationError


class TestResponseParserValidation:
    """Test enhanced validation in ResponseParser."""
    
    def test_validate_update_with_missing_value(self):
        """Test that update operation with missing 'value' is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "updates": [
                            {"row": 1, "col": 1}  # Missing 'value'
                        ]
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "missing required field 'value'" in str(exc_info.value).lower()
        assert "expected format" in str(exc_info.value).lower()
    
    def test_validate_update_with_invalid_row(self):
        """Test that update operation with invalid row is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "updates": [
                            {"row": 0, "col": 1, "value": "test"}  # Row must be >= 1
                        ]
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "invalid 'row'" in str(exc_info.value).lower()
    
    def test_validate_update_with_empty_updates_list(self):
        """Test that update operation with empty updates list is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "updates": []  # Empty list
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "cannot be empty" in str(exc_info.value).lower()
    
    def test_validate_update_with_valid_updates(self):
        """Test that valid update operation passes validation."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "updates": [
                            {"row": 1, "col": 1, "value": "test1"},
                            {"row": 2, "col": 2, "value": "test2"}
                        ]
                    }
                }
            ],
            "explanation": "Test"
        }
        
        # Should not raise any exception
        ResponseParser._validate_structure(response)
    
    def test_validate_format_without_range(self):
        """Test that format operation without 'range' is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "format",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "formatting": {
                            "bold": True
                            # Missing 'range'
                        }
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "range" in str(exc_info.value).lower()
    
    def test_validate_format_with_valid_range(self):
        """Test that valid format operation passes validation."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "format",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "formatting": {
                            "range": "A1:C10",
                            "bold": True
                        }
                    }
                }
            ],
            "explanation": "Test"
        }
        
        # Should not raise any exception
        ResponseParser._validate_structure(response)
    
    def test_validate_formula_with_missing_fields(self):
        """Test that formula operation with missing fields is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "formula",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "formulas": [
                            {"row": 1, "col": 1}  # Missing 'formula'
                        ]
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "missing required field 'formula'" in str(exc_info.value).lower()
    
    def test_validate_append_without_rows(self):
        """Test that append operation without 'rows' is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "append",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1"
                        # Missing 'rows'
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "requires 'rows' parameter" in str(exc_info.value).lower()
    
    def test_validate_merge_with_insufficient_files(self):
        """Test that PDF merge with less than 2 files is rejected."""
        response = {
            "actions": [
                {
                    "tool": "pdf",
                    "operation": "merge",
                    "target_file": "output.pdf",
                    "parameters": {
                        "file_paths": ["file1.pdf"]  # Only 1 file
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "at least 2 files" in str(exc_info.value).lower()
    
    def test_validate_single_cell_update(self):
        """Test that single cell update with all required fields passes."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "row": 1,
                        "col": 1,
                        "value": "test"
                    }
                }
            ],
            "explanation": "Test"
        }
        
        # Should not raise any exception
        ResponseParser._validate_structure(response)
    
    def test_validate_single_cell_update_missing_value(self):
        """Test that single cell update without 'value' is rejected."""
        response = {
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "test.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "row": 1,
                        "col": 1
                        # Missing 'value'
                    }
                }
            ],
            "explanation": "Test"
        }
        
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser._validate_structure(response)
        
        assert "missing required field 'value'" in str(exc_info.value).lower()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
