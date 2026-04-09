"""Unit tests for response_parser module.

Tests the response parsing and validation functionality.
"""

import pytest
import json
from src.response_parser import ResponseParser
from src.exceptions import ValidationError


class TestResponseParser:
    """Test suite for ResponseParser class."""
    
    def test_parse_valid_json_response(self):
        """Test parsing a valid JSON response."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "read",
                    "target_file": "data/sales.xlsx",
                    "parameters": {}
                }
            ],
            "explanation": "Reading sales data"
        })
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert 'actions' in result
        assert len(result['actions']) == 1
        assert result['actions'][0]['tool'] == 'excel'
        assert result['explanation'] == 'Reading sales data'
    
    def test_parse_json_with_markdown_code_block(self):
        """Test parsing JSON wrapped in markdown code block."""
        # Arrange
        response_text = """```json
{
    "actions": [
        {
            "tool": "word",
            "operation": "create",
            "target_file": "report.docx",
            "parameters": {"content": "Test"}
        }
    ],
    "explanation": "Creating document"
}
```"""
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert 'actions' in result
        assert result['actions'][0]['tool'] == 'word'
    
    def test_parse_json_embedded_in_text(self):
        """Test parsing JSON embedded in surrounding text."""
        # Arrange
        response_text = """Here is the action:
{
    "actions": [
        {
            "tool": "powerpoint",
            "operation": "add",
            "target_file": "presentation.pptx",
            "parameters": {}
        }
    ],
    "explanation": "Adding slide"
}
That's what I'll do."""
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert 'actions' in result
        assert result['actions'][0]['tool'] == 'powerpoint'
    
    def test_parse_response_with_multiple_actions(self):
        """Test parsing response with multiple actions."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "read",
                    "target_file": "data.xlsx",
                    "parameters": {}
                },
                {
                    "tool": "word",
                    "operation": "create",
                    "target_file": "report.docx",
                    "parameters": {"content": "Report"}
                }
            ],
            "explanation": "Reading and creating"
        })
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert len(result['actions']) == 2
        assert result['actions'][0]['tool'] == 'excel'
        assert result['actions'][1]['tool'] == 'word'
    
    def test_parse_response_missing_actions_field(self):
        """Test that missing 'actions' field raises ValidationError."""
        # Arrange
        response_text = json.dumps({
            "explanation": "No actions provided"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "missing required field: 'actions'" in str(exc_info.value)
    
    def test_parse_response_empty_actions_list(self):
        """Test that empty actions list raises ValidationError."""
        # Arrange
        response_text = json.dumps({
            "actions": [],
            "explanation": "Empty actions"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "cannot be empty" in str(exc_info.value)
    
    def test_parse_response_invalid_tool(self):
        """Test that invalid tool name raises ValidationError."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "invalid_tool",
                    "operation": "read",
                    "target_file": "file.txt",
                    "parameters": {}
                }
            ],
            "explanation": "Invalid tool"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "invalid tool" in str(exc_info.value).lower()
    
    def test_parse_response_invalid_operation(self):
        """Test that invalid operation raises ValidationError."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "invalid_op",
                    "target_file": "file.xlsx",
                    "parameters": {}
                }
            ],
            "explanation": "Invalid operation"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "invalid operation" in str(exc_info.value).lower()
    
    def test_parse_response_missing_target_file(self):
        """Test that missing target_file raises ValidationError."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "read",
                    "parameters": {}
                }
            ],
            "explanation": "Missing target file"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "missing required field: 'target_file'" in str(exc_info.value)
    
    def test_parse_response_missing_parameters(self):
        """Test that missing parameters field raises ValidationError."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "read",
                    "target_file": "file.xlsx"
                }
            ],
            "explanation": "Missing parameters"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "missing required field: 'parameters'" in str(exc_info.value)
    
    def test_parse_free_text_with_excel_mention(self):
        """Test free-text parsing when Excel is mentioned."""
        # Arrange
        response_text = "I will read the Excel file data/sales.xlsx to get the information."
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert 'actions' in result
        assert len(result['actions']) > 0
        assert result['actions'][0]['tool'] == 'excel'
        assert result['actions'][0]['operation'] == 'read'
    
    def test_parse_free_text_with_word_mention(self):
        """Test free-text parsing when Word is mentioned."""
        # Arrange
        response_text = "I will create a new Word document at reports/summary.docx."
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert result['actions'][0]['tool'] == 'word'
        assert result['actions'][0]['operation'] == 'create'
    
    def test_parse_free_text_with_powerpoint_mention(self):
        """Test free-text parsing when PowerPoint is mentioned."""
        # Arrange
        response_text = "I will update the PowerPoint presentation slides.pptx."
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert result['actions'][0]['tool'] == 'powerpoint'
        assert result['actions'][0]['operation'] == 'update'
    
    def test_parse_free_text_insufficient_information(self):
        """Test that free-text parsing fails when information is insufficient."""
        # Arrange
        response_text = "I will do something with files."
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "Could not extract sufficient action information" in str(exc_info.value)
    
    def test_extract_actions(self):
        """Test extracting actions from parsed response."""
        # Arrange
        parsed_response = {
            "actions": [
                {"tool": "excel", "operation": "read", "target_file": "test.xlsx", "parameters": {}}
            ],
            "explanation": "Test"
        }
        
        # Act
        actions = ResponseParser.extract_actions(parsed_response)
        
        # Assert
        assert len(actions) == 1
        assert actions[0]['tool'] == 'excel'
    
    def test_extract_explanation(self):
        """Test extracting explanation from parsed response."""
        # Arrange
        parsed_response = {
            "actions": [],
            "explanation": "This is the explanation"
        }
        
        # Act
        explanation = ResponseParser.extract_explanation(parsed_response)
        
        # Assert
        assert explanation == "This is the explanation"
    
    def test_extract_explanation_missing(self):
        """Test extracting explanation when it's missing."""
        # Arrange
        parsed_response = {
            "actions": []
        }
        
        # Act
        explanation = ResponseParser.extract_explanation(parsed_response)
        
        # Assert
        assert explanation == ""
    
    def test_parse_response_with_complex_parameters(self):
        """Test parsing response with complex parameters."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "update",
                    "target_file": "budget.xlsx",
                    "parameters": {
                        "sheet": "Sheet1",
                        "row": 5,
                        "col": 3,
                        "value": 1500
                    }
                }
            ],
            "explanation": "Updating cell"
        })
        
        # Act
        result = ResponseParser.parse_response(response_text)
        
        # Assert
        assert result['actions'][0]['parameters']['sheet'] == 'Sheet1'
        assert result['actions'][0]['parameters']['row'] == 5
        assert result['actions'][0]['parameters']['col'] == 3
        assert result['actions'][0]['parameters']['value'] == 1500
    
    def test_parse_response_actions_not_list(self):
        """Test that actions field must be a list."""
        # Arrange
        response_text = json.dumps({
            "actions": "not a list",
            "explanation": "Invalid"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "must be a list" in str(exc_info.value)
    
    def test_parse_response_action_not_dict(self):
        """Test that each action must be a dictionary."""
        # Arrange
        response_text = json.dumps({
            "actions": ["not a dict"],
            "explanation": "Invalid"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "must be a dictionary" in str(exc_info.value)
    
    def test_parse_response_empty_target_file(self):
        """Test that target_file cannot be empty string."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "read",
                    "target_file": "",
                    "parameters": {}
                }
            ],
            "explanation": "Empty target"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "invalid target_file" in str(exc_info.value).lower()
    
    def test_parse_response_parameters_not_dict(self):
        """Test that parameters must be a dictionary."""
        # Arrange
        response_text = json.dumps({
            "actions": [
                {
                    "tool": "excel",
                    "operation": "read",
                    "target_file": "file.xlsx",
                    "parameters": "not a dict"
                }
            ],
            "explanation": "Invalid params"
        })
        
        # Act & Assert
        with pytest.raises(ValidationError) as exc_info:
            ResponseParser.parse_response(response_text)
        
        assert "invalid parameters" in str(exc_info.value).lower()
