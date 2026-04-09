"""Testes para operações em lote."""

import pytest
from unittest.mock import Mock, patch, MagicMock

from src.models import AgentResponse, BatchResult, ActionResult
from src.agent import Agent


@pytest.fixture
def mock_config():
    """Mock configuration."""
    config = Mock()
    config.api_key = "test_api_key"
    config.model_name = "gemini-2.5-flash-lite"
    config.fallback_models = ["gemini-2.5-flash", "gemini-2.5-pro"]
    config.root_path = "/tmp/test"
    return config


@pytest.fixture
def agent(mock_config):
    """Create Agent with mocked dependencies."""
    with patch('src.agent.FileScanner'), \
         patch('src.agent.GeminiClient'), \
         patch('src.agent.ExcelTool'), \
         patch('src.agent.WordTool'), \
         patch('src.agent.PowerPointTool'), \
         patch('src.agent.SecurityValidator'):
        return Agent(mock_config)


def test_batch_result_structure():
    """Test BatchResult data structure."""
    action_results = [
        ActionResult(0, "excel", "update", "file1.xlsx", True, None),
        ActionResult(1, "excel", "update", "file2.xlsx", True, None),
        ActionResult(2, "excel", "update", "file3.xlsx", False, "Error message")
    ]
    
    batch = BatchResult(
        total_actions=3,
        successful_actions=2,
        failed_actions=1,
        action_results=action_results,
        overall_success=True
    )
    
    assert batch.total_actions == 3
    assert batch.successful_actions == 2
    assert batch.failed_actions == 1
    assert len(batch.action_results) == 3
    assert batch.overall_success is True


def test_agent_response_with_batch_result():
    """Test AgentResponse with batch_result field."""
    batch = BatchResult(
        total_actions=2,
        successful_actions=2,
        failed_actions=0,
        action_results=[],
        overall_success=True
    )
    
    response = AgentResponse(
        success=True,
        message="2 de 2 operações bem-sucedidas",
        files_modified=["file1.xlsx", "file2.xlsx"],
        error=None,
        batch_result=batch
    )
    
    assert response.batch_result is not None
    assert response.batch_result.total_actions == 2
    assert response.batch_result.successful_actions == 2


def test_batch_all_successful(agent):
    """Test batch operation where all actions succeed."""
    # Mock Gemini response with multiple actions
    gemini_response = """
    {
        "actions": [
            {"tool": "excel", "operation": "update", "target_file": "test1.xlsx", "parameters": {"sheet": "Sheet1", "row": 1, "col": 1, "value": 10}},
            {"tool": "excel", "operation": "update", "target_file": "test2.xlsx", "parameters": {"sheet": "Sheet1", "row": 1, "col": 1, "value": 20}}
        ],
        "explanation": "Updating cells"
    }
    """
    
    # Mock dependencies
    agent.security_validator.validate_operation = Mock()
    agent.security_validator.validate_file_path = Mock(side_effect=lambda x: x)
    agent._execute_single_action = Mock()
    
    response = agent._execute_actions(gemini_response)
    
    assert response.success is True
    assert response.batch_result is not None
    assert response.batch_result.total_actions == 2
    assert response.batch_result.successful_actions == 2
    assert response.batch_result.failed_actions == 0
    assert response.batch_result.overall_success is True


def test_batch_partial_failure(agent):
    """Test batch operation where some actions fail."""
    gemini_response = """
    {
        "actions": [
            {"tool": "excel", "operation": "update", "target_file": "test1.xlsx", "parameters": {}},
            {"tool": "excel", "operation": "update", "target_file": "test2.xlsx", "parameters": {}},
            {"tool": "excel", "operation": "update", "target_file": "test3.xlsx", "parameters": {}}
        ],
        "explanation": "Updating cells"
    }
    """
    
    # Mock dependencies - second action fails
    agent.security_validator.validate_operation = Mock()
    agent.security_validator.validate_file_path = Mock(side_effect=lambda x: x)
    
    def mock_execute(tool, op, file, params):
        if file == "test2.xlsx":
            raise Exception("File not found")
    
    agent._execute_single_action = Mock(side_effect=mock_execute)
    
    response = agent._execute_actions(gemini_response)
    
    assert response.success is True  # Overall success because some succeeded
    assert response.batch_result is not None
    assert response.batch_result.total_actions == 3
    assert response.batch_result.successful_actions == 2
    assert response.batch_result.failed_actions == 1
    assert response.batch_result.overall_success is True


def test_batch_all_failed(agent):
    """Test batch operation where all actions fail."""
    gemini_response = """
    {
        "actions": [
            {"tool": "excel", "operation": "update", "target_file": "test1.xlsx", "parameters": {}},
            {"tool": "excel", "operation": "update", "target_file": "test2.xlsx", "parameters": {}}
        ],
        "explanation": "Updating cells"
    }
    """
    
    # Mock dependencies - all actions fail
    agent.security_validator.validate_operation = Mock()
    agent.security_validator.validate_file_path = Mock(side_effect=lambda x: x)
    agent._execute_single_action = Mock(side_effect=Exception("Error"))
    
    response = agent._execute_actions(gemini_response)
    
    assert response.success is False
    assert response.batch_result is not None
    assert response.batch_result.total_actions == 2
    assert response.batch_result.successful_actions == 0
    assert response.batch_result.failed_actions == 2
    assert response.batch_result.overall_success is False


def test_single_operation_no_batch_result(agent):
    """Test that single operation doesn't create batch_result."""
    gemini_response = """
    {
        "actions": [
            {"tool": "excel", "operation": "read", "target_file": "test.xlsx", "parameters": {}}
        ],
        "explanation": "Reading file"
    }
    """
    
    agent.security_validator.validate_operation = Mock()
    agent.security_validator.validate_file_path = Mock(side_effect=lambda x: x)
    agent._execute_single_action = Mock()
    
    response = agent._execute_actions(gemini_response)
    
    assert response.success is True
    assert response.batch_result is None  # Single operation


def test_action_result_details(agent):
    """Test that ActionResult contains correct details."""
    gemini_response = """
    {
        "actions": [
            {"tool": "excel", "operation": "update", "target_file": "test1.xlsx", "parameters": {}},
            {"tool": "word", "operation": "create", "target_file": "test2.docx", "parameters": {}}
        ],
        "explanation": "Multiple operations"
    }
    """
    
    agent.security_validator.validate_operation = Mock()
    agent.security_validator.validate_file_path = Mock(side_effect=lambda x: x)
    
    def mock_execute(tool, op, file, params):
        if file == "test2.docx":
            raise Exception("Permission denied")
    
    agent._execute_single_action = Mock(side_effect=mock_execute)
    
    response = agent._execute_actions(gemini_response)
    
    assert len(response.batch_result.action_results) == 2
    
    # First action succeeded
    result1 = response.batch_result.action_results[0]
    assert result1.action_index == 0
    assert result1.tool == "excel"
    assert result1.operation == "update"
    assert result1.target_file == "test1.xlsx"
    assert result1.success is True
    assert result1.error is None
    
    # Second action failed
    result2 = response.batch_result.action_results[1]
    assert result2.action_index == 1
    assert result2.tool == "word"
    assert result2.operation == "create"
    assert result2.target_file == "test2.docx"
    assert result2.success is False
    assert result2.error == "Permission denied"


def test_batch_message_formatting(agent):
    """Test that batch messages are formatted correctly."""
    gemini_response = """
    {
        "actions": [
            {"tool": "excel", "operation": "update", "target_file": "test1.xlsx", "parameters": {}},
            {"tool": "excel", "operation": "update", "target_file": "test2.xlsx", "parameters": {}},
            {"tool": "excel", "operation": "update", "target_file": "test3.xlsx", "parameters": {}}
        ],
        "explanation": "Test"
    }
    """
    
    agent.security_validator.validate_operation = Mock()
    agent.security_validator.validate_file_path = Mock(side_effect=lambda x: x)
    
    # Test all successful
    agent._execute_single_action = Mock()
    response = agent._execute_actions(gemini_response)
    assert "✅" in response.message
    assert "3" in response.message
    
    # Test partial failure
    def mock_execute_partial(tool, op, file, params):
        if file == "test2.xlsx":
            raise Exception("Error")
    
    agent._execute_single_action = Mock(side_effect=mock_execute_partial)
    response = agent._execute_actions(gemini_response)
    assert "⚠️" in response.message
    assert "2 de 3" in response.message
    
    # Test all failed
    agent._execute_single_action = Mock(side_effect=Exception("Error"))
    response = agent._execute_actions(gemini_response)
    assert "❌" in response.message
