"""Unit tests for Streamlit application."""

import pytest
from unittest.mock import Mock, patch
from pathlib import Path

# Import the app module
import app
from src.models import AgentResponse


class TestStreamlitAppBasic:
    """Basic test suite for Streamlit application - tests that don't require full Streamlit mocking."""
    
    def test_app_module_imports(self):
        """Test that the app module can be imported successfully."""
        assert hasattr(app, 'main')
        assert hasattr(app, 'initialize_session_state')
        assert hasattr(app, 'initialize_agent')
        assert hasattr(app, 'discover_files')
        assert hasattr(app, 'display_file_selector')
        assert hasattr(app, 'display_conversation_history')
        assert hasattr(app, 'process_user_request')
    
    def test_discover_files_failure(self):
        """Test file discovery failure handling."""
        mock_agent = Mock()
        mock_agent.file_scanner.scan_office_files.side_effect = Exception("Scan error")
        
        # Create a mock session state object with attribute access
        class MockSessionState:
            def __init__(self):
                self.agent = mock_agent
        
        mock_session_state = MockSessionState()
        
        with patch('app.st.session_state', mock_session_state), \
             patch('app.st.error') as mock_error, \
             patch('app.logger'):
            
            files = app.discover_files()
            
            assert files == []
            mock_error.assert_called_once()
    
    def test_display_file_selector_no_files(self):
        """Test file selector display with no files."""
        with patch('app.st.info') as mock_info:
            app.display_file_selector([])
            mock_info.assert_called_once()
    
    def test_agent_response_structure(self):
        """Test that AgentResponse has the expected structure."""
        response = AgentResponse(
            success=True,
            message="Test message",
            files_modified=['/path/to/file.xlsx'],
            error=None
        )
        
        assert response.success is True
        assert response.message == "Test message"
        assert len(response.files_modified) == 1
        assert response.error is None
    
    def test_agent_response_with_error(self):
        """Test AgentResponse with error."""
        response = AgentResponse(
            success=False,
            message="Operation failed",
            files_modified=[],
            error="File not found"
        )
        
        assert response.success is False
        assert response.error == "File not found"
        assert len(response.files_modified) == 0


class TestStreamlitAppIntegration:
    """Integration tests that verify the app works with real components."""
    
    def test_app_has_correct_structure(self):
        """Test that the app has the correct function structure."""
        # Verify all required functions exist
        assert callable(app.main)
        assert callable(app.initialize_session_state)
        assert callable(app.initialize_agent)
        assert callable(app.discover_files)
        assert callable(app.display_file_selector)
        assert callable(app.display_conversation_history)
        assert callable(app.process_user_request)
    
    def test_path_handling(self):
        """Test that Path is used correctly for file handling."""
        test_path = "/path/to/file.xlsx"
        p = Path(test_path)
        
        assert p.name == "file.xlsx"
        assert p.suffix == ".xlsx"
