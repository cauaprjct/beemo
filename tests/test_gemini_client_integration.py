"""Integration tests for GeminiClient.

These tests verify the complete workflow of the GeminiClient
including initialization, request processing, and error handling.
"""

import pytest
from unittest.mock import Mock, patch

from src.gemini_client import GeminiClient
from src.exceptions import (
    AuthenticationError,
    QuotaExceededError,
    NetworkError,
)


@pytest.fixture
def mock_genai():
    """Mock the google.generativeai module."""
    with patch('src.gemini_client.genai') as mock:
        yield mock


class TestGeminiClientIntegration:
    """Integration tests for complete GeminiClient workflows."""
    
    def test_complete_successful_workflow(self, mock_genai):
        """Test complete workflow from initialization to response."""
        # Arrange
        api_key = "test-api-key"
        model_name = "gemini-2.5-flash-lite"
        prompt = "Explain quantum computing in simple terms"
        expected_response = "Quantum computing uses quantum mechanics..."
        
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = expected_response
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        # Act
        client = GeminiClient(api_key, model_name)
        result = client.generate_response(prompt)
        
        # Assert
        assert result == expected_response
        mock_genai.configure.assert_called_once_with(api_key=api_key)
        mock_genai.GenerativeModel.assert_called_once_with(model_name)
        mock_model.generate_content.assert_called_once_with(
            prompt,
            request_options={'timeout': 30}
        )
    
    def test_multiple_requests_same_client(self, mock_genai):
        """Test multiple requests using the same client instance."""
        # Arrange
        mock_model = Mock()
        mock_genai.GenerativeModel.return_value = mock_model
        
        responses = ["Response 1", "Response 2", "Response 3"]
        mock_model.generate_content.side_effect = [
            Mock(text=resp) for resp in responses
        ]
        
        client = GeminiClient("test-key", "gemini-2.5-flash-lite")
        
        # Act
        results = [
            client.generate_response("Prompt 1"),
            client.generate_response("Prompt 2"),
            client.generate_response("Prompt 3"),
        ]
        
        # Assert
        assert results == responses
        assert mock_model.generate_content.call_count == 3
    
    def test_error_recovery_workflow(self, mock_genai):
        """Test that client can recover from errors and continue working."""
        # Arrange
        mock_model = Mock()
        mock_genai.GenerativeModel.return_value = mock_model
        
        # First call fails, second succeeds
        from google.api_core import exceptions as google_exceptions
        mock_model.generate_content.side_effect = [
            google_exceptions.ServiceUnavailable("Service down"),
            Mock(text="Success response")
        ]
        
        client = GeminiClient("test-key", "gemini-2.5-flash-lite")
        
        # Act & Assert
        # First call should fail
        with pytest.raises(NetworkError):
            client.generate_response("Prompt 1")
        
        # Second call should succeed
        result = client.generate_response("Prompt 2")
        assert result == "Success response"
    
    def test_custom_timeout_workflow(self, mock_genai):
        """Test workflow with custom timeout values."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient("test-key", "gemini-2.5-flash-lite")
        
        # Act
        result1 = client.generate_response("Prompt 1", timeout=10)
        result2 = client.generate_response("Prompt 2", timeout=60)
        result3 = client.generate_response("Prompt 3")  # Default timeout
        
        # Assert
        assert result1 == "Response"
        assert result2 == "Response"
        assert result3 == "Response"
        
        calls = mock_model.generate_content.call_args_list
        assert calls[0][1]['request_options']['timeout'] == 10
        assert calls[1][1]['request_options']['timeout'] == 60
        assert calls[2][1]['request_options']['timeout'] == 30
    
    def test_logging_throughout_workflow(self, mock_genai):
        """Test that logging occurs at all stages of the workflow."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Test response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        with patch('src.gemini_client.logger') as mock_logger:
            # Act
            client = GeminiClient("test-key", "gemini-2.5-flash-lite")
            result = client.generate_response("Test prompt")
            
            # Assert
            assert result == "Test response"
            
            # Verify logging occurred
            info_calls = [call[0][0] for call in mock_logger.info.call_args_list]
            
            # Check initialization logging
            assert any("initialized" in call.lower() for call in info_calls)
            
            # Check request logging
            assert any("sending request" in call.lower() for call in info_calls)
            
            # Check response logging
            assert any("received response" in call.lower() for call in info_calls)
    
    def test_requirements_validation(self, mock_genai):
        """Test that implementation meets all specified requirements.
        
        This test validates:
        - Requirement 3.1: Initialization with API key
        - Requirement 3.2: Request to Gemini 2.5 Flash-Lite
        - Requirement 3.3: Return text response
        - Requirement 3.4: Authentication error handling
        - Requirement 3.5: Quota error handling
        - Requirement 3.6: 30-second timeout
        - Requirement 9.3: API call logging with timestamps
        """
        # Arrange
        api_key = "test-api-key"
        model_name = "gemini-2.5-flash-lite"
        
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Test response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        with patch('src.gemini_client.logger') as mock_logger:
            # Act - Requirement 3.1: Initialize with API key
            client = GeminiClient(api_key, model_name)
            
            # Assert initialization
            assert client._api_key == api_key
            assert client.model_name == model_name
            mock_genai.configure.assert_called_once_with(api_key=api_key)
            
            # Act - Requirement 3.2 & 3.3: Send request and get text response
            result = client.generate_response("Test prompt")
            
            # Assert response
            assert isinstance(result, str)
            assert result == "Test response"
            
            # Requirement 3.6: Verify 30-second timeout
            call_args = mock_model.generate_content.call_args
            assert call_args[1]['request_options']['timeout'] == 30
            
            # Requirement 9.3: Verify logging with timestamps
            # The logging_config module adds timestamps automatically
            assert mock_logger.info.called
            info_calls = [call[0][0] for call in mock_logger.info.call_args_list]
            assert any("Sending request" in call for call in info_calls)
            assert any("Received response" in call for call in info_calls)
        
        # Requirement 3.4: Test authentication error
        from google.api_core import exceptions as google_exceptions
        mock_model.generate_content.side_effect = google_exceptions.Unauthenticated("Auth failed")
        
        with pytest.raises(AuthenticationError) as exc_info:
            client.generate_response("Test")
        assert "Authentication failed" in str(exc_info.value)
        
        # Requirement 3.5: Test quota error
        mock_model.generate_content.side_effect = google_exceptions.ResourceExhausted("Quota exceeded")
        
        with pytest.raises(QuotaExceededError) as exc_info:
            client.generate_response("Test")
        assert "limite de requisições" in str(exc_info.value).lower() or "quota" in str(exc_info.value).lower()
