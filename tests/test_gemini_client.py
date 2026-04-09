"""Unit tests for GeminiClient."""

import pytest
from unittest.mock import Mock, patch, MagicMock
from google.api_core import exceptions as google_exceptions

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


@pytest.fixture
def valid_api_key():
    """Provide a valid API key for testing."""
    return "test-api-key-12345"


@pytest.fixture
def model_name():
    """Provide a model name for testing."""
    return "gemini-2.5-flash-lite"


class TestGeminiClientInitialization:
    """Tests for GeminiClient initialization."""
    
    def test_init_success(self, mock_genai, valid_api_key, model_name):
        """Test successful initialization with valid API key."""
        # Arrange
        mock_model = Mock()
        mock_genai.GenerativeModel.return_value = mock_model
        
        # Act
        client = GeminiClient(valid_api_key, model_name)
        
        # Assert
        assert client._api_key == valid_api_key
        assert client.model_name == model_name
        assert client.model == mock_model
        mock_genai.configure.assert_called_once_with(api_key=valid_api_key)
        mock_genai.GenerativeModel.assert_called_once_with(model_name)
    
    def test_init_with_authentication_error(self, mock_genai, valid_api_key, model_name):
        """Test initialization fails with invalid API key."""
        # Arrange
        mock_genai.configure.side_effect = Exception("Invalid API key")
        
        # Act & Assert
        with pytest.raises(AuthenticationError) as exc_info:
            GeminiClient(valid_api_key, model_name)
        
        assert "Failed to authenticate" in str(exc_info.value)
        assert "Invalid API key" in str(exc_info.value)
    
    def test_init_with_invalid_model_name(self, mock_genai, valid_api_key):
        """Test initialization with invalid model name."""
        # Arrange
        mock_genai.GenerativeModel.side_effect = Exception("Model not found")
        
        # Act & Assert
        with pytest.raises(AuthenticationError) as exc_info:
            GeminiClient(valid_api_key, "invalid-model")
        
        assert "Failed to authenticate" in str(exc_info.value)


class TestGeminiClientGenerateResponse:
    """Tests for generate_response method."""
    
    def test_generate_response_success(self, mock_genai, valid_api_key, model_name):
        """Test successful response generation."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "This is a generated response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act
        result = client.generate_response("Test prompt")
        
        # Assert
        assert result == "This is a generated response"
        mock_model.generate_content.assert_called_once_with(
            "Test prompt",
            request_options={'timeout': 30}
        )
    
    def test_generate_response_with_custom_timeout(self, mock_genai, valid_api_key, model_name):
        """Test response generation with custom timeout."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act
        result = client.generate_response("Test prompt", timeout=60)
        
        # Assert
        assert result == "Response"
        mock_model.generate_content.assert_called_once_with(
            "Test prompt",
            request_options={'timeout': 60}
        )
    
    def test_generate_response_authentication_error(self, mock_genai, valid_api_key, model_name):
        """Test authentication error during response generation."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = google_exceptions.Unauthenticated("Invalid credentials")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(AuthenticationError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "Authentication failed" in str(exc_info.value)
        assert "API key" in str(exc_info.value)
    
    def test_generate_response_quota_exceeded(self, mock_genai, valid_api_key, model_name):
        """Test quota exceeded error."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = google_exceptions.ResourceExhausted("Quota exceeded")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(QuotaExceededError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "limite de requisições" in str(exc_info.value).lower() or "quota" in str(exc_info.value).lower()
    
    def test_generate_response_timeout(self, mock_genai, valid_api_key, model_name):
        """Test timeout error."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = google_exceptions.DeadlineExceeded("Timeout")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(TimeoutError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "timed out" in str(exc_info.value).lower()
        assert "30 seconds" in str(exc_info.value)
    
    def test_generate_response_service_unavailable(self, mock_genai, valid_api_key, model_name):
        """Test service unavailable error."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = google_exceptions.ServiceUnavailable("Service down")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(NetworkError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "Network error" in str(exc_info.value)
    
    def test_generate_response_internal_server_error(self, mock_genai, valid_api_key, model_name):
        """Test internal server error."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = google_exceptions.InternalServerError("Server error")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(NetworkError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "Network error" in str(exc_info.value)
    
    def test_generate_response_connection_error(self, mock_genai, valid_api_key, model_name):
        """Test connection error."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = ConnectionError("Connection failed")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(NetworkError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "Network error" in str(exc_info.value)
    
    def test_generate_response_unexpected_error(self, mock_genai, valid_api_key, model_name):
        """Test unexpected error handling."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = RuntimeError("Unexpected error")
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act & Assert
        with pytest.raises(NetworkError) as exc_info:
            client.generate_response("Test prompt")
        
        assert "Unexpected error" in str(exc_info.value)
    
    def test_generate_response_empty_prompt(self, mock_genai, valid_api_key, model_name):
        """Test response generation with empty prompt."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Empty prompt response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        
        # Act
        result = client.generate_response("")
        
        # Assert
        assert result == "Empty prompt response"
        mock_model.generate_content.assert_called_once()
    
    def test_generate_response_long_prompt(self, mock_genai, valid_api_key, model_name):
        """Test response generation with very long prompt."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Long prompt response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        long_prompt = "A" * 10000
        
        # Act
        result = client.generate_response(long_prompt)
        
        # Assert
        assert result == "Long prompt response"
        mock_model.generate_content.assert_called_once_with(
            long_prompt,
            request_options={'timeout': 30}
        )
    
    def test_generate_response_special_characters(self, mock_genai, valid_api_key, model_name):
        """Test response generation with special characters in prompt."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Special chars response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        client = GeminiClient(valid_api_key, model_name)
        special_prompt = "Test with émojis 🎉 and symbols @#$%"
        
        # Act
        result = client.generate_response(special_prompt)
        
        # Assert
        assert result == "Special chars response"


class TestGeminiClientLogging:
    """Tests for logging functionality."""
    
    def test_initialization_logging(self, mock_genai, valid_api_key, model_name):
        """Test that initialization is logged."""
        # Arrange
        mock_model = Mock()
        mock_genai.GenerativeModel.return_value = mock_model
        
        with patch('src.gemini_client.logger') as mock_logger:
            # Act
            client = GeminiClient(valid_api_key, model_name)
            
            # Assert
            mock_logger.info.assert_called_with(
                f"GeminiClient initialized with model: {model_name}"
            )
    
    def test_generate_response_logging(self, mock_genai, valid_api_key, model_name):
        """Test that API calls are logged with timestamps."""
        # Arrange
        mock_model = Mock()
        mock_response = Mock()
        mock_response.text = "Test response"
        mock_model.generate_content.return_value = mock_response
        mock_genai.GenerativeModel.return_value = mock_model
        
        with patch('src.gemini_client.logger') as mock_logger:
            client = GeminiClient(valid_api_key, model_name)
            
            # Act
            client.generate_response("Test prompt")
            
            # Assert
            # Check that info logs were called for request and response
            info_calls = [call[0][0] for call in mock_logger.info.call_args_list]
            assert any("Sending request to Gemini API" in call for call in info_calls)
            assert any("Received response from" in call for call in info_calls)
    
    def test_error_logging(self, mock_genai, valid_api_key, model_name):
        """Test that errors are logged with stack traces."""
        # Arrange
        mock_model = Mock()
        mock_model.generate_content.side_effect = google_exceptions.Unauthenticated("Auth error")
        mock_genai.GenerativeModel.return_value = mock_model
        
        with patch('src.gemini_client.logger') as mock_logger:
            client = GeminiClient(valid_api_key, model_name)
            
            # Act
            try:
                client.generate_response("Test prompt")
            except AuthenticationError:
                pass
            
            # Assert
            mock_logger.error.assert_called()
            # Verify exc_info=True was passed for stack trace
            call_kwargs = mock_logger.error.call_args[1]
            assert call_kwargs.get('exc_info') is True
