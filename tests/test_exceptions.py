"""Unit tests for custom exceptions."""

import pytest
from src.exceptions import (
    ConfigurationError,
    AuthenticationError,
    QuotaExceededError,
    CorruptedFileError,
    ValidationError,
    NetworkError,
)


def test_configuration_error():
    """Test ConfigurationError can be raised and caught."""
    with pytest.raises(ConfigurationError) as exc_info:
        raise ConfigurationError("Missing API key")
    assert "Missing API key" in str(exc_info.value)


def test_authentication_error():
    """Test AuthenticationError can be raised and caught."""
    with pytest.raises(AuthenticationError) as exc_info:
        raise AuthenticationError("Invalid API key")
    assert "Invalid API key" in str(exc_info.value)


def test_quota_exceeded_error():
    """Test QuotaExceededError can be raised and caught."""
    with pytest.raises(QuotaExceededError) as exc_info:
        raise QuotaExceededError("Daily quota exceeded")
    assert "Daily quota exceeded" in str(exc_info.value)


def test_corrupted_file_error():
    """Test CorruptedFileError can be raised and caught."""
    with pytest.raises(CorruptedFileError) as exc_info:
        raise CorruptedFileError("File is corrupted: test.xlsx")
    assert "File is corrupted" in str(exc_info.value)


def test_validation_error():
    """Test ValidationError can be raised and caught."""
    with pytest.raises(ValidationError) as exc_info:
        raise ValidationError("Invalid data structure")
    assert "Invalid data structure" in str(exc_info.value)


def test_network_error():
    """Test NetworkError can be raised and caught."""
    with pytest.raises(NetworkError) as exc_info:
        raise NetworkError("Connection timeout")
    assert "Connection timeout" in str(exc_info.value)


def test_exception_inheritance():
    """Test that all custom exceptions inherit from Exception."""
    assert issubclass(ConfigurationError, Exception)
    assert issubclass(AuthenticationError, Exception)
    assert issubclass(QuotaExceededError, Exception)
    assert issubclass(CorruptedFileError, Exception)
    assert issubclass(ValidationError, Exception)
    assert issubclass(NetworkError, Exception)


def test_exception_with_no_message():
    """Test exceptions can be raised without a message."""
    with pytest.raises(ConfigurationError):
        raise ConfigurationError()
    
    with pytest.raises(AuthenticationError):
        raise AuthenticationError()


def test_exception_with_detailed_message():
    """Test exceptions with detailed error messages."""
    detailed_message = (
        "Configuration error: GEMINI_API_KEY environment variable not set. "
        "Please set it in your .env file or environment."
    )
    with pytest.raises(ConfigurationError) as exc_info:
        raise ConfigurationError(detailed_message)
    assert "GEMINI_API_KEY" in str(exc_info.value)
    assert ".env" in str(exc_info.value)
