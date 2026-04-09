"""Custom exceptions for Gemini Office Agent.

This module defines custom exception classes for different error scenarios
that can occur during system operation, providing clear error categorization
and handling.
"""


class ConfigurationError(Exception):
    """Raised when there are configuration issues.
    
    Examples:
        - Missing API key
        - Invalid paths
        - Invalid model name
    """
    pass


class AuthenticationError(Exception):
    """Raised when API authentication fails.
    
    Examples:
        - Invalid API key
        - Expired credentials
        - Unauthorized access
    """
    pass


class QuotaExceededError(Exception):
    """Raised when API quota or rate limits are exceeded.
    
    Examples:
        - Daily quota exceeded
        - Rate limit hit
        - Too many requests
    """
    pass


class CorruptedFileError(Exception):
    """Raised when an Office file is corrupted or cannot be read.
    
    Examples:
        - Malformed Excel file
        - Corrupted Word document
        - Invalid PowerPoint structure
    """
    pass


class ValidationError(Exception):
    """Raised when data validation fails.
    
    Examples:
        - Invalid Excel data structure
        - Invalid Word content format
        - Invalid PowerPoint slide data
    """
    pass


class NetworkError(Exception):
    """Raised when network-related errors occur.
    
    Examples:
        - Connection timeout
        - DNS resolution failure
        - Network unreachable
    """
    pass
