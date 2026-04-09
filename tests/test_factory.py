"""Tests for the factory module."""

import os
import pytest
from unittest.mock import patch, MagicMock

from src.factory import create_agent
from src.agent import Agent
from src.exceptions import ConfigurationError


class TestFactory:
    """Test suite for factory module."""
    
    def test_create_agent_success(self):
        """Test successful agent creation with valid configuration."""
        # Set up environment variables
        with patch.dict(os.environ, {
            'GEMINI_API_KEY': 'test-api-key-12345',
            'ROOT_PATH': '/tmp/test',
            'MODEL_NAME': 'gemini-2.5-flash-lite'
        }):
            # Mock the GeminiClient to avoid actual API calls
            with patch('src.agent.GeminiClient') as mock_gemini:
                mock_gemini.return_value = MagicMock()
                
                # Create agent
                agent = create_agent()
                
                # Verify agent is created
                assert agent is not None
                assert isinstance(agent, Agent)
                
                # Verify config is set correctly
                assert agent.config.api_key == 'test-api-key-12345'
                assert agent.config.root_path == '/tmp/test'
                assert agent.config.model_name == 'gemini-2.5-flash-lite'
    
    def test_create_agent_missing_api_key(self):
        """Test that factory raises ConfigurationError when API key is missing."""
        # Remove API key from environment
        with patch.dict(os.environ, {}, clear=True):
            with pytest.raises(ConfigurationError) as exc_info:
                create_agent()
            
            assert "GEMINI_API_KEY" in str(exc_info.value)
    
    def test_create_agent_default_values(self):
        """Test that factory uses default values when optional config is not provided."""
        # Set only required API key
        with patch.dict(os.environ, {
            'GEMINI_API_KEY': 'test-api-key-12345'
        }, clear=True):
            # Mock the GeminiClient to avoid actual API calls
            with patch('src.agent.GeminiClient') as mock_gemini:
                mock_gemini.return_value = MagicMock()
                
                # Create agent
                agent = create_agent()
                
                # Verify defaults are used
                assert agent.config.model_name == 'gemini-2.5-flash-lite'
                assert agent.config.root_path == os.getcwd()
    
    def test_create_agent_dependencies_injected(self):
        """Test that all dependencies are properly injected into Agent."""
        with patch.dict(os.environ, {
            'GEMINI_API_KEY': 'test-api-key-12345',
            'ROOT_PATH': '/tmp/test'
        }):
            # Mock the GeminiClient to avoid actual API calls
            with patch('src.agent.GeminiClient') as mock_gemini:
                mock_gemini.return_value = MagicMock()
                
                # Create agent
                agent = create_agent()
                
                # Verify all dependencies are present
                assert agent.file_scanner is not None
                assert agent.gemini_client is not None
                assert agent.excel_tool is not None
                assert agent.word_tool is not None
                assert agent.powerpoint_tool is not None
                assert agent.security_validator is not None
