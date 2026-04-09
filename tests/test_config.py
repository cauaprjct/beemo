"""Unit tests for Config class."""

import os
import pytest
from config import Config
from src.exceptions import ConfigurationError


class TestConfig:
    """Test suite for Config class."""
    
    def test_config_with_valid_api_key(self, monkeypatch):
        """Test Config initialization with valid API key."""
        monkeypatch.setenv("GEMINI_API_KEY", "test_api_key_123")
        
        config = Config()
        
        assert config.api_key == "test_api_key_123"
    
    def test_config_missing_api_key_raises_error(self, monkeypatch):
        """Test Config raises ConfigurationError when API key is missing."""
        monkeypatch.delenv("GEMINI_API_KEY", raising=False)
        
        with pytest.raises(ConfigurationError) as exc_info:
            Config()
        
        assert "GEMINI_API_KEY não está configurada" in str(exc_info.value)
    
    def test_config_default_root_path(self, monkeypatch):
        """Test Config uses current directory as default root_path."""
        monkeypatch.setenv("GEMINI_API_KEY", "test_api_key_123")
        monkeypatch.delenv("ROOT_PATH", raising=False)
        
        config = Config()
        
        assert config.root_path == os.getcwd()
    
    def test_config_custom_root_path(self, monkeypatch):
        """Test Config uses custom root_path from environment."""
        monkeypatch.setenv("GEMINI_API_KEY", "test_api_key_123")
        monkeypatch.setenv("ROOT_PATH", "/custom/path")
        
        config = Config()
        
        assert config.root_path == "/custom/path"
    
    def test_config_default_model_name(self, monkeypatch):
        """Test Config uses default model name."""
        monkeypatch.setenv("GEMINI_API_KEY", "test_api_key_123")
        monkeypatch.delenv("MODEL_NAME", raising=False)
        
        config = Config()
        
        assert config.model_name == "gemini-2.5-flash-lite"
    
    def test_config_custom_model_name(self, monkeypatch):
        """Test Config uses custom model name from environment."""
        monkeypatch.setenv("GEMINI_API_KEY", "test_api_key_123")
        monkeypatch.setenv("MODEL_NAME", "gemini-pro")
        
        config = Config()
        
        assert config.model_name == "gemini-pro"
    
    def test_config_properties_are_readonly(self, monkeypatch):
        """Test Config properties return correct values."""
        monkeypatch.setenv("GEMINI_API_KEY", "test_key")
        monkeypatch.setenv("ROOT_PATH", "/test/path")
        monkeypatch.setenv("MODEL_NAME", "test-model")
        
        config = Config()
        
        assert config.api_key == "test_key"
        assert config.root_path == "/test/path"
        assert config.model_name == "test-model"
