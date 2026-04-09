"""Factory module for initializing the Gemini Office Agent system.

This module provides a factory function to create a fully configured Agent
with all dependencies properly instantiated and injected.
"""

import logging
from config import Config
from src.agent import Agent
from src.version_manager import VersionManager
from src.response_cache import ResponseCache
from src.logging_config import get_logger, setup_logging

logger = get_logger(__name__)


def create_agent() -> Agent:
    """Create and configure a fully initialized Agent instance.
    
    This factory function handles the complete initialization workflow:
    1. Instantiate and validate Config
    2. Create Agent with the validated configuration
    3. All dependencies (FileScanner, GeminiClient, Tools) are automatically
       instantiated by the Agent constructor
    
    Returns:
        Agent: Fully configured Agent instance ready for use
    
    Raises:
        ConfigurationError: If configuration is invalid (e.g., missing API key)
        AuthenticationError: If Gemini API authentication fails
    
    Example:
        >>> agent = create_agent()
        >>> response = agent.process_user_request("Create a new Excel file")
    """
    # 0. Initialize logging system
    setup_logging()
    
    logger.info("=== Initializing Gemini Office Agent ===")
    
    # 1. Instantiate and validate Config
    logger.info("Loading configuration...")
    config = Config()
    logger.info(f"Configuration loaded successfully:")
    logger.info(f"  - Model: {config.model_name}")
    logger.info(f"  - Fallback models: {', '.join(config.fallback_models)}")
    logger.info(f"  - Root path: {config.root_path}")
    logger.info(f"  - Max versions: {config.max_versions}")
    logger.info(f"  - Cache enabled: {config.cache_enabled}")
    logger.info(f"  - Cache TTL: {config.cache_ttl_hours}h")
    logger.info(f"  - Cache max entries: {config.cache_max_entries}")
    logger.info(f"  - API key: {'*' * 8} (configured)")
    
    # 2. Create VersionManager
    logger.info("Initializing VersionManager...")
    version_manager = VersionManager(config.root_path, config.max_versions)
    
    # 3. Create ResponseCache
    logger.info("Initializing ResponseCache...")
    cache_dir = f"{config.root_path}/.cache"
    response_cache = ResponseCache(
        cache_dir,
        max_entries=config.cache_max_entries,
        ttl_hours=config.cache_ttl_hours,
        enabled=config.cache_enabled
    )
    
    # 4. Create Agent with validated configuration
    # The Agent constructor handles instantiation of:
    #   - FileScanner (with root_path from Config)
    #   - GeminiClient (with api_key and model_name from Config + ResponseCache)
    #   - ExcelTool
    #   - WordTool
    #   - PowerPointTool
    #   - SecurityValidator (with root_path from Config)
    logger.info("Creating Agent with all dependencies...")
    agent = Agent(config, version_manager, response_cache)
    
    logger.info("=== Agent initialized successfully ===")
    
    return agent
