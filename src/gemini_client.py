"""Gemini API client for natural language processing.

This module provides a wrapper around the Google Generative AI API,
handling authentication, request/response processing, error handling,
and automatic fallback between models when rate limits are hit.
"""

import logging
from typing import List, Optional, Any
from google import genai
from google.genai import types
from google.genai import errors as google_exceptions
from src.exceptions import (
    AuthenticationError,
    QuotaExceededError,
    NetworkError,
)
from src.logging_config import get_logger

logger = get_logger(__name__)


class GeminiClient:
    """Client for interacting with Google Gemini API with model fallback.
    
    This class encapsulates all communication with the Gemini API,
    providing a simple interface for generating responses from prompts.
    When the primary model hits rate/quota limits, automatically falls
    back to alternative models.
    
    Attributes:
        model_name: The name of the primary Gemini model
        fallback_models: List of alternative model names for fallback
        client: The configured GenAI client instance
    """
    
    def __init__(self, api_key: str, model_name: str, response_cache=None,
                 fallback_models: Optional[List[str]] = None):
        """Initialize the Gemini client.
        
        Args:
            api_key: Google API key for authentication
            model_name: Name of the primary Gemini model (e.g., 'gemini-2.5-flash-lite')
            response_cache: Optional ResponseCache instance
            fallback_models: Optional list of model names to try when primary fails
                             due to rate limit / quota. Tried in order.
        
        Raises:
            AuthenticationError: If API key is invalid or authentication fails
        """
        self._api_key = api_key
        self.model_name = model_name
        self.fallback_models = fallback_models or []
        self.response_cache = response_cache
        self._active_model_name = model_name
        
        try:
            # Configure the API with the provided key (timeout on client level)
            self.client = genai.Client(
                api_key=api_key,
                http_options={'timeout': 60000} # 60s timeout defaults
            )
            
            logger.info(f"GeminiClient initialized with model: {model_name}")
            if self.fallback_models:
                logger.info(f"Fallback models: {', '.join(self.fallback_models)}")
            
        except Exception as e:
            logger.error(f"Failed to initialize GeminiClient: {e}", exc_info=True)
            raise AuthenticationError(f"Failed to authenticate with Gemini: {e}")
    
    @property
    def active_model_name(self) -> str:
        """Retorna o nome do modelo que foi usado na última chamada com sucesso."""
        return self._active_model_name
    
    def generate_response(self, prompt: str, timeout: int = 30) -> str:
        """Generate a response from the Gemini model with automatic fallback.
        
        Sends a prompt to the Gemini API and returns the generated text response.
        Checks cache first to avoid duplicate API calls. If the primary model
        hits a rate limit or quota error, automatically tries fallback models
        in order.
        
        Args:
            prompt: The input prompt to send to the model
            timeout: Request timeout in seconds (default: 30) (Currently ignored via argument as it is set in HTTP client)
        
        Returns:
            The generated text response from the model
            
        Raises:
            AuthenticationError: If authentication fails during the request
            QuotaExceededError: If ALL models (primary + fallbacks) hit quota
            TimeoutError: If the request times out
            NetworkError: If a network error occurs
        """
        # Check cache first
        if self.response_cache:
            cached_response = self.response_cache.get(prompt)
            if cached_response:
                logger.info("⚡ Response retrieved from cache")
                return cached_response
        
        # Build model list with sticky preference
        all_models = [self.model_name] + self.fallback_models
        if self._active_model_name != self.model_name and self._active_model_name in all_models:
            idx = all_models.index(self._active_model_name)
            models_to_try = [self._active_model_name] + all_models[:idx] + all_models[idx+1:]
        else:
            models_to_try = all_models
        
        last_error = None
        
        for current_model in models_to_try:
            try:
                response_text = self._call_model(current_model, prompt)
                self._active_model_name = current_model
                
                # Save to cache
                if self.response_cache:
                    self.response_cache.set(prompt, response_text)
                
                return response_text
                
            except google_exceptions.APIError as e:
                if e.code in (429, 503):
                    last_error = e
                    logger.warning(
                        f"{'Rate limit' if e.code == 429 else 'Unavailable'} "
                        f"on {current_model}, trying next..."
                    )
                    continue
                else:
                    raise
        
        logger.error("Todos os modelos indisponíveis")
        raise QuotaExceededError(
            f"Todos os modelos indisponíveis. "
            f"Tentados: {', '.join(models_to_try)}. "
            f"Último erro: {last_error}"
        )
    
    def generate_with_tools(
        self,
        contents: List[Any],
        tools: List[Any],
        system_instruction: Optional[str] = None,
    ) -> Any:
        """Generate a response with native function calling support.

        Unlike generate_response(), this method does NOT force JSON output.
        The model natively decides whether to call a tool or reply with text.

        Args:
            contents: List of types.Content objects (conversation history)
            tools: List of types.Tool objects with function declarations
            system_instruction: Optional system prompt

        Returns:
            types.GenerateContentResponse — caller inspects .candidates[0].content
        """
        # Build model list starting from the last successful model ("sticky")
        # to avoid wasting calls on rate-limited models
        all_models = [self.model_name] + self.fallback_models
        if self._active_model_name != self.model_name and self._active_model_name in all_models:
            # Put the last-successful model first, then the rest
            idx = all_models.index(self._active_model_name)
            models_to_try = [self._active_model_name] + all_models[:idx] + all_models[idx+1:]
        else:
            models_to_try = all_models
        last_error = None

        config_kwargs: dict = {
            "tools": tools,
            # Disable Automatic Function Calling so the SDK does NOT consume
            # function_call parts internally — we execute them in the agentic loop.
            "automatic_function_calling": {"disable": True},
        }
        if system_instruction:
            config_kwargs["system_instruction"] = system_instruction

        config = types.GenerateContentConfig(**config_kwargs)

        for current_model in models_to_try:
            try:
                logger.info(f"[FC] Sending tool-enabled request to {current_model}")
                response = self.client.models.generate_content(
                    model=current_model,
                    contents=contents,
                    config=config,
                )
                self._active_model_name = current_model
                return response

            except google_exceptions.APIError as e:
                if e.code in (429, 503):
                    # 429 = rate limit, 503 = high demand — both are transient
                    last_error = e
                    logger.warning(
                        f"{'Rate limit' if e.code == 429 else 'Service unavailable'} "
                        f"on {current_model}, trying next model..."
                    )
                    continue
                elif e.code == 400 and "thought_signature" in str(e):
                    # Gemini 3.x models require thought_signature support;
                    # skip to next model if not fully compatible
                    last_error = e
                    logger.warning(
                        f"Model {current_model} requires thought_signature "
                        f"(not fully supported), trying next model..."
                    )
                    continue
                elif e.code in (401, 403):
                    raise AuthenticationError(f"Authentication failed: {e}")
                else:
                    raise NetworkError(f"API error {e.code}: {e}")
            except Exception as e:
                raise NetworkError(f"Unexpected error: {e}")

        raise QuotaExceededError(
            f"All models unavailable. Tried: {', '.join(models_to_try)}. "
            f"Last error: {last_error}"
        )

    def generate_text_response(self, prompt: str) -> str:
        """Generate a plain-text response (no JSON forcing).

        Unlike generate_response() which forces application/json, this method
        returns natural language text — ideal for chat / conversational replies.

        Uses the same fallback chain as generate_response().

        Args:
            prompt: The input prompt to send to the model

        Returns:
            Plain-text response string

        Raises:
            AuthenticationError: If authentication fails
            QuotaExceededError: If all models hit quota
            NetworkError: If a network error occurs
        """
        # Build model list with sticky preference
        all_models = [self.model_name] + self.fallback_models
        if self._active_model_name != self.model_name and self._active_model_name in all_models:
            idx = all_models.index(self._active_model_name)
            models_to_try = [self._active_model_name] + all_models[:idx] + all_models[idx+1:]
        else:
            models_to_try = all_models
        last_error = None

        for current_model in models_to_try:
            try:
                logger.info(f"Sending text request to {current_model}")
                response = self.client.models.generate_content(
                    model=current_model,
                    contents=prompt,
                    # No response_mime_type → plain text output
                )
                self._active_model_name = current_model
                text = response.text or ""
                logger.info(f"Text response from {current_model} ({len(text)} chars)")
                return text

            except google_exceptions.APIError as e:
                if e.code in (429, 503):
                    last_error = e
                    logger.warning(f"{'Rate limit' if e.code == 429 else 'Unavailable'} on {current_model}, trying next...")
                    continue
                elif e.code in (401, 403):
                    raise AuthenticationError(f"Authentication failed: {e}")
                else:
                    raise NetworkError(f"API error {e.code}: {e}")
            except Exception as e:
                if "timeout" in str(e).lower():
                    raise TimeoutError(f"Request timed out: {e}")
                raise NetworkError(f"Unexpected error: {e}")

        raise QuotaExceededError(
            f"All models unavailable. Tried: {', '.join(models_to_try)}. "
            f"Last error: {last_error}"
        )

    def _call_model(self, model_name: str, prompt: str) -> str:
        """Executa chamada a um modelo específico, forçando saída no formato JSON estrito.
        
        Args:
            model_name: Nome do modelo (ex: gemini-2.5-flash)
            prompt: Prompt a enviar ao modelo
        
        Returns:
            Texto da resposta JSON
        
        Raises:
            google_exceptions.APIError: Erros gerados pelo SDK oficial (autenticação, erro 400, 429)
            NetworkError: Se erro de rede ocorrer
        """
        logger.info(f"Sending request to Gemini API (model: {model_name})")
        logger.debug(f"Prompt length: {len(prompt)} characters")
        
        try:
            # We enforce exactly valid JSON response by specifying response_mime_type.
            config = types.GenerateContentConfig(
                response_mime_type="application/json",
            )
            
            response = self.client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=config
            )
            
            response_text = response.text or ""
            
            logger.info(f"Received response from {model_name} (length: {len(response_text)} characters)")
            logger.debug(f"Response preview: {response_text[:100]}...")
            
            return response_text
            
        except google_exceptions.APIError as e:
            if e.code == 401 or e.code == 403:
                logger.error(f"Authentication error: {e}", exc_info=True)
                raise AuthenticationError(
                    f"Authentication failed. Please check your API key: {e}"
                )
            elif e.code == 429:
                # Deixar propagar pro fallback
                raise
            else:
                logger.error(f"Network error (HTTP {e.code}): {e}", exc_info=True)
                raise NetworkError(
                    f"Network error occurred while communicating with Gemini API: {e}"
                )
        except Exception as e:
            logger.error(f"Unexpected error during API call: {e}", exc_info=True)
            if "timeout" in str(e).lower():
                raise TimeoutError(f"Request timed out: {e}")
            raise NetworkError(f"Unexpected error occurred: {e}")
