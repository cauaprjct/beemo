"""Módulo de configuração do Gemini Office Agent."""

import os
from dotenv import load_dotenv
from src.exceptions import ConfigurationError

load_dotenv()


# Modelos gratuitos do Gemini, ordenados por preferência de uso
# Usados como fallback automático quando o modelo principal atinge rate/quota limit
# Estáveis:
#   gemini-2.5-flash-lite  — 15 RPM / 250k TPM / ~1000 RPD (mais rápido)
#   gemini-2.5-flash       — 10 RPM / 250k TPM / ~250-500 RPD (equilibrado)
#   gemini-2.5-pro         — 5 RPM / 250k TPM / ~100 RPD (mais inteligente)
# Preview (gratuitos, limites mais restritos):
#   Modelos estáveis da família 2.5 (GA - Generally Available)
DEFAULT_FALLBACK_MODELS = [
    "gemini-2.5-flash",             # Estável — 10 RPM (equilibrado)
    "gemini-2.5-flash-lite",        # Estável — 15 RPM (mais rápido)
    "gemini-2.5-pro",               # Estável — 5 RPM (mais inteligente)
]


class Config:
    """Gerencia configurações do sistema através de variáveis de ambiente.
    
    Attributes:
        api_key: API key do Gemini (obrigatória)
        root_path: Caminho da pasta raiz para varredura de arquivos
        model_name: Nome do modelo Gemini principal
        fallback_models: Lista de modelos para fallback em caso de rate limit
    
    Raises:
        ConfigurationError: Se a API key não estiver configurada
    """
    
    def __init__(self):
        """Inicializa configuração lendo variáveis de ambiente.
        
        Raises:
            ConfigurationError: Se GEMINI_API_KEY não estiver definida
        """
        self._api_key = os.getenv("GEMINI_API_KEY")
        if not self._api_key:
            raise ConfigurationError(
                "GEMINI_API_KEY não está configurada. "
                "Por favor, defina a variável de ambiente GEMINI_API_KEY com sua API key do Gemini."
            )
        
        self._root_path = os.getenv("ROOT_PATH", os.getcwd())
        self._model_name = os.getenv("MODEL_NAME", "gemini-2.5-pro")
        
        # Fallback models: lista separada por vírgula ou usa default
        fallback_env = os.getenv("FALLBACK_MODELS")
        if fallback_env:
            self._fallback_models = [m.strip() for m in fallback_env.split(",") if m.strip()]
        else:
            self._fallback_models = [m for m in DEFAULT_FALLBACK_MODELS if m != self._model_name]
        self._max_versions = self._parse_int_env("MAX_VERSIONS", 10)
        self._cache_enabled = os.getenv("CACHE_ENABLED", "true").lower() == "true"
        self._cache_ttl_hours = self._parse_int_env("CACHE_TTL_HOURS", 24)
        self._cache_max_entries = self._parse_int_env("CACHE_MAX_ENTRIES", 100)
    
    @staticmethod
    def _parse_int_env(var_name: str, default: int) -> int:
        """Parseia variável de ambiente como inteiro com tratamento de erro.
        
        Args:
            var_name: Nome da variável de ambiente
            default: Valor padrão se não definida
        
        Returns:
            Valor inteiro da variável
        
        Raises:
            ConfigurationError: Se o valor não for um inteiro válido
        """
        raw = os.getenv(var_name)
        if raw is None:
            return default
        try:
            return int(raw)
        except ValueError:
            raise ConfigurationError(
                f"Valor inválido para {var_name}: '{raw}'. Esperado um número inteiro."
            )
    
    @property
    def api_key(self) -> str:
        """Retorna API key do Gemini.
        
        Returns:
            str: API key configurada
        """
        return self._api_key
    
    @property
    def root_path(self) -> str:
        """Retorna caminho da pasta raiz.
        
        Returns:
            str: Caminho da pasta raiz para varredura de arquivos
        """
        return self._root_path
    
    @property
    def model_name(self) -> str:
        """Retorna nome do modelo Gemini principal.
        
        Returns:
            str: Nome do modelo Gemini a ser utilizado
        """
        return self._model_name
    
    @property
    def fallback_models(self) -> list:
        """Retorna lista de modelos para fallback em caso de rate limit.
        
        Returns:
            list: Lista de nomes de modelos alternativos
        """
        return self._fallback_models
    
    @property
    def max_versions(self) -> int:
        """Retorna número máximo de versões por arquivo.
        
        Returns:
            int: Número máximo de versões a manter
        """
        return self._max_versions
    
    @property
    def cache_enabled(self) -> bool:
        """Retorna se o cache está habilitado.
        
        Returns:
            bool: True se cache está habilitado
        """
        return self._cache_enabled
    
    @property
    def cache_ttl_hours(self) -> int:
        """Retorna tempo de vida do cache em horas.
        
        Returns:
            int: TTL do cache em horas
        """
        return self._cache_ttl_hours
    
    @property
    def cache_max_entries(self) -> int:
        """Retorna número máximo de entradas no cache.
        
        Returns:
            int: Número máximo de entradas
        """
        return self._cache_max_entries
