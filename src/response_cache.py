"""Módulo de cache de respostas do Gemini API.

Este módulo fornece funcionalidade de cache para respostas da API do Gemini,
reduzindo custos e latência ao evitar chamadas duplicadas para prompts similares.
"""

import json
import hashlib
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Dict, Any, List
from dataclasses import dataclass, asdict

from src.logging_config import get_logger

logger = get_logger(__name__)


@dataclass
class CacheEntry:
    """Representa uma entrada no cache de respostas."""
    cache_key: str
    prompt_preview: str  # Primeiros 100 chars do prompt (para debug)
    response: str
    timestamp: str
    hit_count: int
    ttl_expires: str
    prompt_hash: str


@dataclass
class CacheStats:
    """Estatísticas do cache."""
    total_entries: int
    total_hits: int
    total_misses: int
    hit_rate: float
    cache_size_bytes: int


class ResponseCache:
    """Gerenciador de cache de respostas do Gemini API."""
    
    def __init__(
        self,
        cache_dir: str,
        max_entries: int = 100,
        ttl_hours: int = 24,
        enabled: bool = True
    ):
        """
        Inicializa o cache de respostas.
        
        Args:
            cache_dir: Diretório para armazenar o cache
            max_entries: Número máximo de entradas no cache
            ttl_hours: Tempo de vida das entradas em horas
            enabled: Se o cache está habilitado
        """
        self.cache_dir = Path(cache_dir).resolve()
        self.max_entries = max_entries
        self.ttl_hours = ttl_hours
        self.enabled = enabled
        
        # Criar diretório de cache se não existir
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        
        self.cache_file = self.cache_dir / "gemini_cache.json"
        self.stats_file = self.cache_dir / "cache_stats.json"
        
        # Carregar cache e estatísticas
        self.cache_data: Dict[str, CacheEntry] = self._load_cache()
        self.stats = self._load_stats()
        self._stats_dirty = False
        
        # Limpar entradas expiradas na inicialização
        self._cleanup_expired()
        
        logger.info(
            f"ResponseCache inicializado: enabled={enabled}, "
            f"max_entries={max_entries}, ttl_hours={ttl_hours}, "
            f"current_entries={len(self.cache_data)}"
        )
    
    def get(self, prompt: str) -> Optional[str]:
        """
        Busca resposta no cache.
        
        Args:
            prompt: Prompt completo (incluindo contexto)
        
        Returns:
            Resposta cacheada ou None se não encontrada/expirada
        """
        if not self.enabled:
            return None
        
        cache_key = self._generate_cache_key(prompt)
        
        if cache_key not in self.cache_data:
            self.stats['total_misses'] += 1
            self._stats_dirty = True
            logger.debug(f"Cache miss: {cache_key[:16]}...")
            return None
        
        entry = self.cache_data[cache_key]
        
        # Verificar se expirou
        if self._is_expired(entry):
            logger.debug(f"Cache entry expired: {cache_key[:16]}...")
            del self.cache_data[cache_key]
            self._save_cache()
            self.stats['total_misses'] += 1
            self._stats_dirty = True
            return None
        
        # Cache hit!
        entry.hit_count += 1
        self.stats['total_hits'] += 1
        self._stats_dirty = True
        
        logger.info(
            f"Cache hit: {cache_key[:16]}... "
            f"(hit_count={entry.hit_count})"
        )
        
        return entry.response
    
    def set(self, prompt: str, response: str) -> None:
        """
        Salva resposta no cache.
        
        Args:
            prompt: Prompt completo
            response: Resposta do Gemini
        """
        if not self.enabled:
            return
        
        cache_key = self._generate_cache_key(prompt)
        
        # Criar entrada
        now = datetime.now()
        expires = now + timedelta(hours=self.ttl_hours)
        
        entry = CacheEntry(
            cache_key=cache_key,
            prompt_preview=prompt[:100],
            response=response,
            timestamp=now.isoformat(),
            hit_count=0,
            ttl_expires=expires.isoformat(),
            prompt_hash=cache_key
        )
        
        # Adicionar ao cache
        self.cache_data[cache_key] = entry
        
        # Verificar limite de entradas
        if len(self.cache_data) > self.max_entries:
            self._evict_entries()
        
        # Persistir cache e stats pendentes
        self._save_cache()
        self._flush_stats()
        
        logger.debug(f"Cache set: {cache_key[:16]}...")
    
    def invalidate(self) -> None:
        """Limpa todo o cache."""
        self.cache_data.clear()
        self._save_cache()
        logger.info("Cache invalidado completamente")
    
    def get_stats(self) -> CacheStats:
        """
        Retorna estatísticas do cache.
        
        Returns:
            CacheStats com informações do cache
        """
        total_requests = self.stats['total_hits'] + self.stats['total_misses']
        hit_rate = (
            self.stats['total_hits'] / total_requests
            if total_requests > 0
            else 0.0
        )
        
        # Calcular tamanho do cache
        cache_size = 0
        if self.cache_file.exists():
            cache_size = self.cache_file.stat().st_size
        
        return CacheStats(
            total_entries=len(self.cache_data),
            total_hits=self.stats['total_hits'],
            total_misses=self.stats['total_misses'],
            hit_rate=hit_rate,
            cache_size_bytes=cache_size
        )
    
    def get_recent_entries(self, limit: int = 10) -> List[Dict[str, Any]]:
        """
        Retorna entradas mais recentes do cache.
        
        Args:
            limit: Número máximo de entradas a retornar
        
        Returns:
            Lista de entradas ordenadas por timestamp (mais recentes primeiro)
        """
        entries = sorted(
            self.cache_data.values(),
            key=lambda e: e.timestamp,
            reverse=True
        )
        
        return [
            {
                'prompt_preview': e.prompt_preview,
                'timestamp': e.timestamp,
                'hit_count': e.hit_count,
                'expires': e.ttl_expires
            }
            for e in entries[:limit]
        ]
    
    def _generate_cache_key(self, prompt: str) -> str:
        """
        Gera chave de cache usando hash SHA256.
        
        Args:
            prompt: Prompt completo
        
        Returns:
            Hash SHA256 do prompt
        """
        return hashlib.sha256(prompt.encode('utf-8')).hexdigest()
    
    def _is_expired(self, entry: CacheEntry) -> bool:
        """
        Verifica se entrada expirou.
        
        Args:
            entry: Entrada do cache
        
        Returns:
            True se expirou
        """
        expires = datetime.fromisoformat(entry.ttl_expires)
        return datetime.now() > expires
    
    def _cleanup_expired(self) -> None:
        """Remove entradas expiradas do cache."""
        expired_keys = [
            key for key, entry in self.cache_data.items()
            if self._is_expired(entry)
        ]
        
        for key in expired_keys:
            del self.cache_data[key]
        
        if expired_keys:
            self._save_cache()
            logger.info(f"Removidas {len(expired_keys)} entradas expiradas")
    
    def _evict_entries(self) -> None:
        """
        Remove entradas menos usadas quando cache excede limite.
        
        Usa estratégia LRU baseada em hit_count e timestamp.
        """
        # Ordenar por hit_count (menor primeiro) e timestamp (mais antigo primeiro)
        entries = sorted(
            self.cache_data.items(),
            key=lambda x: (x[1].hit_count, x[1].timestamp)
        )
        
        # Remover 10% das entradas menos usadas
        to_remove = max(1, int(self.max_entries * 0.1))
        
        for key, _ in entries[:to_remove]:
            del self.cache_data[key]
        
        logger.info(f"Removidas {to_remove} entradas do cache (LRU)")

    
    def _load_cache(self) -> Dict[str, CacheEntry]:
        """
        Carrega cache do disco.
        
        Returns:
            Dicionário de entradas do cache
        """
        if not self.cache_file.exists():
            return {}
        
        try:
            with open(self.cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Converter dicionários em objetos CacheEntry
            cache = {}
            for key, entry_dict in data.items():
                cache[key] = CacheEntry(**entry_dict)
            
            logger.debug(f"Cache carregado: {len(cache)} entradas")
            return cache
        except Exception as e:
            logger.error(f"Erro ao carregar cache: {e}", exc_info=True)
            return {}
    
    def _save_cache(self) -> None:
        """Salva cache no disco."""
        try:
            # Converter objetos em dicionários
            data = {
                key: asdict(entry)
                for key, entry in self.cache_data.items()
            }
            
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            logger.debug("Cache salvo no disco")
        except Exception as e:
            logger.error(f"Erro ao salvar cache: {e}", exc_info=True)
    
    def _load_stats(self) -> Dict[str, int]:
        """
        Carrega estatísticas do disco.
        
        Returns:
            Dicionário com estatísticas
        """
        if not self.stats_file.exists():
            return {
                'total_hits': 0,
                'total_misses': 0
            }
        
        try:
            with open(self.stats_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Erro ao carregar estatísticas: {e}", exc_info=True)
            return {
                'total_hits': 0,
                'total_misses': 0
            }
    
    def _flush_stats(self) -> None:
        """Salva estatísticas no disco apenas se houver mudanças pendentes."""
        if self._stats_dirty:
            self._save_stats()
            self._stats_dirty = False
    
    def _save_stats(self) -> None:
        """Salva estatísticas no disco."""
        try:
            with open(self.stats_file, 'w', encoding='utf-8') as f:
                json.dump(self.stats, f, indent=2)
        except Exception as e:
            logger.error(f"Erro ao salvar estatísticas: {e}", exc_info=True)
