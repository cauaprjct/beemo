"""Testes para o módulo response_cache."""

import pytest
import tempfile
import shutil
from pathlib import Path
from datetime import datetime, timedelta

from src.response_cache import ResponseCache, CacheEntry, CacheStats


@pytest.fixture
def temp_dir():
    """Cria diretório temporário para testes."""
    temp_path = tempfile.mkdtemp()
    yield temp_path
    shutil.rmtree(temp_path, ignore_errors=True)


@pytest.fixture
def cache(temp_dir):
    """Cria ResponseCache para testes."""
    return ResponseCache(temp_dir, max_entries=5, ttl_hours=1)


def test_cache_initialization(temp_dir):
    """Testa inicialização do cache."""
    cache = ResponseCache(temp_dir, max_entries=10, ttl_hours=24)
    
    assert cache.cache_dir == Path(temp_dir).resolve()
    assert cache.max_entries == 10
    assert cache.ttl_hours == 24
    assert cache.enabled is True
    assert len(cache.cache_data) == 0


def test_cache_disabled(temp_dir):
    """Testa cache desabilitado."""
    cache = ResponseCache(temp_dir, enabled=False)
    
    # Set não deve fazer nada
    cache.set("test prompt", "test response")
    assert len(cache.cache_data) == 0
    
    # Get deve retornar None
    result = cache.get("test prompt")
    assert result is None


def test_cache_set_and_get(cache):
    """Testa salvamento e recuperação do cache."""
    prompt = "Create an Excel file with sales data"
    response = "I'll create the Excel file..."
    
    # Salvar no cache
    cache.set(prompt, response)
    
    # Recuperar do cache
    cached_response = cache.get(prompt)
    assert cached_response == response


def test_cache_miss(cache):
    """Testa cache miss."""
    result = cache.get("non-existent prompt")
    assert result is None
    
    stats = cache.get_stats()
    assert stats.total_misses == 1
    assert stats.total_hits == 0


def test_cache_hit_increments_counter(cache):
    """Testa que cache hit incrementa contador."""
    prompt = "test prompt"
    response = "test response"
    
    cache.set(prompt, response)
    
    # Primeiro hit
    cache.get(prompt)
    
    # Segundo hit
    cache.get(prompt)
    
    # Verificar hit_count
    cache_key = cache._generate_cache_key(prompt)
    entry = cache.cache_data[cache_key]
    assert entry.hit_count == 2
    
    stats = cache.get_stats()
    assert stats.total_hits == 2


def test_cache_different_prompts(cache):
    """Testa que prompts diferentes geram entradas diferentes."""
    cache.set("prompt 1", "response 1")
    cache.set("prompt 2", "response 2")
    
    assert cache.get("prompt 1") == "response 1"
    assert cache.get("prompt 2") == "response 2"
    assert len(cache.cache_data) == 2


def test_cache_eviction(cache):
    """Testa remoção de entradas quando excede limite."""
    # max_entries = 5, adicionar 7 entradas
    for i in range(7):
        cache.set(f"prompt {i}", f"response {i}")
    
    # Deve ter no máximo 5 entradas
    assert len(cache.cache_data) <= 5


def test_cache_invalidate(cache):
    """Testa limpeza completa do cache."""
    cache.set("prompt 1", "response 1")
    cache.set("prompt 2", "response 2")
    
    assert len(cache.cache_data) == 2
    
    cache.invalidate()
    
    assert len(cache.cache_data) == 0


def test_cache_persistence(temp_dir):
    """Testa persistência do cache entre instâncias."""
    # Criar cache e adicionar entrada
    cache1 = ResponseCache(temp_dir)
    cache1.set("test prompt", "test response")
    
    # Criar nova instância
    cache2 = ResponseCache(temp_dir)
    
    # Deve carregar entrada do disco
    result = cache2.get("test prompt")
    assert result == "test response"


def test_cache_stats(cache):
    """Testa estatísticas do cache."""
    cache.set("prompt 1", "response 1")
    cache.get("prompt 1")  # hit
    cache.get("prompt 2")  # miss
    
    stats = cache.get_stats()
    
    assert stats.total_entries == 1
    assert stats.total_hits == 1
    assert stats.total_misses == 1
    assert stats.hit_rate == 0.5


def test_get_recent_entries(cache):
    """Testa recuperação de entradas recentes."""
    for i in range(3):
        cache.set(f"prompt {i}", f"response {i}")
    
    recent = cache.get_recent_entries(limit=2)
    
    assert len(recent) == 2
    assert 'prompt_preview' in recent[0]
    assert 'timestamp' in recent[0]
    assert 'hit_count' in recent[0]


def test_cache_key_generation(cache):
    """Testa geração de chave de cache."""
    key1 = cache._generate_cache_key("test prompt")
    key2 = cache._generate_cache_key("test prompt")
    key3 = cache._generate_cache_key("different prompt")
    
    # Mesmo prompt deve gerar mesma chave
    assert key1 == key2
    
    # Prompts diferentes devem gerar chaves diferentes
    assert key1 != key3
    
    # Chave deve ser hash SHA256 (64 caracteres hex)
    assert len(key1) == 64


def test_cache_with_context(cache):
    """Testa que contexto diferente gera cache diferente."""
    prompt_base = "Update cell A1"
    context1 = "File: sales.xlsx\nContent: ..."
    context2 = "File: inventory.xlsx\nContent: ..."
    
    full_prompt1 = f"{prompt_base}\n{context1}"
    full_prompt2 = f"{prompt_base}\n{context2}"
    
    cache.set(full_prompt1, "response 1")
    cache.set(full_prompt2, "response 2")
    
    # Devem ser entradas diferentes
    assert cache.get(full_prompt1) == "response 1"
    assert cache.get(full_prompt2) == "response 2"
    assert len(cache.cache_data) == 2
