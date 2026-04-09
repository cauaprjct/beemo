# Implementação de Cache de Respostas do Gemini

## Visão Geral

Sistema de cache inteligente que armazena respostas da API do Gemini para evitar chamadas duplicadas, reduzindo custos e latência.

## Arquitetura

### Componentes Principais

1. **ResponseCache** (`src/response_cache.py`)
   - Gerencia cache de respostas
   - Armazena em disco para persistência
   - Implementa TTL e LRU eviction
   - Rastreia estatísticas de uso

2. **Integração com GeminiClient** (`src/gemini_client.py`)
   - Verifica cache antes de chamar API
   - Salva respostas no cache após API call
   - Transparente para o usuário

3. **Interface Streamlit** (`app.py`)
   - Seção "Cache de Respostas" na sidebar
   - Estatísticas em tempo real
   - Controle para limpar cache

### Estrutura de Dados

```python
CacheEntry:
  - cache_key: hash SHA256 do prompt
  - prompt_preview: primeiros 100 chars (para debug)
  - response: resposta do Gemini
  - timestamp: quando foi criado
  - hit_count: quantas vezes foi usado
  - ttl_expires: quando expira
  - prompt_hash: hash completo

CacheStats:
  - total_entries: número de entradas
  - total_hits: cache hits
  - total_misses: cache misses
  - hit_rate: taxa de acerto
  - cache_size_bytes: tamanho em bytes
```

### Armazenamento

```
ROOT_PATH/
├── .cache/
│   ├── gemini_cache.json      # Entradas do cache
│   └── cache_stats.json       # Estatísticas
└── [arquivos do usuário]
```

## Fluxo de Operação

### Cache Hit (Resposta Encontrada)

1. Usuário envia prompt
2. Agent constrói prompt completo com contexto
3. GeminiClient calcula hash SHA256 do prompt
4. ResponseCache verifica se hash existe
5. Verifica se não expirou (TTL)
6. Incrementa hit_count
7. Retorna resposta cacheada
8. ⚡ Resposta instantânea!

### Cache Miss (Resposta Não Encontrada)

1. Usuário envia prompt
2. Agent constrói prompt completo com contexto
3. GeminiClient calcula hash SHA256 do prompt
4. ResponseCache não encontra entrada
5. GeminiClient chama API do Gemini
6. Recebe resposta da API
7. ResponseCache salva resposta
8. Retorna resposta ao usuário

### Expiração e Limpeza

1. **TTL (Time To Live)**: Entradas expiram após X horas
2. **Cleanup na inicialização**: Remove entradas expiradas
3. **Verificação no get**: Checa TTL antes de retornar
4. **LRU Eviction**: Remove menos usadas quando excede limite

## Estratégia de Cache

### Hash do Prompt

O cache usa SHA256 do prompt **completo**, incluindo:
- Prompt do usuário
- Contexto dos arquivos
- Conteúdo dos arquivos

Isso garante que:
- Mudanças nos arquivos invalidam o cache
- Contextos diferentes geram entradas diferentes
- Colisões são praticamente impossíveis

### Exemplo

```python
# Prompt 1: "Atualizar célula A1" + contexto de vendas.xlsx
hash1 = sha256("Atualizar célula A1\nFile: vendas.xlsx\nContent: ...")

# Prompt 2: "Atualizar célula A1" + contexto de estoque.xlsx  
hash2 = sha256("Atualizar célula A1\nFile: estoque.xlsx\nContent: ...")

# hash1 != hash2 → Entradas diferentes no cache
```

## Configuração

### Variáveis de Ambiente

```env
CACHE_ENABLED=true          # Habilitar/desabilitar cache
CACHE_TTL_HOURS=24          # Tempo de vida em horas
CACHE_MAX_ENTRIES=100       # Número máximo de entradas
```

### Valores Padrão

- Habilitado: `true`
- TTL: 24 horas
- Max entries: 100
- Eviction: 10% das entradas menos usadas

## Uso

### Programático

```python
from src.response_cache import ResponseCache

# Criar cache
cache = ResponseCache(
    cache_dir="/path/to/cache",
    max_entries=100,
    ttl_hours=24,
    enabled=True
)

# Buscar no cache
response = cache.get(prompt)
if response:
    print("Cache hit!")
else:
    # Chamar API
    response = api_call(prompt)
    # Salvar no cache
    cache.set(prompt, response)

# Estatísticas
stats = cache.get_stats()
print(f"Hit rate: {stats.hit_rate:.1%}")

# Limpar cache
cache.invalidate()
```

### Interface Streamlit

1. Veja estatísticas na sidebar
2. Monitore taxa de hit
3. Veja entradas recentes
4. Limpe o cache quando necessário

## Estatísticas

### Métricas Rastreadas

- **Total de entradas**: Número atual de entradas no cache
- **Total de hits**: Quantas vezes o cache foi usado
- **Total de misses**: Quantas vezes não encontrou no cache
- **Taxa de hit**: hits / (hits + misses)
- **Tamanho do cache**: Espaço em disco usado

### Interpretação

- **Taxa de hit > 50%**: Cache está sendo efetivo
- **Taxa de hit < 20%**: Considere aumentar TTL ou max_entries
- **Muitas entradas**: Considere reduzir TTL
- **Poucas entradas**: Usuário não repete operações

## Testes

```bash
# Testes do ResponseCache
pytest tests/test_response_cache.py -v

# Testes de integração
pytest tests/test_gemini_client.py -v
```

### Cobertura de Testes

- ✅ Inicialização do cache
- ✅ Cache habilitado/desabilitado
- ✅ Set e get de entradas
- ✅ Cache hit e miss
- ✅ Incremento de hit_count
- ✅ Prompts diferentes
- ✅ Eviction quando excede limite
- ✅ Invalidação completa
- ✅ Persistência entre sessões
- ✅ Estatísticas
- ✅ Entradas recentes
- ✅ Geração de chave
- ✅ Cache com contexto diferente

## Limitações

1. **Cache exato**: Apenas prompts idênticos são cacheados
   - Prompts similares mas não idênticos não são reconhecidos
   - Futuro: implementar cache por similaridade com embeddings

2. **Espaço em disco**: Cache pode crescer
   - Limite de max_entries ajuda
   - Considere monitorar espaço em disco

3. **Dados sensíveis**: Cache armazena prompts e respostas
   - Não use com dados confidenciais
   - Ou desabilite o cache

4. **Concorrência**: Não há proteção contra múltiplas instâncias
   - Para MVP, assumir instância única
   - Futuro: implementar file locking

5. **Invalidação manual**: Não há invalidação automática
   - Usuário deve limpar cache manualmente se necessário
   - TTL ajuda a manter cache atualizado

## Melhorias Futuras

- [ ] Cache por similaridade usando embeddings
- [ ] Compressão de entradas para economizar espaço
- [ ] Invalidação seletiva (por arquivo, por tipo de operação)
- [ ] Exportar/importar cache
- [ ] Sincronização entre múltiplas instâncias
- [ ] Cache distribuído (Redis, Memcached)
- [ ] Análise de padrões de uso
- [ ] Sugestões de otimização baseadas em estatísticas
- [ ] Cache warming (pré-carregar entradas comuns)
- [ ] Versionamento de cache (invalidar quando modelo muda)

## Segurança

### Considerações

1. **Dados sensíveis**: Cache pode expor informações
   - Solução: Desabilitar cache para dados sensíveis
   - Ou implementar criptografia

2. **Acesso ao cache**: Arquivos JSON são legíveis
   - Solução: Permissões de arquivo apropriadas
   - Ou criptografar cache no disco

3. **Injeção de cache**: Manipulação maliciosa do cache
   - Solução: Validar integridade ao carregar
   - Ou usar assinatura digital

## Arquivos Modificados

### Novos Arquivos
- `src/response_cache.py` - Implementação do cache
- `tests/test_response_cache.py` - Testes unitários
- `docs/cache_implementation.md` - Esta documentação

### Arquivos Modificados
- `config/__init__.py` - Adicionadas configurações de cache
- `src/factory.py` - Inicialização do ResponseCache
- `src/gemini_client.py` - Integração com cache
- `src/agent.py` - Passar cache para GeminiClient
- `app.py` - UI de estatísticas de cache
- `.env.example` - Adicionadas variáveis de cache
- `README.md` - Documentação da funcionalidade

## Conclusão

O sistema de cache está totalmente funcional e integrado ao Gemini Office Agent. Reduz custos de API e melhora a experiência do usuário com respostas instantâneas para prompts repetidos.
