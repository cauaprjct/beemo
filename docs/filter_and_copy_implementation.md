# 🔍 Funcionalidade de Filtrar e Copiar Dados (filter_and_copy)

## 📋 Resumo

Implementada funcionalidade completa para filtrar dados de planilhas Excel baseado em critérios específicos e copiar os resultados para nova sheet ou arquivo, permitindo extração e segmentação programática de dados.

## 🎯 Objetivo

Permitir que usuários filtrem e copiem dados de planilhas Excel através de comandos em linguagem natural, como:
- "Copie apenas vendas acima de R$5000 para nova sheet"
- "Filtre clientes com email Gmail para novo arquivo"
- "Extraia produtos da categoria Electronics"

## ✨ Características Implementadas

### 1. Método `filter_and_copy()` em `src/excel_tool.py`

**Capacidades:**
- ✅ Filtrar por operadores numéricos (>, <, >=, <=)
- ✅ Filtrar por igualdade/desigualdade (==, !=)
- ✅ Filtrar por texto (contains, starts_with, ends_with)
- ✅ Copiar para nova sheet no mesmo arquivo
- ✅ Copiar para novo arquivo Excel
- ✅ Suporte para linha de cabeçalho
- ✅ Opção de copiar ou não o cabeçalho
- ✅ Retorna estatísticas detalhadas
- ✅ Case-insensitive para comparações de texto

**Parâmetros:**
```python
config = {
    'column': 'F',                    # Coluna para filtrar (letra ou número) - OBRIGATÓRIO
    'operator': '>',                  # Operador de comparação - OBRIGATÓRIO
    'value': 5000,                    # Valor para comparar - OBRIGATÓRIO
    'destination_sheet': 'High Sales', # Nome da sheet destino (mesmo arquivo)
    'destination_file': 'output.xlsx', # Ou caminho do arquivo destino
    'has_header': True,               # Se primeira linha é cabeçalho - padrão: True
    'copy_header': True               # Se deve copiar cabeçalho - padrão: True
}
```

**Operadores Suportados:**
- Numéricos: `>`, `<`, `>=`, `<=`
- Igualdade: `==`, `!=`
- Texto: `contains`, `starts_with`, `ends_with`

**Retorno:**
```python
{
    'filtered_count': 4,        # Número de linhas que passaram no filtro
    'destination': 'High Sales', # Onde os dados foram copiados
    'destination_type': 'sheet'  # 'sheet' ou 'file'
}
```

### 2. Método Auxiliar `_apply_filter()`

Aplica o operador de filtro ao valor da célula, com tratamento robusto de tipos e conversões automáticas.

### 3. Validação em `src/response_parser.py`

**Validações adicionadas:**
- ✅ Verifica se 'config' está presente (obrigatório)
- ✅ Valida campos obrigatórios: column, operator, value
- ✅ Valida operador contra lista de operadores válidos
- ✅ Valida que destination_sheet OU destination_file está presente
- ✅ Validação integrada ao sistema existente

### 4. Integração em `src/agent.py`

```python
elif operation == 'filter_and_copy':
    source_sheet = parameters.get('sheet')
    config = parameters.get('config', {})
    result = self.excel_tool.filter_and_copy(target_file, source_sheet, config)
    logger.info(f"Filtered {result['filtered_count']} row(s)")
```

### 5. Segurança em `src/security_validator.py`

- ✅ Adicionado 'filter_and_copy' à lista `ALLOWED_OPERATIONS`

### 6. Documentação em `src/prompt_templates.py`

**Adicionado:**
- ✅ "Filter and copy data (by criteria to new sheet/file)" na lista de capacidades
- ✅ Seção completa "FILTER AND COPY OPERATIONS" com 3 exemplos detalhados
- ✅ Dicas de uso para filtrar e copiar
- ✅ Explicação de todos os operadores e parâmetros

## 🧪 Testes Implementados

**Arquivo:** `tests/test_filter_and_copy.py`

### Testes de Operadores Numéricos (4 testes)
1. ✅ Filtrar maior que (>)
2. ✅ Filtrar menor que (<)
3. ✅ Filtrar maior ou igual (>=)
4. ✅ Filtrar menor ou igual (<=)

### Testes de Igualdade (2 testes)
5. ✅ Filtrar igual (==)
6. ✅ Filtrar diferente (!=)

### Testes de Texto (3 testes)
7. ✅ Filtrar contém (contains)
8. ✅ Filtrar começa com (starts_with)
9. ✅ Filtrar termina com (ends_with)

### Testes de Destino (2 testes)
10. ✅ Copiar para novo arquivo
11. ✅ Copiar sem cabeçalho

### Testes de Casos Extremos (4 testes)
12. ✅ Nenhuma linha corresponde ao filtro
13. ✅ Todas as linhas correspondem
14. ✅ Planilha vazia
15. ✅ Sobrescrever sheet existente

### Testes de Validação (10 testes)
16. ✅ Erro para coluna ausente
17. ✅ Erro para operador ausente
18. ✅ Erro para valor ausente
19. ✅ Erro para operador inválido
20. ✅ Erro para destino ausente
21. ✅ Erro para ambos destinos especificados
22. ✅ Erro para tipo de coluna inválido
23. ✅ Erro para coluna fora do intervalo
24. ✅ Erro para sheet inexistente
25. ✅ Erro para arquivo não encontrado

### Testes de Integração (2 testes)
26. ✅ Filtrar por número de coluna
27. ✅ Filtrar case-insensitive

**Resultado:** 27/27 testes passando (100%) ✅

## 📊 Exemplos de Uso

### Exemplo 1: Filtrar valores numéricos para nova sheet
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "filter_and_copy",
            "target_file": "vendas.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "config": {
                    "column": "F",
                    "operator": ">",
                    "value": 5000,
                    "destination_sheet": "Vendas Altas",
                    "has_header": true
                }
            }
        }
    ],
    "explanation": "Filtrando vendas maiores que R$5000 para nova sheet"
}
```

### Exemplo 2: Filtrar texto para novo arquivo
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "filter_and_copy",
            "target_file": "clientes.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "config": {
                    "column": "C",
                    "operator": "contains",
                    "value": "@gmail.com",
                    "destination_file": "clientes_gmail.xlsx",
                    "has_header": true
                }
            }
        }
    ],
    "explanation": "Filtrando clientes com email Gmail para novo arquivo"
}
```

### Exemplo 3: Filtrar por categoria
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "filter_and_copy",
            "target_file": "produtos.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "config": {
                    "column": "D",
                    "operator": "==",
                    "value": "Electronics",
                    "destination_sheet": "Eletrônicos",
                    "has_header": true
                }
            }
        }
    ],
    "explanation": "Extraindo apenas produtos da categoria Electronics"
}
```

## 🔍 Casos de Uso Reais

### Caso 1: Segmentação de Vendas
```python
# Separar vendas por faixa de valor
tool.filter_and_copy('vendas.xlsx', 'Sheet1',
    {'column': 'F', 'operator': '>', 'value': 10000,
     'destination_sheet': 'VIP', 'has_header': True})

tool.filter_and_copy('vendas.xlsx', 'Sheet1',
    {'column': 'F', 'operator': '<=', 'value': 1000,
     'destination_sheet': 'Pequenas', 'has_header': True})
```

### Caso 2: Extração de Clientes por Domínio
```python
# Extrair clientes corporativos
tool.filter_and_copy('clientes.xlsx', 'Sheet1',
    {'column': 'C', 'operator': 'ends_with', 'value': '.com.br',
     'destination_file': 'clientes_br.xlsx', 'has_header': True})
```

### Caso 3: Filtrar Produtos em Estoque Baixo
```python
# Alertar sobre produtos com estoque baixo
tool.filter_and_copy('estoque.xlsx', 'Sheet1',
    {'column': 'D', 'operator': '<', 'value': 10,
     'destination_sheet': 'Estoque Baixo', 'has_header': True})
```

## 🔍 Considerações Técnicas

### Estratégia de Filtragem

**Operadores Numéricos:**
- Converte valores para float antes de comparar
- Funciona com inteiros, decimais e datas (como números)
- Retorna False se conversão falhar

**Operadores de Igualdade:**
- Compara valores diretamente (type-aware)
- Funciona com qualquer tipo de dado
- None é tratado corretamente

**Operadores de Texto:**
- Converte ambos os valores para string
- Case-insensitive (converte para lowercase)
- Funciona com qualquer tipo que possa ser convertido para string

### Preservação de Dados
- Copia linha inteira quando filtro corresponde
- Todas as colunas são preservadas
- Formatação não é copiada (apenas valores)
- Fórmulas são convertidas para valores

### Destino: Sheet vs File

**destination_sheet:**
- Cria nova sheet no mesmo arquivo
- Sobrescreve se sheet já existir
- Mais rápido (um único arquivo)
- Útil para organização interna

**destination_file:**
- Cria novo arquivo Excel
- Arquivo independente
- Útil para distribuição
- Sheet tem mesmo nome da origem

### Performance
- Lê todos os dados na memória
- Filtragem O(n) onde n = número de linhas
- Eficiente para planilhas de tamanho médio (milhares de linhas)
- Para planilhas muito grandes (>100k linhas), considerar processamento em lotes

### Tratamento de Erros
- `FileNotFoundError`: Arquivo não existe
- `ValueError`: Parâmetros inválidos (coluna, operador, destino)
- `CorruptedFileError`: Arquivo Excel corrompido
- `IOError`: Falha na operação de filtrar/copiar

## 📚 Arquivos Modificados/Criados

1. ✅ `src/excel_tool.py` - Adicionado método `filter_and_copy()` e `_apply_filter()` (~250 linhas)
2. ✅ `src/agent.py` - Integração da operação filter_and_copy
3. ✅ `src/response_parser.py` - Adicionado 'filter_and_copy' e validação completa
4. ✅ `src/security_validator.py` - Adicionado 'filter_and_copy' a `ALLOWED_OPERATIONS`
5. ✅ `src/prompt_templates.py` - Documentação completa com 3 exemplos
6. ✅ `tests/test_filter_and_copy.py` - Suite completa de testes (27 testes)
7. ✅ `docs/filter_and_copy_implementation.md` - Documentação técnica detalhada

## 📊 Métricas

- **Esforço:** 3-4 horas
- **Linhas de código:** ~250 linhas
- **Testes:** 27 testes (100% passando)
- **Cobertura:** Completa (operadores, destinos, edge cases, validação, integração)
- **Valor:** Alto (extração de dados é muito comum)

## 🎁 Benefícios

### Para Análise de Dados
- ✅ Segmentação rápida de dados
- ✅ Extração de subconjuntos relevantes
- ✅ Preparação de dados para análise
- ✅ Criação de relatórios específicos

### Para Automação
- ✅ Separação automática de dados por critérios
- ✅ Geração de arquivos para diferentes departamentos
- ✅ Alertas baseados em condições (estoque baixo, vendas altas)
- ✅ Distribuição de dados filtrados

### Para Usuários
- ✅ Comandos simples em linguagem natural
- ✅ Flexibilidade: múltiplos operadores e tipos de dados
- ✅ Destino flexível: sheet ou arquivo
- ✅ Feedback: estatísticas do resultado

## 🎯 Comandos em Linguagem Natural

O agente entende comandos como:
- "Copie apenas vendas acima de R$5000 para nova sheet 'Vendas Altas'"
- "Filtre clientes com email Gmail para arquivo clientes_gmail.xlsx"
- "Extraia produtos da categoria Electronics"
- "Separe transações maiores que 1000 para nova planilha"
- "Crie arquivo com apenas linhas onde Status é 'Ativo'"

## ✅ Status da Implementação

**Status:** ✅ COMPLETO

- [x] Método `filter_and_copy()` implementado
- [x] Método auxiliar `_apply_filter()` implementado
- [x] Suporte para operadores numéricos
- [x] Suporte para operadores de igualdade
- [x] Suporte para operadores de texto
- [x] Copiar para nova sheet
- [x] Copiar para novo arquivo
- [x] Suporte para cabeçalho
- [x] Validação completa de parâmetros
- [x] Integração no agent
- [x] Integração no response_parser
- [x] Integração no security_validator
- [x] Documentação no prompt (3 exemplos)
- [x] Capacidade adicionada à lista
- [x] Testes criados (27 testes)
- [x] Todos os testes passando (100%)
- [x] Documentação técnica criada

## 🚀 Como Testar

### Teste Manual
```bash
streamlit run app.py
```

**Comandos de teste:**
```
1. "Copie vendas acima de 5000 para nova sheet 'Vendas Altas' no arquivo vendas.xlsx"
2. "Filtre clientes com email Gmail para arquivo clientes_gmail.xlsx"
3. "Extraia produtos da categoria Electronics para nova sheet"
```

**Lembre-se:** Feche o Excel antes de executar modificações (Windows file locking).

### Testes Automatizados
```bash
python -m pytest tests/test_filter_and_copy.py -v
```

**Resultado esperado:** 27 passed ✅

## 🔗 Integração com Outras Funcionalidades

### Workflow: Sort + Filter + Copy
```python
# 1. Ordenar por valor
tool.sort_data('vendas.xlsx', 'Sheet1', {'column': 'F', 'order': 'desc'})

# 2. Filtrar top vendas
tool.filter_and_copy('vendas.xlsx', 'Sheet1',
    {'column': 'F', 'operator': '>', 'value': 5000,
     'destination_sheet': 'Top Vendas'})
```

### Workflow: Filter + Remove Duplicates
```python
# 1. Filtrar categoria
tool.filter_and_copy('produtos.xlsx', 'Sheet1',
    {'column': 'C', 'operator': '==', 'value': 'Electronics',
     'destination_sheet': 'Electronics'})

# 2. Remover duplicatas por ID
tool.remove_duplicates('produtos.xlsx', 'Electronics',
    {'column': 'A', 'keep': 'first'})
```

---

**Implementado em:** 28/03/2026
**Versão:** 1.7.0
**Status:** ✅ Produção
**Testes:** 27/27 passando (100%)
