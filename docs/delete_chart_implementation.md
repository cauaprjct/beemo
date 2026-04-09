# Implementação de Deleção de Gráficos (delete_chart)

## Visão Geral

Implementada funcionalidade completa para remover gráficos de planilhas Excel, permitindo limpeza programática de gráficos antigos ou indesejados, completando o ciclo de gerenciamento de gráficos (criar, listar, deletar).

## Problema Resolvido

### Antes da Implementação
- Impossível remover gráficos programaticamente
- Usuário precisava abrir Excel manualmente para deletar gráficos
- Gráficos antigos se acumulavam em automações repetidas
- Sem forma de limpar gráficos obsoletos

### Depois da Implementação
- Remove gráficos por índice ou título
- Limpeza programática de gráficos antigos
- Integrado ao sistema de versionamento (undo/redo)
- Mensagens de erro claras e informativas

## Funcionalidades

### 1. Deletar Gráfico por Índice

```python
# Deletar primeiro gráfico (índice 0)
tool.delete_chart('vendas.xlsx', 'Sheet1', 0)

# Deletar segundo gráfico (índice 1)
tool.delete_chart('vendas.xlsx', 'Sheet1', 1)
```

**Vantagens:**
- Rápido e direto
- Útil em loops e automações
- Não depende de títulos

### 2. Deletar Gráfico por Título

```python
# Deletar gráfico específico por nome
tool.delete_chart('vendas.xlsx', 'Sheet1', 'Vendas por Produto')

# Deletar gráfico com caracteres especiais
tool.delete_chart('vendas.xlsx', 'Sheet1', 'Vendas 2025 (R$) - 100%')
```

**Vantagens:**
- Mais legível e intuitivo
- Não precisa saber o índice
- Busca case-sensitive (exata)

### 3. Workflow Integrado: Listar → Deletar

```python
# Listar gráficos existentes
result = tool.list_charts('vendas.xlsx', 'Sheet1')

# Encontrar gráfico específico
for chart in result['charts']:
    if 'Antigo' in chart['title']:
        # Deletar por índice ou título
        tool.delete_chart('vendas.xlsx', 'Sheet1', chart['index'])
```

## Uso via Linguagem Natural

### Exemplo 1: Deletar por Índice
```
"Delete o primeiro gráfico da Sheet1 do arquivo vendas.xlsx"
```

**Resposta JSON:**
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "delete_chart",
            "target_file": "vendas.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "identifier": 0
            }
        }
    ],
    "explanation": "Deletando o primeiro gráfico (índice 0) da Sheet1"
}
```

### Exemplo 2: Deletar por Título
```
"Remova o gráfico 'Vendas por Produto' do arquivo vendas.xlsx"
```

**Resposta JSON:**
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "delete_chart",
            "target_file": "vendas.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "identifier": "Vendas por Produto"
            }
        }
    ],
    "explanation": "Removendo o gráfico 'Vendas por Produto' da Sheet1"
}
```

### Exemplo 3: Listar e Deletar
```
"Liste os gráficos do arquivo vendas.xlsx e delete o gráfico 'Gráfico Antigo'"
```

**Resposta JSON:**
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "list_charts",
            "target_file": "vendas.xlsx",
            "parameters": {
                "sheet": "Sheet1"
            }
        },
        {
            "tool": "excel",
            "operation": "delete_chart",
            "target_file": "vendas.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "identifier": "Gráfico Antigo"
            }
        }
    ],
    "explanation": "Listando gráficos e removendo 'Gráfico Antigo'"
}
```

## Casos de Uso

### Caso 1: Limpeza de Gráficos Antigos
```python
# Automação mensal que recria gráficos
# Deletar gráficos antigos antes de criar novos
result = tool.list_charts('relatorio_mensal.xlsx', 'Dashboard')

for chart in result['charts']:
    if 'Mês Anterior' in chart['title']:
        tool.delete_chart('relatorio_mensal.xlsx', 'Dashboard', chart['index'])

# Criar novos gráficos atualizados
tool.add_chart('relatorio_mensal.xlsx', 'Dashboard', new_chart_config)
```

### Caso 2: Substituir Gráfico Específico
```python
# Substituir gráfico mantendo a posição
result = tool.list_charts('vendas.xlsx', 'Sheet1')

# Encontrar gráfico a substituir
old_chart = next(c for c in result['charts'] if c['title'] == 'Vendas Q1')
position = old_chart['position']

# Deletar antigo
tool.delete_chart('vendas.xlsx', 'Sheet1', 'Vendas Q1')

# Criar novo na mesma posição
new_config = {
    'type': 'column',
    'title': 'Vendas Q2',
    'categories': 'A2:A10',
    'values': 'B2:B10',
    'position': position
}
tool.add_chart('vendas.xlsx', 'Sheet1', new_config)
```

### Caso 3: Deletar Todos os Gráficos
```python
# Limpar todos os gráficos de uma sheet
result = tool.list_charts('vendas.xlsx', 'Sheet1')

# Deletar todos (sempre deletando o índice 0)
for _ in range(result['total_count']):
    tool.delete_chart('vendas.xlsx', 'Sheet1', 0)
```

### Caso 4: Deletar Gráficos por Padrão
```python
# Deletar todos os gráficos que contêm "Teste" no título
result = tool.list_charts('vendas.xlsx', 'Sheet1')

for chart in result['charts']:
    if 'Teste' in chart['title']:
        tool.delete_chart('vendas.xlsx', 'Sheet1', chart['title'])
```

## Implementação Técnica

### Método Principal: `delete_chart()`

```python
def delete_chart(self, file_path: str, sheet: str, identifier: Any) -> None:
    """
    Remove um gráfico de uma planilha Excel.
    
    Args:
        file_path: Path to the Excel file
        sheet: Name of the sheet containing the chart
        identifier: Chart identifier - can be:
                   - int: Index of chart in the sheet (0-based)
                   - str: Title of the chart to delete
    
    Raises:
        FileNotFoundError: If the file does not exist
        ValueError: If the sheet doesn't exist or chart not found
        CorruptedFileError: If the file is corrupted
        IOError: If the chart cannot be deleted
    """
```

**Funcionamento:**
1. Valida que o arquivo existe
2. Abre o workbook e valida que a sheet existe
3. Determina o tipo de identificador (int ou str)
4. Se int: valida índice e obtém gráfico diretamente
5. Se str: busca gráfico por título (case-sensitive)
6. Extrai informações do gráfico para logging
7. Remove gráfico da worksheet
8. Salva o arquivo
9. Loga sucesso com detalhes

### Validação de Identificador

**Por Índice (int):**
```python
if isinstance(identifier, int):
    if identifier < 0 or identifier >= len(worksheet._charts):
        raise ValueError(
            f"Chart index {identifier} out of range. "
            f"Sheet has {len(worksheet._charts)} chart(s)"
        )
    chart_to_delete = worksheet._charts[identifier]
```

**Por Título (str):**
```python
elif isinstance(identifier, str):
    for index, chart in enumerate(worksheet._charts):
        chart_title = self._get_chart_title(chart)
        if chart_title and chart_title == identifier:
            chart_to_delete = chart
            break
    
    if chart_to_delete is None:
        # Lista gráficos disponíveis para mensagem de erro
        available_titles = [...]
        raise ValueError(
            f"Chart with title '{identifier}' not found. "
            f"Available charts: {', '.join(available_titles)}"
        )
```

### Mensagens de Erro Informativas

**Índice fora do intervalo:**
```
ValueError: Chart index 5 out of range. Sheet has 3 chart(s) (indices 0-2)
```

**Título não encontrado:**
```
ValueError: Chart with title 'Gráfico Inexistente' not found in sheet 'Sheet1'. 
Available charts: Vendas por Produto, Distribuição, Untitled
```

**Tipo de identificador inválido:**
```
ValueError: Invalid identifier type: float. Must be int (index) or str (title)
```

## Integração com Sistema

### 1. Versionamento (Undo/Redo)

A operação `delete_chart` está integrada ao sistema de versionamento:

```python
# Em src/agent.py
if self.version_manager and operation in [..., 'delete_chart', ...]:
    version_id = self.version_manager.create_backup(
        target_file, operation, user_prompt
    )
```

**Benefício:** Se deletar um gráfico por engano, pode usar undo para restaurar!

### 2. Validação de Segurança

```python
# Em src/security_validator.py
ALLOWED_OPERATIONS = {
    ...,
    'delete_chart',
    ...
}
```

### 3. Validação de Parâmetros

```python
# Em src/response_parser.py
elif operation == 'delete_chart':
    if 'identifier' not in parameters:
        raise ValidationError("delete_chart requires 'identifier' parameter")
    
    identifier = parameters['identifier']
    if not isinstance(identifier, (int, str)):
        raise ValidationError("'identifier' must be int or str")
```

### 4. Documentação no Prompt

Adicionada seção completa "DELETE CHART OPERATIONS" no `src/prompt_templates.py` com:
- 3 exemplos detalhados
- Explicação de parâmetros
- Dicas de uso
- Integração com list_charts

## Testes

Criado arquivo `tests/test_delete_chart.py` com 18 testes:

### Testes de Funcionalidade (6 testes)
1. ✅ Deletar primeiro gráfico por índice
2. ✅ Deletar gráfico do meio por índice
3. ✅ Deletar último gráfico por índice
4. ✅ Deletar gráfico por título
5. ✅ Deletar todos os gráficos sequencialmente
6. ✅ Verificar que deleção persiste após reabrir arquivo

### Testes de Validação (7 testes)
7. ✅ Erro para índice negativo
8. ✅ Erro para índice muito alto
9. ✅ Erro para título inexistente (com lista de disponíveis)
10. ✅ Erro para sheet inexistente
11. ✅ Erro para arquivo inexistente
12. ✅ Erro para tipo de identificador inválido
13. ✅ Erro ao deletar de sheet sem gráficos

### Testes de Casos Extremos (3 testes)
14. ✅ Deletar gráfico com caracteres especiais no título
15. ✅ Busca por título é case-sensitive
16. ✅ Índices são atualizados após deleção

### Testes de Integração (2 testes)
17. ✅ Deletar e depois adicionar novo gráfico
18. ✅ Workflow de listar e depois deletar

**Resultado:** 18/18 testes passando (100%) ✅

## Benefícios

### Para Automação
- ✅ Limpeza programática de gráficos antigos
- ✅ Evita acúmulo de gráficos em processos repetitivos
- ✅ Permite substituição controlada de gráficos
- ✅ Integrado ao ciclo completo: criar → listar → deletar

### Para Manutenção
- ✅ Remove gráficos obsoletos facilmente
- ✅ Limpeza de gráficos de teste
- ✅ Reorganização de dashboards

### Para Usuários
- ✅ Comandos simples em linguagem natural
- ✅ Flexibilidade: deletar por índice ou título
- ✅ Mensagens de erro claras e úteis
- ✅ Undo/redo disponível

## Integração com Outras Funcionalidades

### Ciclo Completo de Gerenciamento de Gráficos

```python
# 1. Criar gráfico
tool.add_chart('vendas.xlsx', 'Sheet1', chart_config)

# 2. Listar gráficos
result = tool.list_charts('vendas.xlsx', 'Sheet1')

# 3. Deletar gráfico
tool.delete_chart('vendas.xlsx', 'Sheet1', 0)
```

### Com Validação de Posição

```python
# Deletar gráfico em posição específica
result = tool.list_charts('vendas.xlsx', 'Sheet1')

for chart in result['charts']:
    if chart['position'] == 'H2':
        tool.delete_chart('vendas.xlsx', 'Sheet1', chart['index'])
        break

# Adicionar novo gráfico na posição liberada
tool.add_chart('vendas.xlsx', 'Sheet1', new_config)
```

## Limitações e Considerações

### Limitações
- Busca por título é case-sensitive (exata)
- Não suporta regex ou busca parcial por título
- Não deleta múltiplos gráficos em uma chamada

### Considerações de Performance
- Operação rápida (< 100ms para arquivos típicos)
- Salva arquivo após cada deleção
- Para deletar múltiplos gráficos, use loop

### Boas Práticas
- Use `list_charts` antes de deletar para verificar gráficos existentes
- Delete por título quando possível (mais legível)
- Delete por índice em loops (sempre deletando índice 0)
- Verifique mensagens de erro para gráficos disponíveis

## Compatibilidade

- ✅ Funciona com todos os tipos de gráficos
- ✅ Funciona com gráficos com ou sem título
- ✅ Integrado ao sistema de versionamento
- ✅ Compatível com arquivos criados no Excel ou pelo projeto
- ✅ Não quebra referências de dados

## Próximos Passos Possíveis

Esta implementação completa o ciclo básico de gerenciamento de gráficos. Melhorias futuras poderiam incluir:

1. 🔜 **update_chart** - Modificar gráficos existentes (título, dados, posição)
2. 🔜 **move_chart** - Mover gráfico para nova posição
3. 🔜 **copy_chart** - Copiar gráfico para outra sheet
4. 🔜 **delete_charts_batch** - Deletar múltiplos gráficos em uma chamada

## Resumo da Implementação

**Status:** ✅ COMPLETO

- [x] Método `delete_chart()` implementado
- [x] Suporte para identificador por índice (int)
- [x] Suporte para identificador por título (str)
- [x] Validação completa de parâmetros
- [x] Mensagens de erro informativas
- [x] Integração no agent
- [x] Integração no response_parser
- [x] Integração no security_validator
- [x] Integração no versionamento
- [x] Documentação no prompt (3 exemplos)
- [x] Capacidade adicionada à lista
- [x] Testes criados (18 testes)
- [x] Todos os testes passando (100%)
- [x] Documentação técnica criada

---

**Implementado em:** 28/03/2026
**Versão:** 1.5.0
**Status:** ✅ Produção
**Testes:** 18/18 passando (100%)
**Esforço:** 2-3 horas
**Valor:** Alto (completa ciclo de gerenciamento de gráficos)
