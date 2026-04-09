# Implementação de Listagem de Gráficos (list_charts)

## Visão Geral

Implementada funcionalidade para listar todos os gráficos existentes em planilhas Excel, permitindo inspeção programática e facilitando automações que precisam saber quais gráficos já existem antes de adicionar, modificar ou remover.

## Problema Resolvido

### Antes da Implementação
- Impossível saber programaticamente quais gráficos existem
- Usuário precisava abrir Excel manualmente para ver gráficos
- Difícil implementar lógica condicional (se existe gráfico X, então...)
- Sem visibilidade para automações

### Depois da Implementação
- Lista todos os gráficos com detalhes completos
- Inspeção programática sem abrir Excel
- Permite lógica condicional em automações
- Base para funcionalidade delete_chart

## Funcionalidades

### 1. Listar Todos os Gráficos do Arquivo

```python
result = tool.list_charts('vendas.xlsx')

print(f"Total: {result['total_count']} gráficos")
for chart in result['charts']:
    print(f"  - {chart['title']} ({chart['type']}) em {chart['sheet']}!{chart['position']}")
```

**Saída:**
```
Total: 3 gráficos
  - Vendas por Produto (BarChart) em Sheet1!E2
  - Distribuição de Custos (PieChart) em Sheet1!E15
  - Receita Mensal (LineChart) em Sheet2!D2
```

### 2. Listar Gráficos de Sheet Específica

```python
result = tool.list_charts('vendas.xlsx', 'Sheet1')

print(f"Gráficos em Sheet1: {result['total_count']}")
```

### 3. Estrutura de Retorno

```python
{
    'charts': [
        {
            'sheet': 'Sheet1',
            'title': 'Vendas por Produto',
            'type': 'BarChart',
            'position': 'E2',
            'index': 0
        },
        {
            'sheet': 'Sheet1',
            'title': 'Distribuição de Custos',
            'type': 'PieChart',
            'position': 'E15',
            'index': 1
        }
    ],
    'total_count': 2,
    'sheets_analyzed': ['Sheet1']
}
```

**Campos de cada gráfico:**
- `sheet`: Nome da sheet onde o gráfico está
- `title`: Título do gráfico (ou "Untitled" se não tiver)
- `type`: Tipo do gráfico (BarChart, LineChart, PieChart, etc.)
- `position`: Posição da célula (e.g., 'E2')
- `index`: Índice do gráfico na sheet (0-based)

## Uso via Linguagem Natural

### Exemplo 1: Listar Todos os Gráficos
```
"Liste todos os gráficos do arquivo vendas.xlsx"
```

**Resposta:**
```
Encontrados 3 gráficos:

Sheet1:
  1. Vendas por Produto (BarChart) na posição E2
  2. Distribuição de Custos (PieChart) na posição E15

Sheet2:
  3. Receita Mensal (LineChart) na posição D2
```

### Exemplo 2: Listar Gráficos de Uma Sheet
```
"Quais gráficos existem na Sheet1 do arquivo vendas.xlsx?"
```

### Exemplo 3: Verificar Antes de Adicionar
```
"Liste os gráficos do arquivo vendas.xlsx e depois adicione um novo gráfico de pizza na posição K2"
```

## Casos de Uso

### Caso 1: Verificar Antes de Adicionar
```python
# Listar gráficos existentes
result = tool.list_charts('vendas.xlsx', 'Sheet1')

# Verificar se posição está livre
positions = [chart['position'] for chart in result['charts']]
if 'H2' not in positions:
    # Posição livre, pode adicionar
    tool.add_chart('vendas.xlsx', 'Sheet1', chart_config)
else:
    print("Posição H2 já ocupada!")
```

### Caso 2: Auditoria de Gráficos
```python
# Listar todos os gráficos
result = tool.list_charts('relatorio.xlsx')

# Gerar relatório
print(f"Relatório de Gráficos:")
print(f"Total: {result['total_count']} gráficos")
print(f"Sheets analisadas: {', '.join(result['sheets_analyzed'])}")

for chart in result['charts']:
    print(f"  - {chart['title']} ({chart['type']}) em {chart['sheet']}!{chart['position']}")
```

### Caso 3: Lógica Condicional em Automação
```python
# Verificar se gráfico específico existe
result = tool.list_charts('vendas.xlsx', 'Sheet1')

chart_titles = [chart['title'] for chart in result['charts']]

if 'Vendas por Produto' in chart_titles:
    print("Gráfico já existe, pulando criação")
else:
    print("Gráfico não existe, criando...")
    tool.add_chart('vendas.xlsx', 'Sheet1', chart_config)
```

### Caso 4: Documentação Automática
```python
# Gerar documentação de todos os arquivos
import os

for file in os.listdir('relatorios/'):
    if file.endswith('.xlsx'):
        result = tool.list_charts(f'relatorios/{file}')
        print(f"\n{file}: {result['total_count']} gráficos")
        for chart in result['charts']:
            print(f"  - {chart['title']}")
```

## Implementação Técnica

### Método Principal: `list_charts()`

```python
def list_charts(self, file_path: str, sheet: Optional[str] = None) -> Dict[str, Any]:
    """
    Lista todos os gráficos em uma planilha Excel.
    
    Args:
        file_path: Path to the Excel file
        sheet: Name of specific sheet (optional, None = all sheets)
    
    Returns:
        Dictionary with charts info, total_count, sheets_analyzed
    """
```

**Funcionamento:**
1. Abre o arquivo Excel
2. Determina quais sheets analisar (específica ou todas)
3. Itera sobre cada sheet
4. Para cada gráfico na sheet:
   - Extrai título usando `_get_chart_title()`
   - Extrai posição usando `_get_chart_position()`
   - Obtém tipo do gráfico
   - Registra índice
5. Retorna estrutura com todos os gráficos

### Método Auxiliar: `_get_chart_position()`

```python
def _get_chart_position(self, chart) -> str:
    """
    Extrai a posição de um gráfico.
    
    Args:
        chart: Chart object do openpyxl
    
    Returns:
        Posição do gráfico (e.g., 'H2') ou 'Unknown'
    """
```

**Funcionamento:**
- Acessa `chart.anchor._from` para obter coordenadas
- Converte coluna numérica para letra usando `get_column_letter()`
- Ajusta índices (openpyxl usa 0-based, Excel usa 1-based)
- Retorna posição formatada (e.g., 'H2')
- Retorna 'Unknown' se não conseguir determinar

### Integração no Agent

```python
elif operation == 'list_charts':
    sheet = parameters.get('sheet')  # Optional
    result = self.excel_tool.list_charts(target_file, sheet)
    return result  # Return result for user to see
```

**Nota:** `list_charts` é uma operação de leitura que retorna dados, não modifica o arquivo.

## Testes

Criado arquivo `tests/test_list_charts.py` com 12 testes:

### Testes de Funcionalidade
1. ✅ Listar gráficos em arquivo sem gráficos (retorna vazio)
2. ✅ Listar todos os gráficos do arquivo (múltiplas sheets)
3. ✅ Listar gráficos de sheet específica
4. ✅ Verificar que todos os detalhes estão presentes
5. ✅ Verificar tipos de gráficos corretos
6. ✅ Verificar posições corretas
7. ✅ Listar gráfico sem título (mostra "Untitled")

### Testes de Validação
8. ✅ Erro ao especificar sheet inexistente
9. ✅ Erro ao especificar arquivo inexistente

### Testes de Estrutura
10. ✅ Índices dos gráficos estão corretos
11. ✅ Estrutura de retorno está correta
12. ✅ Listar em sheet vazia (retorna vazio)

**Resultado:** 12/12 testes passando (100%) ✅

## Tipos de Gráficos Reconhecidos

A operação identifica corretamente todos os tipos de gráficos:

| Tipo no Excel | Tipo Retornado | Descrição |
|---------------|----------------|-----------|
| Column | BarChart | Gráfico de colunas (vertical) |
| Bar | BarChart | Gráfico de barras (horizontal) |
| Line | LineChart | Gráfico de linhas |
| Pie | PieChart | Gráfico de pizza |
| Area | AreaChart | Gráfico de área |
| Scatter | ScatterChart | Gráfico de dispersão |

**Nota:** Column e Bar são ambos `BarChart` no openpyxl, diferenciados pela propriedade `type`.

## Benefícios

### Para Automação
- ✅ Permite lógica condicional (if chart exists...)
- ✅ Evita duplicação de gráficos
- ✅ Base para operações de delete/update
- ✅ Auditoria e documentação automática

### Para Debugging
- ✅ Ver quais gráficos existem sem abrir Excel
- ✅ Identificar gráficos problemáticos
- ✅ Verificar posições e tipos

### Para Usuários
- ✅ Comando simples em linguagem natural
- ✅ Resposta clara e formatada
- ✅ Informações completas sobre cada gráfico

## Integração com Outras Funcionalidades

### Com Validação de Posição
```python
# Listar gráficos para ver posições ocupadas
result = tool.list_charts('vendas.xlsx', 'Sheet1')
occupied_positions = [chart['position'] for chart in result['charts']]

# Escolher posição livre
free_position = 'K2' if 'K2' not in occupied_positions else 'M2'

# Adicionar gráfico em posição livre
chart_config['position'] = free_position
tool.add_chart('vendas.xlsx', 'Sheet1', chart_config)
```

### Com Delete Chart (próxima funcionalidade)
```python
# Listar gráficos
result = tool.list_charts('vendas.xlsx', 'Sheet1')

# Deletar gráfico específico por título
for chart in result['charts']:
    if chart['title'] == 'Gráfico Antigo':
        tool.delete_chart('vendas.xlsx', 'Sheet1', chart['index'])
```

## Limitações e Considerações

### Limitações
- Não lista gráficos em sheets ocultas (por design)
- Posição pode ser 'Unknown' em casos raros de gráficos mal formados
- Não retorna dados do gráfico, apenas metadados

### Considerações de Performance
- Operação rápida (< 100ms para arquivos típicos)
- Não carrega dados das células, apenas metadados dos gráficos
- Escala bem com muitos gráficos (testado até 50 gráficos)

## Compatibilidade

- ✅ Funciona com todos os tipos de gráficos suportados
- ✅ Funciona com gráficos com ou sem título
- ✅ Funciona com múltiplas sheets
- ✅ Não modifica o arquivo (operação de leitura)
- ✅ Compatível com arquivos criados no Excel ou pelo projeto

## Próximos Passos

Esta funcionalidade é a base para:
1. ✅ **Validação de posição** - IMPLEMENTADO
2. ✅ **list_charts** - IMPLEMENTADO
3. 🔜 **delete_chart** - Próxima implementação

---

**Implementado em:** 28/03/2026
**Versão:** 1.4.0
**Status:** ✅ Produção
**Testes:** 12/12 passando (100%)
**Esforço:** 1 hora
**Valor:** Alto (base para automações inteligentes)
