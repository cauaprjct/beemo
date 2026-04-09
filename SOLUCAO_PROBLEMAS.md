# Solução dos Problemas Identificados

## Problema 1: Permission Denied ao Deletar Gráficos

### Sintoma
```
PermissionError: [Errno 13] Permission denied: 'C:\\Users\\ngb\\Documents\\GeminiOfficeFiles\\vendas_2025.xlsx'
```

### Causa
O arquivo Excel estava aberto no Microsoft Excel quando o agente tentou modificá-lo. O Windows bloqueia arquivos abertos para escrita, impedindo que outros programas os modifiquem.

### Solução
Feche o arquivo Excel antes de executar operações de modificação através do agente.

### Melhoria Futura Sugerida
Adicionar detecção de arquivo aberto e mensagem mais clara para o usuário:

```python
# Em excel_tool.py, antes de salvar:
try:
    workbook.save(file_path)
except PermissionError:
    raise IOError(
        f"❌ Não foi possível salvar o arquivo '{file_path}'. "
        f"O arquivo pode estar aberto no Excel. "
        f"Por favor, feche o arquivo e tente novamente."
    )
```

---

## Problema 2: Gemini Não Executava Múltiplas Ações

### Sintoma
Quando o usuário pedia "analise o conteúdo e crie um gráfico de pizza", o Gemini executava apenas a operação `read`, não criando o gráfico.

### Causa
O prompt do sistema não tinha instruções EXPLÍCITAS sobre quando executar múltiplas ações em uma única resposta. O modelo estava sendo conservador e executando apenas a primeira etapa.

### Solução Implementada

#### 1. Adicionada instrução crítica no prompt (src/prompt_templates.py)

```
CRITICAL INSTRUCTION - MULTIPLE ACTIONS:
When a user request requires multiple steps (e.g., "analyze the content and create a chart"), 
you MUST include ALL necessary actions in a SINGLE response.
DO NOT return only a "read" action and wait for another request. 
Complete the entire task in one response.

Examples:
- "Analyze data and create a chart" → Return BOTH read AND add_chart actions
- "Read the file and format the header" → Return BOTH read AND format actions
- "Create a file and add a chart" → Return BOTH create AND add_chart actions

The ONLY exception is when you genuinely need to see the data first to determine 
what to do next (e.g., user says "do something appropriate based on the data").
```

#### 2. Adicionado exemplo específico (Example 10)

```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "read",
            "target_file": "data/sales.xlsx",
            "parameters": {}
        },
        {
            "tool": "excel",
            "operation": "add_chart",
            "target_file": "data/sales.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "chart_config": {
                    "type": "pie",
                    "title": "Sales Distribution",
                    "categories": "A2:A10",
                    "values": "B2:B10",
                    "position": "H2"
                }
            }
        }
    ],
    "explanation": "Reading sales data and creating a pie chart based on the content"
}
```

---

## Problema 3: Gráfico de Pizza Sem Dados e Mal Posicionado

### Sintoma
O Gemini criava o gráfico, mas:
- O gráfico aparecia sem fatias visíveis (apenas legendas)
- O gráfico era posicionado sobre os dados existentes (E2), causando sobreposição

### Causa
1. **Dados invisíveis**: O Gemini estava usando `data_range` em vez de `categories` + `values` separados, ou passando ranges incorretos
2. **Posição ruim**: O Gemini escolhia posições que sobrepunham dados (E2 está no meio dos dados)

### Solução Implementada

#### 1. Adicionada instrução CRÍTICA sobre gráficos de pizza

```
CRITICAL - PIE CHART REQUIREMENTS:
For pie charts, you MUST use separate "categories" and "values" ranges (NOT "data_range").
- categories: Range with labels (e.g., product names in "C2:C11")
- values: Range with numbers (e.g., sales values in "F2:F11")
- Position: Choose a cell that doesn't overlap data (e.g., "H2", "K2", "M2")
- The ranges MUST have actual data, not empty cells

Example of CORRECT pie chart:
{
    "chart_config": {
        "type": "pie",
        "title": "Sales by Product",
        "categories": "C2:C11",  // Product names
        "values": "F2:F11",      // Sales totals
        "position": "H2"         // Right side, away from data
    }
}
```

#### 2. Corrigidos todos os exemplos de gráficos de pizza

Antes (incorreto):
```json
"chart_config": {
    "type": "pie",
    "title": "Budget Distribution",
    "data_range": "A1:B5",  // ❌ Não funciona bem para pizza
    "position": "E2"
}
```

Depois (correto):
```json
"chart_config": {
    "type": "pie",
    "title": "Budget Distribution",
    "categories": "A2:A5",  // ✅ Labels separados
    "values": "B2:B5",      // ✅ Valores separados
    "position": "E2"
}
```

#### 3. Adicionado logging detalhado

Agora o sistema loga o JSON completo de cada ação para facilitar debug:
```python
logger.debug(f"Ação {idx + 1}: {json.dumps(action, indent=2, ensure_ascii=False)}")
```

### Como Testar a Correção

1. **Reinicie o Streamlit** (importante para carregar o novo prompt):
   ```bash
   # Pressione Ctrl+C no terminal onde o Streamlit está rodando
   streamlit run app.py
   ```

2. **Teste com comandos de gráfico de pizza**:
   - "analise o conteúdo e crie um gráfico de pizza mostrando vendas por produto"
   - "crie um gráfico de pizza com os top 10 produtos"
   - "adicione um gráfico de pizza na posição K2"

3. **Verifique os logs**: O log deve mostrar:
   ```
   2026-03-28 18:xx:xx - src.agent - DEBUG - Ação 2: {
     "tool": "excel",
     "operation": "add_chart",
     "target_file": "...",
     "parameters": {
       "sheet": "Vendas",
       "chart_config": {
         "type": "pie",
         "categories": "C2:C11",  // ✅ Deve ter categories
         "values": "F2:F11",      // ✅ Deve ter values
         "position": "H2"         // ✅ Posição longe dos dados
       }
     }
   }
   ```

4. **Resultado esperado**:
   - Gráfico com fatias visíveis e coloridas
   - Legendas mostrando os nomes dos produtos
   - Posicionado à direita dos dados (H2, K2, M2, etc.)

---

## Observações Adicionais

### Cache de Respostas
O sistema tem cache habilitado. Se você testar o mesmo prompt várias vezes, pode receber a resposta cacheada (antiga). Para forçar uma nova chamada à API:
- Use um prompt ligeiramente diferente
- Ou desabilite o cache temporariamente em `.env`: `CACHE_ENABLED=false`

### Aviso sobre google.generativeai
O log mostra um FutureWarning sobre o pacote `google.generativeai` estar deprecated. Considere migrar para `google.genai` no futuro:

```python
# Atual (deprecated):
import google.generativeai as genai

# Recomendado:
import google.genai as genai
```

Veja: https://github.com/google-gemini/deprecated-generative-ai-python/blob/main/README.md
