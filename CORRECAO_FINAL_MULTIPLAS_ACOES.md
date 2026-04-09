# Correção Final - Problema de Múltiplas Ações

## Problema Persistente

Mesmo após todas as correções, o Gemini ainda executava apenas `list_charts` quando o usuário pedia "remova os graficos", sem executar `delete_chart`.

**Log:**
```
2026-03-28 19:34:00 - src.agent - INFO - Resposta parseada: 1 ação(ões) identificada(s)
2026-03-28 19:34:00 - src.agent - INFO - Explicação: Listing all charts in the Excel file to identify them before deletion.
2026-03-28 19:34:00 - src.agent - INFO - Executando ação: excel.list_charts
```

## Causa Raiz

A instrução sobre múltiplas ações era **genérica demais**. Dizia:
- "Complete a tarefa inteira"
- "Não retorne apenas read"

Mas NÃO dizia explicitamente:
- ❌ "NÃO retorne apenas list_charts sem delete_chart"
- ❌ "list_charts DEVE ser seguido de delete_chart"

O Gemini interpretava `list_charts` como "completar a tarefa" porque tecnicamente ele estava "listando os gráficos" como pedido.

## Correção Implementada

### 1. Instrução MUITO Mais Específica ✅

**Antes:**
```
CRITICAL INSTRUCTION - MULTIPLE ACTIONS:
When a user request requires multiple steps, you MUST include ALL necessary actions.
```

**Depois:**
```
CRITICAL INSTRUCTION - MULTIPLE ACTIONS:

IMPORTANT RULES:
1. DO NOT use "list_charts" as a standalone action
2. DO NOT return only discovery actions (read, list_charts) without the actual work
3. Complete the ENTIRE task in ONE response

Examples of CORRECT multi-action responses:
- "Remove charts" → Return list_charts AND delete_chart actions (NOT just list_charts!)

Examples of INCORRECT responses (DO NOT DO THIS):
- ❌ "Remove charts" → Only list_charts (WRONG! Must also delete!)
```

### 2. Exemplo Específico de Remover Todos os Gráficos ✅

Adicionado exemplo completo:

```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "list_charts",
            "target_file": "data/sales.xlsx",
            "parameters": {}
        },
        {
            "tool": "excel",
            "operation": "delete_chart",
            "target_file": "data/sales.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "identifier": 0
            }
        },
        {
            "tool": "excel",
            "operation": "delete_chart",
            "target_file": "data/sales.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "identifier": 0
            }
        }
    ],
    "explanation": "Listing all charts, then deleting first chart twice (index shifts)"
}
```

### 3. Nota Crítica Sobre Índices ✅

Adicionado aviso importante:

```
CRITICAL NOTE: When deleting multiple charts, always use index 0 repeatedly 
because after deleting the first chart, the second chart becomes the new 
first chart (index 0).
```

## Por Que Isso Deve Funcionar

Agora o prompt tem:

1. ✅ **Regra explícita**: "DO NOT use list_charts as standalone"
2. ✅ **Exemplo negativo**: "❌ Only list_charts (WRONG!)"
3. ✅ **Exemplo positivo**: Mostra list_charts + delete_chart + delete_chart
4. ✅ **Explicação técnica**: Por que usar índice 0 repetidamente

O Gemini agora tem ZERO ambiguidade sobre o que fazer quando o usuário pede "remova os gráficos".

## Teste de Validação

### 1. Reiniciar Streamlit

```bash
Ctrl+C
streamlit run app.py
```

### 2. Testar Comando

```
"remova os graficos"
```

**Esperado:**
```
✅ Resposta parseada: 3 ação(ões) identificada(s)
✅ Ação 1: list_charts
✅ Ação 2: delete_chart (identifier: 0)
✅ Ação 3: delete_chart (identifier: 0)
```

**NÃO esperado:**
```
❌ Resposta parseada: 1 ação(ões) identificada(s)
❌ Apenas list_charts
```

### 3. Verificar Resultado

Abra o arquivo Excel e confirme que os gráficos foram removidos.

## Se Ainda Falhar

Se o Gemini AINDA executar apenas `list_charts`:

### Opção 1: Usar Modelo Mais Potente

```env
# Em .env
MODEL_NAME=gemini-2.5-flash
```

O gemini-2.5-flash-lite pode ter limitações de compreensão.

### Opção 2: Simplificar o Comando

Em vez de "remova os graficos", tente:

```
"delete todos os graficos da planilha"
"execute list_charts e depois delete_chart para cada grafico"
```

### Opção 3: Forçar Múltiplas Ações no Código

Modificar `agent.py` para detectar quando apenas `list_charts` é retornado e avisar:

```python
# Após parse da resposta
if len(actions) == 1 and actions[0]['operation'] == 'list_charts':
    if 'remov' in user_prompt.lower() or 'delet' in user_prompt.lower():
        logger.warning("Gemini returned only list_charts for a delete request!")
        # Poderia até rejeitar e pedir nova resposta
```

## Arquivos Modificados

```
src/prompt_templates.py
  ✅ Linha 122: Instrução MUITO mais específica sobre múltiplas ações
  ✅ Linha 1177: Exemplo de remover TODOS os gráficos
  ✅ Linha 1205: Nota crítica sobre índices
```

## Conclusão

O problema era que o Gemini estava sendo **tecnicamente correto** mas **praticamente inútil**.

Ele listava os gráficos (como pedido) mas não os deletava (também como pedido).

Agora, com instruções MUITO EXPLÍCITAS e exemplos NEGATIVOS (o que NÃO fazer), o Gemini deve entender que:

**"remova os graficos" = list_charts + delete_chart + delete_chart**

Não apenas:

**"remova os graficos" = list_charts ✋ (parar aqui)**

Reinicie o Streamlit e teste!
