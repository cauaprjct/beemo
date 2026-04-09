# Problema: Gemini Respondendo em Texto Livre

## Sintoma

O Gemini está respondendo em texto livre em vez de JSON:

```
"I need more information to proceed. You asked to 'remove all charts 
and create a new one,' but you haven't specified:
1. Which file to operate on...
2. Which chart to create..."
```

**Log:**
```
2026-03-28 19:22:31 - src.response_parser - WARNING - JSON parsing failed: 
Expecting value: line 1 column 1 (char 0). Attempting free-text parsing.
```

## Causa

O Gemini está ignorando a instrução de responder em JSON e está:
1. Fazendo perguntas em vez de analisar os dados disponíveis
2. Pedindo informações que já estão no contexto (arquivo selecionado)
3. Respondendo como um chatbot conversacional em vez de um agente de ação

## Por Que Isso Acontece

1. **Instrução não é forte o suficiente** - "Please respond with JSON" é muito educado
2. **Instrução está no meio do prompt** - Pode ser "perdida" entre muito texto
3. **Modelo pode estar confuso** - gemini-2.5-flash-lite pode ter limitações

## Correções Implementadas

### 1. Instrução CRÍTICA no Início ✅

Adicionado logo no início do system prompt:

```
You are an Office Agent assistant...

CRITICAL RULE: You MUST ALWAYS respond with valid JSON. 
NEVER respond with free text or questions.
```

### 2. Instrução EXPLÍCITA no Final ✅

Adicionado no final do prompt contextualizado:

```
================================================================================
CRITICAL - YOUR RESPONSE FORMAT:
================================================================================

You MUST respond with VALID JSON ONLY. NO free text, NO explanations outside JSON.

Your response must be a JSON object with this EXACT structure:
{
    "actions": [...],
    "explanation": "..."
}

DO NOT respond with free text like 'I need more information'.
DO NOT ask questions - analyze the available files and make intelligent decisions.
The file is already provided in the context above - use it!

Respond with JSON now:
```

### 3. Instrução de Análise Inteligente ✅

Reforçado que o Gemini deve:
- Analisar os arquivos disponíveis
- Tomar decisões inteligentes
- NÃO fazer perguntas
- NÃO pedir mais informações

## Teste de Validação

Após reiniciar o Streamlit, testar com:

```
"remova todos os graficos e crie um novo"
```

**Esperado:**
```json
{
  "actions": [
    {
      "tool": "excel",
      "operation": "list_charts",
      "target_file": "vendas_2025.xlsx",
      "parameters": {"sheet": "Vendas"}
    },
    {
      "tool": "excel",
      "operation": "delete_chart",
      "target_file": "vendas_2025.xlsx",
      "parameters": {"sheet": "Vendas", "identifier": 0}
    },
    {
      "tool": "excel",
      "operation": "add_chart",
      "target_file": "vendas_2025.xlsx",
      "parameters": {
        "sheet": "Vendas",
        "chart_config": {
          "type": "pie",
          "categories": "C2:C101",
          "values": "F2:F101",
          "position": "H2"
        }
      }
    }
  ],
  "explanation": "Listing, deleting existing chart, and creating new pie chart"
}
```

**NÃO esperado:**
```
I need more information to proceed...
```

## Se o Problema Persistir

### Opção 1: Desabilitar Cache Temporariamente

O cache pode estar retornando respostas antigas em texto livre.

```env
# Em .env
CACHE_ENABLED=false
```

### Opção 2: Testar com Modelo Diferente

O gemini-2.5-flash-lite pode ter limitações. Testar com:

```env
# Em .env
MODEL_NAME=gemini-2.5-flash
```

### Opção 3: Adicionar Validação Mais Rígida

Modificar `response_parser.py` para REJEITAR respostas em texto livre:

```python
def parse_response(response_text: str) -> Dict[str, Any]:
    try:
        return json.loads(response_text)
    except json.JSONDecodeError:
        # REJEITAR em vez de tentar parse de texto livre
        raise ValueError(
            "Gemini returned invalid response (not JSON). "
            "This indicates a prompt engineering issue. "
            "Response preview: " + response_text[:200]
        )
```

### Opção 4: Usar JSON Mode (se disponível)

Verificar se a API do Gemini suporta "JSON mode" forçado:

```python
# Em gemini_client.py
generation_config = {
    "response_mime_type": "application/json"  # Força JSON
}
```

## Próximos Passos

1. **Reiniciar Streamlit** (OBRIGATÓRIO)
2. **Limpar cache** (opcional mas recomendado)
3. **Testar novamente**
4. **Verificar logs** para confirmar que resposta é JSON

## Arquivos Modificados

```
src/prompt_templates.py
  - Linha 31: Adicionada instrução CRÍTICA no início
  - Linhas 1724-1745: Adicionada instrução EXPLÍCITA no final
```

## Conclusão

O problema não é técnico, é de COMUNICAÇÃO com o modelo.

O Gemini precisa de instruções MUITO CLARAS e REPETIDAS sobre:
1. Sempre responder em JSON
2. Nunca fazer perguntas
3. Analisar e decidir autonomamente

Essas correções devem resolver o problema de respostas em texto livre.
