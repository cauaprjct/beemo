# Implementação de Ordenação (Sort) para Excel

## Visão Geral

A funcionalidade de ordenação permite ordenar dados em planilhas Excel por qualquer coluna, em ordem crescente ou decrescente, preservando a integridade das linhas.

## Características

### Capacidades
- Ordenar por letra de coluna (A, B, C, etc.) ou número (1, 2, 3, etc.)
- Ordem crescente (asc) ou decrescente (desc)
- Suporte para linha de cabeçalho (header row)
- Ordenar intervalo específico de linhas
- Preserva a integridade das linhas (todas as colunas se movem juntas)
- Funciona com diferentes tipos de dados (texto, números, datas)

### Parâmetros de Configuração

```python
sort_config = {
    'column': 'B',           # Coluna para ordenar (letra ou número) - OBRIGATÓRIO
    'order': 'asc',          # 'asc' ou 'desc' - padrão: 'asc'
    'has_header': True,      # Se a primeira linha é cabeçalho - padrão: True
    'start_row': 2,          # Primeira linha a incluir - padrão: 2 se has_header, senão 1
    'end_row': 50            # Última linha a incluir - padrão: última linha com dados
}
```

## Implementação

### 1. Método `sort_data()` em `src/excel_tool.py`

```python
def sort_data(self, file_path: str, sheet: str, sort_config: Dict[str, Any]) -> None:
    """Ordena dados em uma sheet do Excel.
    
    Args:
        file_path: Caminho para o arquivo Excel
        sheet: Nome da sheet a ordenar
        sort_config: Dicionário com configuração de ordenação
    """
```

O método:
1. Valida o arquivo e a sheet
2. Converte letra de coluna para número (se necessário)
3. Lê todos os dados do intervalo especificado
4. Ordena os dados pela coluna especificada
5. Escreve os dados ordenados de volta na planilha
6. Preserva a integridade das linhas (todas as colunas se movem juntas)

### 2. Validação em `src/response_parser.py`

Adicionado validação para `sort_config`:
- Verifica se 'column' está presente
- Valida que 'order' é 'asc' ou 'desc'

### 3. Integração em `src/agent.py`

```python
elif operation == 'sort':
    sheet = parameters.get('sheet')
    sort_config = parameters.get('sort_config', {})
    self.excel_tool.sort_data(target_file, sheet, sort_config)
```

### 4. Segurança em `src/security_validator.py`

Adicionado 'sort' à lista `ALLOWED_OPERATIONS`.

### 5. Documentação em `src/prompt_templates.py`

Adicionado:
- "Sort data" na lista de capacidades do Excel
- Seção completa "SORT OPERATIONS" com 3 exemplos
- Dicas de uso para ordenação

## Exemplos de Uso

### Exemplo 1: Ordenar por coluna (crescente)
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "sort",
            "target_file": "vendas_2025.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "sort_config": {
                    "column": "B",
                    "order": "asc",
                    "has_header": true
                }
            }
        }
    ],
    "explanation": "Ordenando dados pela coluna B (Data) em ordem crescente"
}
```

### Exemplo 2: Ordenar por valor (decrescente)
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "sort",
            "target_file": "vendas_2025.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "sort_config": {
                    "column": "F",
                    "order": "desc",
                    "has_header": true
                }
            }
        }
    ],
    "explanation": "Ordenando pela coluna F (Total) em ordem decrescente para mostrar maiores valores primeiro"
}
```

### Exemplo 3: Ordenar intervalo específico
```json
{
    "actions": [
        {
            "tool": "excel",
            "operation": "sort",
            "target_file": "vendas_2025.xlsx",
            "parameters": {
                "sheet": "Sheet1",
                "sort_config": {
                    "column": 3,
                    "start_row": 2,
                    "end_row": 50,
                    "order": "asc",
                    "has_header": true
                }
            }
        }
    ],
    "explanation": "Ordenando linhas 2-50 pela coluna 3 (Produto) em ordem crescente"
}
```

## Comandos em Linguagem Natural

O agente entende comandos como:
- "Ordene o arquivo vendas_2025.xlsx pela coluna B em ordem crescente"
- "Organize os dados por data, do mais antigo para o mais recente"
- "Classifique pela coluna Total em ordem decrescente"
- "Ordene alfabeticamente pela coluna Nome"

## Testes

Criado arquivo `tests/test_excel_sort.py` com 16 testes cobrindo:

### Testes de Funcionalidade
1. ✅ Ordenar por letra de coluna (crescente)
2. ✅ Ordenar por letra de coluna (decrescente)
3. ✅ Ordenar por número de coluna
4. ✅ Ordenar com linha de cabeçalho
5. ✅ Ordenar sem linha de cabeçalho
6. ✅ Ordenar intervalo específico de linhas
7. ✅ Ordenar por coluna de data
8. ✅ Ordenar por coluna de texto
9. ✅ Preservar integridade das linhas

### Testes de Validação
10. ✅ Erro para letra de coluna inválida
11. ✅ Erro para tipo de coluna inválido
12. ✅ Erro para ordem inválida
13. ✅ Erro para sheet inexistente
14. ✅ Erro quando coluna está ausente
15. ✅ Erro para arquivo não encontrado
16. ✅ Erro para intervalo de linhas inválido

**Resultado**: 16/16 testes passando (100%)

## Tratamento de Erros

A implementação trata os seguintes erros:
- `FileNotFoundError`: Arquivo não existe
- `ValueError`: Parâmetros inválidos (coluna, ordem, sheet, intervalo)
- `CorruptedFileError`: Arquivo Excel corrompido
- `IOError`: Falha na operação de ordenação

## Considerações Técnicas

### Preservação de Integridade
- Todas as colunas de uma linha se movem juntas durante a ordenação
- Exemplo: Se ordenar por "Nome", os valores de "Data", "Valor" e "Categoria" da mesma linha permanecem juntos

### Tratamento de Valores Nulos
- Valores `None` são colocados no final da ordenação (tanto em asc quanto desc)
- Usa tupla `(x[1][sort_col_idx - 1] is None, x[1][sort_col_idx - 1])` como chave de ordenação

### Performance
- Lê todos os dados do intervalo na memória
- Ordena usando `sorted()` do Python (Timsort - O(n log n))
- Escreve os dados ordenados de volta na planilha

## Arquivos Modificados

1. `src/excel_tool.py` - Adicionado método `sort_data()` (~130 linhas)
2. `src/response_parser.py` - Adicionada validação de `sort_config`
3. `src/agent.py` - Integração da operação sort
4. `src/security_validator.py` - Adicionado 'sort' a `ALLOWED_OPERATIONS`
5. `src/prompt_templates.py` - Documentação e exemplos
6. `tests/test_excel_sort.py` - Suite completa de testes (16 testes)

## Status

✅ **IMPLEMENTAÇÃO COMPLETA**
- Método implementado
- Validação adicionada
- Integração no agent
- Segurança configurada
- Documentação no prompt
- Testes criados e passando (16/16)
- Documentação criada

## Próximos Passos

Para testar manualmente:
```bash
streamlit run app.py
```

Comando de teste:
```
Ordene o arquivo vendas_2025.xlsx pela coluna B (Data) em ordem crescente
```

Lembre-se de fechar o Excel antes de executar modificações (Windows file locking).
