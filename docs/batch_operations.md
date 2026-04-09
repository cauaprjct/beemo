# Operações em Lote

## Visão Geral

O Gemini Office Agent suporta execução de múltiplas operações similares em uma única solicitação, permitindo que usuários realizem tarefas repetitivas de forma eficiente.

## Como Funciona

### Detecção Automática

O sistema detecta automaticamente quando o usuário solicita múltiplas operações:

```
Usuário: "Atualizar células A1, B1, C1 com valores 10, 20, 30"

Sistema detecta: 3 operações de atualização
↓
Gemini retorna: 3 ações separadas
↓
Agent executa: Cada ação sequencialmente
↓
UI mostra: Progresso e resultados individuais
```

### Execução Sequencial

Cada operação é executada em ordem:
1. Validação de segurança
2. Execução da operação
3. Registro do resultado (sucesso/falha)
4. Continua para próxima operação

**Importante:** O sistema NÃO para na primeira falha. Todas as operações são tentadas.

## Exemplos Práticos

### Excel - Atualizar Múltiplas Células

```
Prompt: "Atualizar células A1, B1, C1 da planilha vendas.xlsx com valores 100, 200, 300"

Resultado:
✅ Ação 1: Atualizar A1 → Sucesso
✅ Ação 2: Atualizar B1 → Sucesso  
✅ Ação 3: Atualizar C1 → Sucesso

Resumo: 3 de 3 operações bem-sucedidas
```

### Excel - Criar Múltiplas Planilhas

```
Prompt: "Criar planilhas Excel para Janeiro, Fevereiro e Março com dados de vendas"

Resultado:
✅ Ação 1: Criar vendas_janeiro.xlsx → Sucesso
✅ Ação 2: Criar vendas_fevereiro.xlsx → Sucesso
✅ Ação 3: Criar vendas_marco.xlsx → Sucesso

Resumo: 3 de 3 operações bem-sucedidas
```

### Word - Adicionar Conteúdo em Múltiplos Documentos

```
Prompt: "Adicionar parágrafo de disclaimer em doc1.docx, doc2.docx e doc3.docx"

Resultado:
✅ Ação 1: Adicionar parágrafo em doc1.docx → Sucesso
❌ Ação 2: Adicionar parágrafo em doc2.docx → Falha (arquivo não encontrado)
✅ Ação 3: Adicionar parágrafo em doc3.docx → Sucesso

Resumo: 2 de 3 operações bem-sucedidas (1 falhou)
```

### PowerPoint - Criar Múltiplas Apresentações

```
Prompt: "Criar apresentações para os departamentos: Vendas, Marketing e RH"

Resultado:
✅ Ação 1: Criar vendas.pptx → Sucesso
✅ Ação 2: Criar marketing.pptx → Sucesso
✅ Ação 3: Criar rh.pptx → Sucesso

Resumo: 3 de 3 operações bem-sucedidas
```

## Tratamento de Erros

### Erro Parcial

Quando algumas operações falham:

```
Prompt: "Atualizar células A1, A2, A3 com valores 10, 20, 30"

Resultado:
✅ Ação 1: Atualizar A1 → Sucesso
✅ Ação 2: Atualizar A2 → Sucesso
❌ Ação 3: Atualizar A3 → Falha (célula protegida)

Status: ⚠️ Sucesso parcial
Resumo: 2 de 3 operações bem-sucedidas (1 falhou)
```

O sistema:
- Marca a operação como sucesso parcial
- Mostra quais operações falharam
- Permite reverter operações bem-sucedidas com undo

### Erro Total

Quando todas as operações falham:

```
Prompt: "Atualizar células em arquivo_inexistente.xlsx"

Resultado:
❌ Ação 1: Atualizar A1 → Falha (arquivo não encontrado)
❌ Ação 2: Atualizar A2 → Falha (arquivo não encontrado)
❌ Ação 3: Atualizar A3 → Falha (arquivo não encontrado)

Status: ❌ Falha
Resumo: 0 de 3 operações bem-sucedidas (todas falharam)
```

## Interface do Usuário

### Durante Execução

A interface mostra progresso em tempo real:

```
📂 Descobrindo e lendo arquivos...
🔎 Filtrando arquivos relevantes...
📖 Lendo vendas.xlsx...
🧩 Construindo contexto...
🤖 Consultando Gemini API...
⚙️ Executando ações...
```

### Após Conclusão

Resultados detalhados são exibidos:

```
✅ 3 de 3 operações bem-sucedidas

Resultados Detalhados:
  ✅ Ação 1: update em vendas.xlsx
     Ferramenta: excel
     Operação: update
     Arquivo: vendas.xlsx
     Status: Sucesso
  
  ✅ Ação 2: update em vendas.xlsx
     Ferramenta: excel
     Operação: update
     Arquivo: vendas.xlsx
     Status: Sucesso
  
  ✅ Ação 3: update em vendas.xlsx
     Ferramenta: excel
     Operação: update
     Arquivo: vendas.xlsx
     Status: Sucesso
```

### Histórico

Operações em lote são marcadas com 📦:

```
📜 Histórico
  📦 #3: Atualizar células A1, B1, C1... (3/3)
  #2: Criar planilha vendas.xlsx
  #1: Ler documento relatório.docx
```

## Estrutura de Dados

### ActionResult

Cada ação individual tem seu resultado:

```python
@dataclass
class ActionResult:
    action_index: int          # Índice da ação (0, 1, 2...)
    tool: str                  # excel, word, powerpoint
    operation: str             # create, update, add, etc.
    target_file: str           # Arquivo alvo
    success: bool              # True se sucedeu
    error: Optional[str]       # Mensagem de erro se falhou
```

### BatchResult

Resumo de todas as ações:

```python
@dataclass
class BatchResult:
    total_actions: int         # Total de ações
    successful_actions: int    # Ações bem-sucedidas
    failed_actions: int        # Ações que falharam
    action_results: List[ActionResult]  # Resultados individuais
    overall_success: bool      # True se pelo menos uma sucedeu
```

## Limitações

1. **Execução Sequencial**: Operações são executadas uma por vez, não em paralelo
   - Futuro: implementar execução paralela para operações independentes

2. **Sem Rollback Automático**: Operações bem-sucedidas não são revertidas se outras falharem
   - Use undo manual para reverter operações individuais

3. **Limite de Operações**: Recomendado máximo de 20 operações por lote
   - Operações muito grandes podem exceder timeout

4. **Dependências**: Operações são independentes
   - Não há suporte para operações que dependem do resultado de outras
   - Futuro: implementar pipeline de operações dependentes

## Melhores Práticas

### 1. Agrupe Operações Similares

✅ Bom:
```
"Atualizar células A1, A2, A3 com valores 10, 20, 30"
```

❌ Evite:
```
"Atualizar A1 com 10"
"Atualizar A2 com 20"  
"Atualizar A3 com 30"
(3 solicitações separadas)
```

### 2. Seja Específico

✅ Bom:
```
"Criar planilhas vendas_jan.xlsx, vendas_fev.xlsx, vendas_mar.xlsx"
```

❌ Vago:
```
"Criar algumas planilhas de vendas"
```

### 3. Verifique Arquivos Antes

Para operações de atualização, certifique-se de que os arquivos existem:

```
1. "Listar arquivos disponíveis"
2. "Atualizar células em vendas.xlsx, estoque.xlsx"
```

### 4. Use Undo para Correções

Se algumas operações falharem, use undo nas bem-sucedidas:

```
Resultado: 2 de 3 operações bem-sucedidas
↓
Corrigir problema
↓
Desfazer operações bem-sucedidas
↓
Tentar novamente com todas
```

## Casos de Uso

### 1. Atualização de Dados em Massa

```
"Atualizar preços nas células B2:B10 da planilha produtos.xlsx com valores 
10.50, 12.00, 15.75, 20.00, 8.50, 11.25, 14.00, 16.50, 9.75"
```

### 2. Criação de Documentos Padronizados

```
"Criar documentos Word para cada cliente: cliente1.docx, cliente2.docx, 
cliente3.docx com template de proposta comercial"
```

### 3. Geração de Relatórios Mensais

```
"Criar apresentações PowerPoint para Janeiro, Fevereiro e Março com 
dados de performance mensal"
```

### 4. Adição de Conteúdo Padrão

```
"Adicionar rodapé com informações de copyright em todos os documentos: 
doc1.docx, doc2.docx, doc3.docx, doc4.docx"
```

### 5. Organização de Dados

```
"Criar abas na planilha dados.xlsx para cada departamento: Vendas, 
Marketing, RH, TI, Financeiro"
```

## Testes

```bash
# Testes de operações em lote
pytest tests/test_batch_operations.py -v

# Teste específico
pytest tests/test_batch_operations.py::test_batch_with_partial_failure -v
```

## Conclusão

Operações em lote tornam o Gemini Office Agent muito mais eficiente para tarefas repetitivas, economizando tempo e esforço do usuário enquanto fornece feedback detalhado de cada operação.
