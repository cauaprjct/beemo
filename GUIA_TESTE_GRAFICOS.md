# 🧪 Guia de Teste - Funcionalidades de Gráficos

Este guia mostra como testar as três novas funcionalidades de gerenciamento de gráficos implementadas.

## 📋 Pré-requisitos

1. Feche o Excel se estiver com o arquivo `vendas_2025.xlsx` aberto (Windows file locking)
2. Certifique-se de que o arquivo existe em: `C:\Users\ngb\Documents\GeminiOfficeFiles\vendas_2025.xlsx`
3. Inicie o Streamlit: `streamlit run app.py`

## 🎯 Cenários de Teste

### Cenário 1: Criar Gráficos com Validação de Posição

#### Teste 1.1: Criar primeiro gráfico (deve funcionar)
```
Crie um gráfico de colunas no arquivo vendas_2025.xlsx na posição H2 
mostrando vendas por produto. Use as colunas C (Produto) e F (Total).
```

**Resultado esperado:**
- ✅ Gráfico criado com sucesso na posição H2
- ✅ Mensagem de sucesso

#### Teste 1.2: Tentar criar gráfico na mesma posição (deve falhar)
```
Crie um gráfico de pizza no arquivo vendas_2025.xlsx na posição H2 
mostrando distribuição de vendas.
```

**Resultado esperado:**
- ❌ Erro: "Position 'H2' is already occupied by a chart..."
- ❌ Mensagem indica o título do gráfico existente
- ❌ Sugere usar posição diferente ou replace_existing=True

#### Teste 1.3: Criar gráfico em posição diferente (deve funcionar)
```
Crie um gráfico de pizza no arquivo vendas_2025.xlsx na posição K2 
mostrando distribuição de vendas por produto. Use as colunas C e F.
```

**Resultado esperado:**
- ✅ Gráfico criado com sucesso na posição K2
- ✅ Agora existem 2 gráficos no arquivo

---

### Cenário 2: Listar Gráficos Existentes

#### Teste 2.1: Listar todos os gráficos
```
Liste todos os gráficos do arquivo vendas_2025.xlsx
```

**Resultado esperado:**
```
Encontrados 2 gráficos:

Sheet1:
  1. [Título do gráfico 1] (BarChart) na posição H2
  2. [Título do gráfico 2] (PieChart) na posição K2
```

#### Teste 2.2: Listar gráficos de sheet específica
```
Quais gráficos existem na Sheet1 do arquivo vendas_2025.xlsx?
```

**Resultado esperado:**
- ✅ Lista os gráficos da Sheet1
- ✅ Mostra título, tipo, posição e índice de cada gráfico

#### Teste 2.3: Verificar antes de adicionar
```
Liste os gráficos do arquivo vendas_2025.xlsx. Se já existir um gráfico 
na posição M2, me avise. Caso contrário, crie um gráfico de linhas nessa posição.
```

**Resultado esperado:**
- ✅ Lista os gráficos primeiro
- ✅ Verifica se M2 está livre
- ✅ Cria o gráfico se a posição estiver livre

---

### Cenário 3: Deletar Gráficos

#### Teste 3.1: Deletar gráfico por índice
```
Delete o primeiro gráfico (índice 0) da Sheet1 do arquivo vendas_2025.xlsx
```

**Resultado esperado:**
- ✅ Gráfico deletado com sucesso
- ✅ Mensagem confirma qual gráfico foi removido
- ✅ Número total de gráficos diminui em 1

#### Teste 3.2: Deletar gráfico por título
```
Remova o gráfico de pizza do arquivo vendas_2025.xlsx
```

**Resultado esperado:**
- ✅ Gráfico deletado com sucesso
- ✅ Busca pelo título e remove o gráfico correto

#### Teste 3.3: Tentar deletar gráfico inexistente (deve falhar)
```
Delete o gráfico chamado "Gráfico Inexistente" do arquivo vendas_2025.xlsx
```

**Resultado esperado:**
- ❌ Erro: "Chart with title 'Gráfico Inexistente' not found..."
- ❌ Mensagem lista os gráficos disponíveis
- ❌ Ajuda o usuário a identificar o nome correto

---

### Cenário 4: Workflow Completo (Listar → Deletar → Criar)

#### Teste 4.1: Substituir gráfico antigo por novo
```
Liste os gráficos do arquivo vendas_2025.xlsx, delete todos os gráficos 
existentes, e depois crie um novo gráfico de barras na posição H2 mostrando 
vendas por produto.
```

**Resultado esperado:**
- ✅ Lista os gráficos existentes
- ✅ Deleta todos os gráficos
- ✅ Cria novo gráfico na posição H2
- ✅ Workflow completo executado com sucesso

#### Teste 4.2: Limpeza e recriação
```
Delete todos os gráficos do arquivo vendas_2025.xlsx e crie dois novos:
1. Gráfico de colunas na posição H2 com vendas por produto
2. Gráfico de linhas na posição K2 com evolução de vendas por data
```

**Resultado esperado:**
- ✅ Todos os gráficos antigos removidos
- ✅ Dois novos gráficos criados nas posições especificadas
- ✅ Operação em lote executada com sucesso

---

### Cenário 5: Testar Undo/Redo

#### Teste 5.1: Deletar e desfazer
```
1. Liste os gráficos do arquivo vendas_2025.xlsx
2. Delete o primeiro gráfico
3. Use o botão "Undo" na sidebar do Streamlit
```

**Resultado esperado:**
- ✅ Gráfico é deletado
- ✅ Undo restaura o gráfico deletado
- ✅ Gráfico volta a aparecer na listagem

#### Teste 5.2: Criar, deletar e refazer
```
1. Crie um gráfico de teste
2. Delete o gráfico
3. Use "Undo" para restaurar
4. Use "Redo" para deletar novamente
```

**Resultado esperado:**
- ✅ Undo/Redo funciona corretamente com gráficos
- ✅ Estado do arquivo é restaurado corretamente

---

## 🔍 Verificação Manual no Excel

Após cada teste, você pode abrir o arquivo no Excel para verificar visualmente:

1. Abra `C:\Users\ngb\Documents\GeminiOfficeFiles\vendas_2025.xlsx`
2. Verifique se os gráficos estão nas posições corretas
3. Verifique se os gráficos deletados realmente foram removidos
4. Verifique se os títulos e tipos estão corretos

## 🧪 Testes Automatizados

Para executar os testes automatizados:

```bash
# Testar validação de posição
python -m pytest tests/test_chart_position_validation.py -v

# Testar listagem de gráficos
python -m pytest tests/test_list_charts.py -v

# Testar deleção de gráficos
python -m pytest tests/test_delete_chart.py -v

# Executar todos os testes de gráficos
python -m pytest tests/test_excel_chart.py tests/test_chart_position_validation.py tests/test_list_charts.py tests/test_delete_chart.py -v
```

**Resultado esperado:** Todos os 50 testes devem passar (100%)

## 📊 Comandos Rápidos para Teste

### Setup Inicial
```
Crie três gráficos no arquivo vendas_2025.xlsx:
1. Gráfico de colunas na posição H2 com vendas por produto
2. Gráfico de pizza na posição K2 com distribuição de vendas
3. Gráfico de linhas na posição H15 com evolução temporal
```

### Teste Rápido de Listagem
```
Liste todos os gráficos do arquivo vendas_2025.xlsx
```

### Teste Rápido de Deleção
```
Delete o segundo gráfico do arquivo vendas_2025.xlsx
```

### Teste Rápido de Validação
```
Tente criar um gráfico na posição H2 do arquivo vendas_2025.xlsx
```

## ⚠️ Problemas Comuns

### Problema 1: "File is locked"
**Causa:** Excel está aberto com o arquivo
**Solução:** Feche o Excel antes de executar operações

### Problema 2: "Chart not found"
**Causa:** Título do gráfico não corresponde exatamente
**Solução:** Use `list_charts` para ver os títulos exatos (case-sensitive)

### Problema 3: "Position already occupied"
**Causa:** Já existe um gráfico na posição especificada
**Solução:** Use posição diferente ou delete o gráfico existente primeiro

### Problema 4: "Sheet not found"
**Causa:** Nome da sheet está incorreto
**Solução:** Verifique o nome exato da sheet (geralmente "Sheet1")

## ✅ Checklist de Validação

Após os testes, verifique:

- [ ] Gráficos são criados nas posições corretas
- [ ] Validação previne sobreposição de gráficos
- [ ] list_charts retorna informações corretas (título, tipo, posição, índice)
- [ ] delete_chart remove gráficos por índice
- [ ] delete_chart remove gráficos por título
- [ ] Mensagens de erro são claras e informativas
- [ ] Undo/Redo funciona com operações de gráficos
- [ ] Workflow completo (listar → deletar → criar) funciona
- [ ] Testes automatizados passam 100%

## 🎉 Teste de Aceitação Final

Execute este comando para testar todas as funcionalidades de uma vez:

```
No arquivo vendas_2025.xlsx:
1. Liste todos os gráficos existentes
2. Delete todos os gráficos
3. Crie um gráfico de colunas na posição H2 mostrando vendas por produto
4. Crie um gráfico de pizza na posição K2 mostrando distribuição
5. Liste os gráficos novamente para confirmar
6. Tente criar outro gráfico na posição H2 (deve falhar)
```

**Resultado esperado:**
- ✅ Todas as 6 operações executadas corretamente
- ✅ Mensagens claras em cada etapa
- ✅ Validação previne sobreposição no passo 6

---

**Documentação criada em:** 28/03/2026
**Versão:** 1.5.0
**Status:** ✅ Pronto para teste
