# 📖 Manual do Usuário - Beemo Office Agent

## 🎯 O que é o Beemo?

O Beemo é um assistente inteligente que permite manipular arquivos Office (Excel, Word, PowerPoint e PDF) usando comandos em linguagem natural, sem precisar conhecer fórmulas ou funções complexas.

---

## 🚀 Instalação Rápida (3 Passos)

### 1️⃣ Instalar Python (apenas primeira vez)

Se você ainda não tem Python instalado:

1. Acesse: https://www.python.org/downloads/
2. Baixe a versão mais recente (3.8 ou superior)
3. **IMPORTANTE:** Marque a opção "Add Python to PATH" durante a instalação
4. Clique em "Install Now"

### 2️⃣ Executar o Instalador

1. Clique duas vezes no arquivo **INSTALAR.bat**
2. Aguarde a instalação das dependências (2-5 minutos)
3. Quando solicitado, configure sua API Key do Google Gemini

### 3️⃣ Configurar API Key (gratuita)

1. Acesse: https://aistudio.google.com/apikey
2. Faça login com sua conta Google
3. Clique em "Create API Key"
4. Copie a chave gerada
5. Abra o arquivo **.env** (ou clique em CONFIGURAR.bat)
6. Cole sua chave no lugar de `your_api_key_here`
7. Configure o caminho da pasta onde seus arquivos ficarão (ROOT_PATH)

**Exemplo de configuração:**
```
GEMINI_API_KEY=AIzaSyABC123def456GHI789jkl
ROOT_PATH=C:\Users\SeuNome\Documents\MeusArquivos
```

---

## ▶️ Como Usar

### Iniciar o Beemo

1. Clique duas vezes no arquivo **EXECUTAR.bat**
2. Aguarde o navegador abrir automaticamente
3. Pronto! A interface do Beemo estará disponível

### Parar o Beemo

- Feche a janela do terminal (cmd)
- Ou pressione **Ctrl + C** no terminal

---

## 💬 Comandos Básicos

### 📊 Excel

**Criar planilha:**
```
Crie uma planilha vendas.xlsx com colunas Produto, Quantidade e Preço
```

**Adicionar dados:**
```
Adicione 5 linhas de exemplo na planilha vendas.xlsx
```

**Criar gráfico:**
```
Adicione um gráfico de colunas mostrando vendas por produto
```

**Ordenar dados:**
```
Ordene a planilha pela coluna Preço em ordem decrescente
```

**Aplicar fórmulas:**
```
Adicione uma coluna Total com a fórmula Quantidade * Preço
```

### 📝 Word

**Criar documento:**
```
Crie um documento relatorio.docx com título "Relatório Mensal"
```

**Adicionar conteúdo:**
```
Adicione uma seção "Introdução" com 2 parágrafos
```

**Inserir tabela:**
```
Adicione uma tabela 3x4 com dados de vendas
```

### 📑 PowerPoint

**Criar apresentação:**
```
Crie uma apresentação vendas.pptx com slide de título "Vendas 2025"
```

**Adicionar slides:**
```
Adicione 3 slides de conteúdo com bullet points
```

**Inserir tabela:**
```
Adicione uma tabela no slide 2 com dados de vendas por região
```

### 📕 PDF

**Criar PDF:**
```
Crie um PDF relatorio.pdf com título "Relatório Anual" e 3 seções
```

**Juntar PDFs:**
```
Junte os arquivos parte1.pdf e parte2.pdf em completo.pdf
```

**Extrair páginas:**
```
Extraia as páginas 1 a 5 do arquivo documento.pdf
```

---

## 🎨 Interface

### Painel Esquerdo (Sidebar)

**📁 Arquivos Disponíveis**
- Lista todos os arquivos Office na pasta configurada
- Marque os arquivos que deseja manipular
- Use "Selecionar Todos" para trabalhar com múltiplos arquivos

**⏱️ Histórico de Versões**
- Mostra todas as operações realizadas em cada arquivo
- Botão **↩️ Desfazer**: volta a última operação
- Botão **↪️ Refazer**: refaz operação desfeita

**⚡ Cache de Respostas**
- Comandos repetidos são instantâneos
- Economiza tempo e chamadas à API
- Botão **🗑️ Limpar Cache** para resetar

**❓ Ajuda**
- Exemplos de comandos
- Dicas de uso
- Funcionalidades disponíveis

### Área Principal

**💬 Nova Solicitação**
- Digite seu comando em linguagem natural
- Seja específico sobre o que deseja fazer
- Clique em **🚀 Enviar** para executar

**📜 Histórico de Conversas**
- Veja todos os comandos executados
- Status de sucesso (✅) ou erro (❌)
- Detalhes de cada operação

---

## 💡 Dicas e Truques

### ✅ Boas Práticas

1. **Seja específico:** "Adicione um gráfico de pizza mostrando vendas por região" é melhor que "adicione gráfico"

2. **Um comando por vez:** Execute operações complexas em etapas
   - ❌ "Crie planilha, adicione dados, formate e crie gráfico"
   - ✅ "Crie uma planilha vendas.xlsx" → depois → "Adicione dados de exemplo"

3. **Use nomes claros:** Nomeie arquivos de forma descritiva
   - ✅ vendas_janeiro_2025.xlsx
   - ❌ planilha1.xlsx

4. **Aproveite o histórico:** Use Desfazer se algo não ficou como esperado

5. **Cache inteligente:** Comandos idênticos são instantâneos

### ⚠️ Evite

- Comandos muito vagos: "faça algo com a planilha"
- Múltiplas ações em um comando: "crie, formate e envie"
- Arquivos sem extensão: use sempre .xlsx, .docx, .pptx, .pdf

---

## 🔧 Solução de Problemas

### ❌ "Python não encontrado"

**Solução:**
1. Instale o Python: https://www.python.org/downloads/
2. **IMPORTANTE:** Marque "Add Python to PATH"
3. Reinicie o computador
4. Execute INSTALAR.bat novamente

### ❌ "API Key não configurada"

**Solução:**
1. Obtenha sua chave em: https://aistudio.google.com/apikey
2. Execute CONFIGURAR.bat
3. Cole sua chave no lugar de `your_api_key_here`
4. Salve o arquivo

### ❌ "Erro ao instalar dependências"

**Solução:**
1. Abra o terminal (cmd) como Administrador
2. Navegue até a pasta do Beemo: `cd C:\caminho\para\beemo`
3. Execute: `pip install -r requirements.txt`
4. Se persistir, atualize o pip: `python -m pip install --upgrade pip`

### ❌ "Porta 8501 ocupada"

**Solução:**
1. Feche outras instâncias do Beemo
2. Ou edite EXECUTAR.bat e adicione ao final da linha do streamlit:
   ```
   streamlit run app.py --server.port 8502
   ```

### ❌ "Arquivo não encontrado"

**Solução:**
1. Verifique se o arquivo está na pasta configurada (ROOT_PATH)
2. Clique em "🔄 Atualizar Lista" na sidebar
3. Certifique-se de que o arquivo tem extensão correta (.xlsx, .docx, etc.)

### ❌ "Rate limit exceeded"

**Solução:**
- A API gratuita tem limites de uso
- Aguarde alguns minutos antes de tentar novamente
- O Beemo usa fallback automático para outros modelos

---

## 📊 Limites da API Gratuita

O Google Gemini oferece uso gratuito com os seguintes limites:

| Modelo | Requisições por Minuto |
|--------|------------------------|
| gemini-2.5-flash-lite | 15 RPM |
| gemini-2.5-flash | 10 RPM |
| gemini-2.5-pro | 5 RPM |

**O Beemo gerencia isso automaticamente:**
- Usa cache para comandos repetidos
- Alterna entre modelos quando atinge limite
- Aguarda automaticamente se necessário

---

## 🎓 Exemplos Práticos

### Exemplo 1: Relatório de Vendas Completo

```
1. Crie uma planilha vendas_2025.xlsx com colunas Data, Produto, Quantidade, Preço, Total

2. Adicione 10 linhas de dados de exemplo com datas de janeiro de 2025

3. Formate o cabeçalho com negrito e fundo azul

4. Adicione uma fórmula na coluna Total multiplicando Quantidade por Preço

5. Adicione um gráfico de colunas mostrando vendas por produto

6. Ordene os dados pela coluna Total em ordem decrescente
```

### Exemplo 2: Apresentação Profissional

```
1. Crie uma apresentação resultados.pptx com slide de título "Resultados Q1 2025"

2. Adicione um slide com bullet points sobre principais conquistas

3. Adicione um slide com uma tabela 4x3 mostrando vendas por região

4. Adicione um slide de conclusão com próximos passos
```

### Exemplo 3: Documento Estruturado

```
1. Crie um documento proposta.docx com título "Proposta Comercial"

2. Adicione uma seção "Introdução" com 2 parágrafos

3. Adicione uma seção "Escopo" com uma lista de 5 itens

4. Adicione uma tabela 3x4 com cronograma de entregas

5. Adicione uma seção "Investimento" com valores
```

---

## 🆘 Suporte

### Documentação Adicional

- **COMO_RODAR.md**: Guia técnico de instalação
- **ANALISE_FRONTEND_STREAMLIT.md**: Detalhes da interface
- **docs/**: Documentação técnica de cada funcionalidade

### Problemas Comuns

Consulte a seção "Solução de Problemas" acima ou os arquivos de documentação na pasta `docs/`.

### Logs

Se encontrar erros, verifique o arquivo de log:
```
logs\agent.log
```

---

## 📝 Notas Importantes

1. **Backup:** O Beemo mantém 10 versões de cada arquivo automaticamente
2. **Privacidade:** Seus arquivos ficam no seu computador, apenas comandos são enviados à API
3. **Internet:** Necessária para comunicação com a API do Google Gemini
4. **Gratuito:** A API do Gemini é gratuita dentro dos limites de uso

---

## 🎉 Aproveite o Beemo!

Agora você está pronto para manipular seus arquivos Office de forma inteligente e natural!

**Dica Final:** Crie um atalho do EXECUTAR.bat na área de trabalho para acesso rápido!

---

**Versão:** 1.0  
**Última Atualização:** Março 2026
