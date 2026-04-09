# 🚀 Como Rodar o Projeto - Guia Rápido

## ✅ Configuração Concluída!

Seu projeto já está configurado e pronto para rodar:

### 📁 Configurações Atuais

- **API Key:** ✅ Configurada
- **Pasta de Arquivos:** `C:\Users\ngb\Documents\GeminiOfficeFiles`
- **Modelo:** gemini-2.5-flash-lite (gratuito)
- **Cache:** Habilitado
- **Versionamento:** Habilitado (10 versões por arquivo)

### 🎯 Como Rodar

1. **Abra o terminal** no diretório do projeto:
   ```
   cd C:\Users\ngb\Desktop\code
   ```

2. **Execute o Streamlit:**
   ```
   streamlit run app.py
   ```

3. **Acesse no navegador:**
   ```
   http://localhost:8501
   ```

### 📂 Sua Pasta de Arquivos

Todos os arquivos serão salvos em:
```
C:\Users\ngb\Documents\GeminiOfficeFiles
```

Você pode:
- Colocar arquivos existentes lá para manipulá-los
- Deixar vazio e pedir ao agente para criar novos arquivos
- Abrir a pasta no Explorer: `explorer C:\Users\ngb\Documents\GeminiOfficeFiles`

### 🧪 Testes Rápidos

Depois que a interface abrir, teste estes comandos:

**1. Criar Excel:**
```
Crie uma planilha vendas.xlsx com colunas Produto, Quantidade e Preço, e adicione 3 linhas de exemplo
```

**2. Criar PDF:**
```
Crie um PDF relatorio.pdf com título "Relatório Mensal", uma seção "Resumo" e uma tabela com dados de vendas
```

**3. Criar Word:**
```
Crie um documento memo.docx com título "Memorando Interno", data de hoje e um parágrafo de introdução
```

**4. Criar PowerPoint:**
```
Crie uma apresentação vendas.pptx com slide de título "Vendas Q1 2025" e 2 slides de conteúdo
```

### 🔧 Comandos Úteis

**Parar o servidor:**
- Pressione `Ctrl + C` no terminal

**Abrir pasta de arquivos:**
```bash
explorer C:\Users\ngb\Documents\GeminiOfficeFiles
```

**Ver logs:**
```bash
type logs\agent.log
```

**Rodar testes:**
```bash
pytest tests/test_pdf_tool.py -v
```

### 💡 Dicas

1. **Primeira vez:** Clique em "🔄 Atualizar Lista" na sidebar para descobrir arquivos
2. **Seleção:** Você pode selecionar arquivos específicos ou deixar o agente descobrir automaticamente
3. **Histórico:** Veja todas as operações no painel "📜 Histórico"
4. **Undo/Redo:** Use os botões na sidebar para desfazer/refazer operações
5. **Cache:** Comandos repetidos são instantâneos (cache ativo)

### ⚠️ Solução de Problemas

**Erro de módulo não encontrado:**
```bash
pip install -r requirements.txt
```

**Porta 8501 ocupada:**
```bash
streamlit run app.py --server.port 8502
```

**Erro de API Key:**
- Verifique se a chave está correta no arquivo `.env`
- Obtenha uma nova em: https://aistudio.google.com/apikey

### 📚 Exemplos Avançados

**Excel com formatação:**
```
Crie uma planilha budget.xlsx com receitas e despesas, formate o cabeçalho com negrito e fundo azul, e adicione uma fórmula de soma no final
```

**PDF com merge:**
```
Junte os arquivos parte1.pdf e parte2.pdf em um único documento completo.pdf
```

**Word estruturado:**
```
Crie um relatório estruturado relatorio.docx com título, 3 seções (Introdução, Análise, Conclusão), uma tabela de dados e uma lista de próximos passos
```

**PowerPoint com tabela:**
```
Crie uma apresentação resultados.pptx com slide de título, adicione uma tabela no slide 2 com dados de vendas por região
```

---

## 🎉 Pronto para Usar!

Execute `streamlit run app.py` e comece a manipular seus arquivos Office e PDF através de comandos em linguagem natural!

**Divirta-se!** 🚀
