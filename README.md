# Beemo — Office Agent

Um agente Python que permite manipular arquivos Office (Excel, Word, PowerPoint) e PDF através de comandos em linguagem natural, utilizando a API do Google Gemini (free tier).

O Beemo usa **native function calling** do Gemini — o modelo decide diretamente quais ferramentas acionar, eliminando o pipeline manual de parsing de JSON. Resultado: mais precisão, menos falsos positivos e suporte nativo a conversas mistas (perguntas + ações).

## Requisitos

- Python 3.8 ou superior
- Uma chave de API do Google Gemini ([obtenha aqui](https://aistudio.google.com/apikey))

## Instalação

```bash
git clone <url-do-repositorio>
cd beemo
pip install -r requirements.txt
cp .env.example .env
# Edite .env com sua GEMINI_API_KEY e ROOT_PATH
```

## Executando

```bash
streamlit run app.py
```

Acesse `http://localhost:8501` no navegador.

## Configuração (.env)

```env
GEMINI_API_KEY=sua_chave_aqui           # Obrigatório
ROOT_PATH=/caminho/para/seus/arquivos   # Pasta com seus arquivos (auto-criada)
MODEL_NAME=gemini-2.5-flash-lite        # Modelo principal (gratuito)
# FALLBACK_MODELS=gemini-2.5-flash,gemini-2.5-pro,gemini-3-flash-preview,gemini-3.1-flash-lite-preview,gemini-3.1-flash-live-preview
MAX_VERSIONS=10                         # Versões de backup por arquivo
CACHE_ENABLED=true                      # Cache de respostas da API
CACHE_TTL_HOURS=24                      # Tempo de vida do cache
CACHE_MAX_ENTRIES=100                   # Máximo de entradas no cache
```

**Modelos gratuitos disponíveis:**

| Modelo | RPM | TPM | RPD | Observações |
|---|---|---|---|---|
| `gemini-2.5-flash-lite` | 15 | 250k | ~1000 | Mais rápido (padrão) |
| `gemini-2.5-flash` | 10 | 250k | ~500 | Equilibrado |
| `gemini-2.5-pro` | 5 | 250k | ~100 | Mais inteligente |
| `gemini-3-flash-preview` | — | — | — | Preview (fallback) |
| `gemini-3.1-flash-lite-preview` | — | — | — | Preview (fallback) |
| `gemini-3.1-flash-live-preview` | — | — | — | Preview (fallback) |

Fallback automático com **sticky model**: o sistema lembra o último modelo que funcionou e tenta ele primeiro, evitando chamadas desperdiçadas em modelos com rate limit.

## Capacidades

### Excel (.xlsx) — 27 operações

| Operação | Descrição |
|---|---|
| `read` | Ler dados e metadata de todas as sheets |
| `create` | Criar arquivo com múltiplas sheets |
| `update` | Atualizar 1 célula ou múltiplas via `update_range` |
| `add` | Adicionar nova sheet com dados |
| `append` | Adicionar linhas ao final de uma sheet |
| `delete_sheet` | Remover sheet |
| `delete_rows` | Remover linhas |
| `format` | Negrito, itálico, cores, bordas, formatação numérica (R$, %, data), alinhamento |
| `auto_width` | Ajustar largura das colunas automaticamente |
| `formula` | Inserir fórmulas (SUM, AVERAGE, COUNT, VLOOKUP, IF, etc.) |
| `merge` | Mesclar células |
| `add_chart` | Adicionar gráficos (coluna, barra, linha, pizza, área, dispersão) com validação de posição |
| `list_charts` | Listar todos os gráficos existentes (título, tipo, posição, índice) |
| `delete_chart` | Remover gráficos por índice ou título |
| `update_chart` | Atualizar propriedades do gráfico (título, tipo, estilo) |
| `move_chart` | Mover gráfico para nova posição (preserva tamanho) |
| `resize_chart` | Redimensionar gráfico (escala % ou dimensões absolutas em cm) |
| `sort` | Ordenar dados por coluna (crescente/decrescente) |
| `filter_and_copy` | Filtrar dados por critérios e copiar para nova sheet ou arquivo |
| `remove_duplicates` | Remover linhas duplicadas por coluna ou linha inteira |
| `insert_rows` | Inserir linhas vazias em posição específica |
| `insert_columns` | Inserir colunas vazias em posição específica |
| `freeze_panes` | Congelar linhas e/ou colunas para visualização |
| `unfreeze_panes` | Remover congelamento de painéis |
| `find_and_replace` | Buscar e substituir texto em células (case-sensitive/insensitive) |

### Word (.docx) — 49 operações

**Conteúdo Base**

| Operação | Descrição |
|---|---|
| `read` | Ler texto completo do documento |
| `create` | Criar documento simples ou estruturado (headings + tabelas + listas) |
| `update` | Atualizar parágrafo por índice |
| `add` | Adicionar parágrafo ao final |
| `add_heading` | Adicionar título (H0 a H9) |
| `add_table` | Adicionar tabela com cabeçalho em negrito |
| `add_list` | Adicionar lista (bullets ou numerada) |
| `delete_paragraph` | Remover parágrafo por índice |
| `replace` | Buscar e substituir texto em todo o documento |
| `format` | Negrito, itálico, sublinhado, tachado, highlight, superscript, subscript, versaletes, fonte, cor, alinhamento, espaçamento |

**Conteúdo Rico** *(novo)*

| Operação | Descrição |
|---|---|
| `add_image` | Inserir imagem ao final com alinhamento, dimensões e legenda opcional |
| `add_image_at_position` | Inserir imagem em posição específica (antes de parágrafo) |
| `add_hyperlink` | Adicionar link clicável em novo parágrafo (URLs, emails mailto:) |
| `add_hyperlink_to_paragraph` | Inserir link inline no final de um parágrafo existente |

**Cabeçalho e Rodapé** *(novo)*

| Operação | Descrição |
|---|---|
| `add_header` | Adicionar cabeçalho com texto, formatação e número de página opcional |
| `add_footer` | Adicionar rodapé com texto, número de página e total de páginas (ex: 1 de 5) |
| `remove_header` | Remover conteúdo do cabeçalho |
| `remove_footer` | Remover conteúdo do rodapé |

**Layout de Página** *(novo)*

| Operação | Descrição |
|---|---|
| `set_page_margins` | Definir margens (cm ou polegadas); padrão ABNT: top=3, bottom=2, left=3, right=2 |
| `set_page_size` | Definir tamanho (A4, A3, A5, Letter, Legal, Tabloid) e orientação (portrait/landscape) |
| `get_page_layout` | Obter tamanho, dimensões, orientação e margens atuais |
| `add_page_break` | Inserir quebra de página (ao final ou após parágrafo específico) |
| `add_section_break` | Inserir quebra de seção (new_page, continuous, even_page, odd_page) |

**Melhoria de Texto com IA**

| Operação | Descrição |
|---|---|
| `improve_text` | Melhorar texto de um parágrafo ou documento inteiro |
| `correct_grammar` | Corrigir erros gramaticais e ortográficos |
| `improve_clarity` | Tornar o texto mais claro e objetivo |
| `adjust_tone` | Ajustar tom (formal, informal, técnico, persuasivo) |
| `simplify_language` | Simplificar linguagem técnica |
| `rewrite_professional` | Reescrever em estilo profissional |

**Geração de Conteúdo com IA**

| Operação | Descrição |
|---|---|
| `generate_summary` | Gerar resumo executivo do documento |
| `extract_key_points` | Extrair pontos-chave como lista |
| `create_resume` | Criar currículo profissional |
| `generate_conclusions` | Gerar seção de conclusões |
| `create_faq` | Criar seção de Perguntas Frequentes |

**Conversão de Formato**

| Operação | Descrição |
|---|---|
| `convert_list_to_table` | Converter lista em tabela |
| `convert_table_to_list` | Converter tabela em lista |
| `extract_tables_to_excel` | Exportar todas as tabelas para arquivo Excel |

**Exportação**

| Operação | Descrição |
|---|---|
| `export_to_txt` | Exportar como arquivo de texto simples |
| `export_to_markdown` | Exportar como Markdown |
| `export_to_html` | Exportar como HTML |
| `export_to_pdf` | Exportar como PDF |

**Análise de Documento**

| Operação | Descrição |
|---|---|
| `analyze_word_count` | Contagem de palavras por seção |
| `analyze_section_length` | Identificar seções muito longas |
| `get_document_statistics` | Estatísticas gerais (parágrafos, tabelas, listas, palavras) |
| `analyze_tone` | Análise de tom e formalidade com IA |
| `identify_jargon` | Identificar jargões técnicos |
| `analyze_readability` | Avaliar legibilidade e complexidade |
| `check_term_consistency` | Verificar consistência de termos |
| `analyze_document` | Análise completa combinando todas as métricas acima |

### PowerPoint (.pptx) — 10 operações

| Operação | Descrição |
|---|---|
| `read` | Ler conteúdo de todos os slides |
| `create` | Criar apresentação com layouts variados (title, content, section, blank, etc.) |
| `add` | Adicionar slide |
| `update` | Atualizar conteúdo de slide |
| `delete_slide` | Remover slide |
| `duplicate_slide` | Duplicar slide |
| `add_textbox` | Adicionar caixa de texto livre com formatação |
| `add_table` | Adicionar tabela dentro de slide |
| `replace` | Buscar e substituir texto em toda a apresentação |
| `extract_text` | Extrair texto de todos os slides |

### PDF (.pdf) — 8 operações

| Operação | Descrição |
|---|---|
| `read` | Extrair texto de todas as páginas |
| `create` | Criar PDF com texto, títulos, tabelas e listas |
| `extract_tables` | Extrair tabelas de PDFs |
| `merge` | Juntar múltiplos PDFs em um |
| `split` | Dividir PDF por range de páginas |
| `add_text` | Adicionar texto/marca d'água a páginas existentes |
| `get_info` | Obter metadata (páginas, autor, tamanho) |
| `rotate` | Rotacionar páginas |

## Exemplos de Comandos

**Excel:**
```
Crie uma planilha vendas.xlsx com colunas: Produto, Quantidade, Preço
Formate o cabeçalho com negrito e fundo azul
Adicione uma fórmula SUM na célula B11
Ajuste automaticamente a largura das colunas
Delete as linhas 5 a 8 da planilha dados.xlsx
Ordene os dados pela coluna Data em ordem crescente
Liste todos os gráficos do arquivo vendas.xlsx
Crie um gráfico de pizza na posição K2 mostrando vendas por produto
Delete o primeiro gráfico da Sheet1
Remova o gráfico 'Vendas por Produto' do arquivo vendas.xlsx
Aumente o gráfico em 50%
Redimensione o gráfico para 20cm de largura e 12cm de altura
Mova o gráfico para a célula G2
Filtre vendas acima de R$1000 e copie para nova sheet 'Vendas Altas'
Remova linhas duplicadas pela coluna Email
Insira 3 linhas vazias na posição 10
Congele a primeira linha para manter o cabeçalho visível
Substitua 'Produto A' por 'Produto Alpha' em toda a planilha
```

**Word:**
```
Crie um relatório estruturado com título, seções, tabela de dados e lista de próximos passos
Adicione uma tabela com colunas Nome, Email e Cargo ao documento equipe.docx
Substitua "2024" por "2025" em todo o documento contrato.docx
Formate o primeiro parágrafo com negrito e centralizado
Insira a imagem logo.png centralizada com 5cm de largura no relatório
Adicione o link do LinkedIn: linkedin.com/in/joao no currículo
Coloque cabeçalho com o nome da empresa e número de página à direita
Adicione rodapé com "Página 1 de 5" centralizado
Defina margens ABNT: superior 3cm, inferior 2cm, esquerda 3cm, direita 2cm
Converta o documento para paisagem A4
Adicione uma quebra de página antes da seção Conclusões
Risque o texto do parágrafo 3 (tachado)
Destaque o parágrafo 5 em amarelo (highlight)
Corriga a gramática do documento contrato.docx
Gere um resumo executivo do relatório
Exporte o documento para Markdown
Analise o tom de escrita do documento
```

**PowerPoint:**
```
Crie uma apresentação com slide de título, 3 slides de conteúdo e slide de encerramento
Adicione uma tabela de resultados no slide 2
Adicione uma caixa de texto "Confidencial" em vermelho no slide 1
Substitua "2024" por "2025" em toda a apresentação
```

**PDF:**
```
Extraia o texto do arquivo contrato.pdf
Junte os arquivos parte1.pdf e parte2.pdf em um único documento
Divida o relatório.pdf extraindo as páginas 1 a 5
Crie um PDF com o resumo executivo do projeto
Adicione "CONFIDENCIAL" como marca d'água no documento
```

## Uso Programático

```python
from src.factory import create_agent

agent = create_agent()
response = agent.process_chat_message(
    "Crie uma planilha vendas.xlsx com dados de exemplo"
)

if response.success:
    print(response.message)
    print("Arquivos modificados:", response.files_modified)
else:
    print("Erro:", response.error)
```

## Estrutura do Projeto

```
beemo/
├── app.py                  # Interface Streamlit (UI do Beemo)
├── beemo.png               # Ícone do Beemo
├── requirements.txt        # Dependências
├── .env.example            # Exemplo de configuração
├── config/                 # Módulo de configuração
├── src/
│   ├── agent.py            # Orquestrador principal (agentic loop com function calling)
│   ├── tool_definitions.py # Declarações nativas de function calling (Gemini SDK)
│   ├── factory.py          # Inicialização do sistema
│   ├── gemini_client.py    # Integração com Gemini API (function calling + fallback)
│   ├── file_scanner.py     # Varredura de arquivos
│   ├── excel_tool.py       # Manipulação de Excel (27 operações)
│   ├── word_tool.py        # Manipulação de Word (62 operações)
│   ├── powerpoint_tool.py  # Manipulação de PowerPoint (10 operações)
│   ├── pdf_tool.py         # Manipulação de PDF (8 operações)
│   ├── security_validator.py
│   ├── response_cache.py   # Cache de respostas da API
│   ├── version_manager.py  # Versionamento com undo/redo
│   ├── models.py
│   ├── exceptions.py
│   └── logging_config.py
├── tests/                  # Testes unitários
├── docs/                   # Documentação técnica detalhada
└── logs/                   # Logs da aplicação
```

## Testes

```bash
pytest tests/                                    # Todos os testes
pytest tests/ --cov=src --cov-report=term-missing # Com cobertura
pytest tests/test_excel_tool.py -v               # Módulo específico
```

## Funcionalidades Avançadas

### Validação de Posição de Gráficos
Ao adicionar gráficos, o sistema verifica automaticamente se a posição está ocupada, prevenindo sobreposição acidental. Mensagens de erro claras indicam o gráfico existente e sugerem soluções. Suporta substituição intencional com `replace_existing=True`.

**Benefícios:**
- Previne gráficos sobrepostos e invisíveis
- Mensagens de erro informativas com título do gráfico existente
- Opção de substituição intencional quando necessário
- Integrado ao sistema de versionamento (undo/redo disponível)

### Listagem de Gráficos
Liste todos os gráficos existentes em um arquivo Excel com informações detalhadas: título, tipo, posição e índice. Útil para automações que precisam verificar gráficos existentes antes de adicionar novos, evitando duplicação.

**Benefícios:**
- Inspeção programática sem abrir Excel
- Permite lógica condicional em automações (if chart exists...)
- Base para operações de delete/update de gráficos
- Auditoria e documentação automática de gráficos

### Deleção de Gráficos
Remova gráficos de planilhas Excel por índice ou título. Permite limpeza programática de gráficos antigos ou indesejados, completando o ciclo de gerenciamento de gráficos (criar, listar, deletar).

**Benefícios:**
- Limpeza programática de gráficos antigos
- Evita acúmulo de gráficos em processos repetitivos
- Permite substituição controlada de gráficos
- Integrado ao sistema de versionamento (undo/redo disponível)
- Mensagens de erro claras com lista de gráficos disponíveis

### Redimensionar e Mover Gráficos
Redimensione gráficos com escala percentual (ex: "aumente 50%") ou dimensões absolutas em cm. Mova gráficos para qualquer célula preservando o tamanho atual.

**Benefícios:**
- Escala proporcional (1.5 = 150%) ou dimensões absolutas (largura/altura em cm)
- Movimento preserva tamanho: redimensione primeiro, mova depois
- Suporta todos os tipos de anchor (TwoCellAnchor, OneCellAnchor)
- Detecção automática do tamanho atual do gráfico
- Integrado ao sistema de versionamento (undo/redo disponível)

### Ordenação de Dados
Ordene dados em planilhas Excel por qualquer coluna, em ordem crescente ou decrescente. Suporta ordenação por letra de coluna (A, B, C) ou número (1, 2, 3), com opção de preservar linha de cabeçalho e ordenar intervalos específicos.

**Benefícios:**
- Comandos em linguagem natural
- Preserva integridade das linhas (todas as colunas se movem juntas)
- Tratamento robusto de diferentes tipos de dados (texto, números, datas)
- Flexível: ordena planilha inteira ou intervalos específicos

### Filtrar e Copiar Dados
Filtre dados por critérios complexos e copie para nova sheet ou arquivo. Suporta operadores numéricos (>, <, >=, <=, ==, !=) e texto (contains, starts_with, ends_with). Preserva cabeçalhos e formatação.

**Benefícios:**
- Extração inteligente de subconjuntos de dados
- Múltiplos operadores de comparação
- Copia para nova sheet (mesmo arquivo) ou novo arquivo
- Útil para relatórios e análises segmentadas

### Remover Duplicatas
Remova linhas duplicadas por coluna específica ou linha inteira. Escolha manter primeira ou última ocorrência. Retorna estatísticas detalhadas de remoção.

**Benefícios:**
- Limpeza automática de dados
- Flexível: por coluna (ex: Email) ou linha inteira
- Estratégia configurável (keep first/last)
- Relatório de duplicatas encontradas

### Inserir Linhas e Colunas
Insira linhas ou colunas vazias em qualquer posição. Dados existentes são deslocados automaticamente. Suporta inserção múltipla (ex: 5 linhas de uma vez).

**Benefícios:**
- Adiciona espaço sem reescrever dados
- Preserva formatação e fórmulas
- Suporta letra de coluna (A, B) ou número (1, 2)
- Útil para adicionar dados no meio da planilha

### Congelar Painéis
Congele linhas superiores e/ou colunas à esquerda para manter cabeçalhos visíveis durante rolagem. Suporta congelamento horizontal, vertical ou ambos.

**Benefícios:**
- Melhora navegação em planilhas grandes
- Mantém contexto (cabeçalhos) sempre visível
- Configuração flexível (linhas, colunas ou ambos)
- Comando para remover congelamento

### Buscar e Substituir
Busque e substitua texto em células com opções avançadas: case-sensitive/insensitive, match entire cell, aplicar em sheet específica ou todas. Retorna contagem de substituições.

**Benefícios:**
- Correções em massa automatizadas
- Opções flexíveis de matching
- Aplica em uma ou todas as sheets
- Relatório detalhado de substituições por sheet

### Imagens em Documentos Word
Insira imagens (PNG, JPG, GIF, BMP, TIFF) em documentos Word com controle de tamanho (largura/altura em polegadas), alinhamento (esquerda, centro, direita) e legenda automática. Suporta inserção ao final do documento ou em posição específica (antes de um parágrafo).

**Benefícios:**
- Inserção de fotos de perfil em currículos, logos em relatórios, gráficos em documentos
- Manutenção automática de proporção (defina só largura ou só altura)
- Legenda em itálico abaixo da imagem
- Inserção posicional via índice de parágrafo

### Hyperlinks em Documentos Word
Adicione links clicáveis em documentos Word com texto de exibição, URL de destino e formatação (negrito, itálico, cor, tamanho). Suporta URLs web e endereços de email (`mailto:`). Insere em novo parágrafo ou inline no final de um parágrafo existente.

**Benefícios:**
- Links de LinkedIn, GitHub, portfólio e email em currículos
- Links automáticos com sublinhado e cor padrão Word (azul)
- Inline ou em parágrafo próprio

### Cabeçalho e Rodapé
Adicione ou substitua cabeçalho e rodapé em todas as seções do documento. Suporta texto livre, formatação (fonte, tamanho, negrito, itálico) e campos automáticos: número de página (`PAGE`) e total de páginas (`NUMPAGES`), gerando o formato profissional "Página 1 de 5".

**Benefícios:**
- Identidade visual consistente (nome da empresa, título do documento)
- Numeração automática de páginas centralizada ou à direita
- Formato "Página X de Y" para relatórios formais
- remove_header / remove_footer para limpar sem deletar a área

### Layout de Página
Control total do layout: defina margens (cm ou polegadas), tamanho de página (A4, A3, A5, Letter, Legal, Tabloid) e orientação (portrait/landscape). Inclui suporte ao padrão ABNT brasileiro (top=3, bottom=2, left=3, right=2) e `get_page_layout` para inspecionar o layout atual.

**Benefícios:**
- Conformidade com ABNT para trabalhos acadêmicos
- Landscape automático (width/height trocados) para tabelas e diagramas
- Detecção automática do tamanho de papel atual
- Base para configurações de layout por seção

### Quebras de Página e Seção
Insira quebras de página (forçam conteúdo à próxima página) ou quebras de seção (permitem layouts diferentes por seção: new_page, continuous, even_page, odd_page). Suporte a inserção posicional via índice de parágrafo.

**Benefícios:**
- Capítulos sempre em nova página
- Seções com orientação ou margens diferentes no mesmo documento
- Odd_page para livros com capítulos em página ímpar
- Continuous para mudança de colunas sem quebra visual

### Formatação Avançada de Parágrafos
Além de negrito, itálico e sublinhado, o `format` agora suporta: tachado (strikethrough), marca-texto em 16 cores (highlight), superscript (x², nota¹), subscript (H₂O), versaletes, todas-maiúsculas, espaço antes/depois do parágrafo e espaçamento entre linhas.

**Benefícios:**
- Destaque visual com marca-texto amarelo, verde, ciano, etc.
- Notação científica com superscript/subscript
- Espaçamento ABNT (1.5x = line_spacing: 18)
- Controle completo de tipografia

### Native Function Calling
O Beemo usa a API de function calling nativa do Gemini SDK. O modelo escolhe diretamente quais ferramentas chamar (`excel_operation`, `word_operation`, `pdf_operation`, `powerpoint_operation`, `read_file`, `list_files`) com os parâmetros corretos — sem JSON manual, sem parsing frágil.

**Benefícios:**
- Conversas mistas: o modelo responde em texto quando é uma pergunta, executa ferramenta quando é uma ação
- Suporte a múltiplas ferramentas por turno
- Parâmetros tipados e validados pelo schema
- Agentic loop: o modelo pode encadear leituras e escritas em sequência

### Fallback Automático de Modelos (Sticky Model)
Quando o modelo principal atinge rate limit (429), alta demanda (503) ou incompatibilidade (400), o sistema tenta automaticamente até 6 modelos alternativos. O **sticky model** memoriza o último modelo bem-sucedido e tenta ele primeiro na próxima chamada, evitando chamadas desperdiçadas.

**Cadeia:** `flash-lite → flash → pro → 3-flash-preview → 3.1-flash-lite-preview → 3.1-flash-live-preview`

### Upload de Arquivos via Interface
Envie arquivos Office diretamente pela sidebar com drag & drop ou botão "Browse files". Suporta `.xlsx`, `.docx`, `.pptx`, `.pdf`. O diretório de trabalho (`ROOT_PATH`) é criado automaticamente se não existir.

### Geração de Dados Simulados
Ao pedir "crie uma planilha com 100 linhas de dados financeiros", o sistema gera dados realistas server-side via Python (não depende do modelo enviar todos os dados). Suporta inferência inteligente de tipos de colunas:

- **Financeiro**: Data, Receita, Despesas, Lucro (calculado)
- **Vendas**: Produto, Quantidade, Valor Unitário, Total
- **RH**: Nome, Cargo, Departamento, Salário
- **Genérico**: Nomes, cidades, emails, telefones brasileiros

### Auto-Inferência Inteligente
- **Sheet name**: quando o modelo não informa a sheet, auto-detecta a primeira do workbook
- **Pie chart data**: separa automaticamente categories (labels) e values (números) para gráficos de pizza
- **JSON wrapper**: respostas acidentalmente em JSON (`{"response":"..."}`) são automaticamente convertidas em texto limpo

### Compatibilidade Gemini 3.x
Suporte a `thought_signature` para modelos Gemini 3.x que exigem assinatura de pensamento nas respostas de function call. Fallback automático quando um modelo preview retorna erro de compatibilidade.

### Versionamento e Undo/Redo
Cada modificação cria backup automático. Desfaça e refaça operações pela sidebar. Limite configurável de versões por arquivo.

### Cache de Respostas
Respostas da API são cacheadas por SHA256 do prompt. Respostas instantâneas para prompts repetidos. TTL e limite configuráveis. Estatísticas de hit rate na interface.

### Operações em Lote
Múltiplas operações em uma única solicitação. Execução sequencial com feedback individual. Continua mesmo se uma operação falhar.

## Limitações

- **Arquivos grandes**: Recomenda-se até ~10MB (limite de tokens do Gemini)
- **Imagens em Word**: Suporte a inserção de imagens (.png, .jpg, .gif, .bmp, .tiff). Imagens já existentes no documento não são lidas/processadas pelo agente.
- **Gráficos embutidos em Word/PDF**: Gráficos incorporados em documentos Word e PDF não são processados
- **Usuário único**: Sem suporte a múltiplos usuários simultâneos
- **Apenas local**: Sem integração com OneDrive, Google Drive ou SharePoint
- **Conectividade**: Requer internet para chamar a API do Gemini
