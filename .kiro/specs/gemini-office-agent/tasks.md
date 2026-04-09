# Implementation Plan: Gemini Office Agent

## Overview

Este plano de implementação detalha as tarefas necessárias para construir o Gemini Office Agent, um sistema Python que permite manipular arquivos Office através de comandos em linguagem natural. A implementação seguirá uma abordagem incremental, começando pela infraestrutura básica, depois os componentes de manipulação de arquivos, integração com Gemini API, e finalmente a interface Streamlit.

## Tasks

- [x] 1. Configurar estrutura do projeto e dependências
  - Criar estrutura de diretórios (src/, tests/, config/)
  - Criar requirements.txt com dependências: google-generativeai, openpyxl, python-docx, python-pptx, streamlit, hypothesis
  - Criar arquivo .env.example para documentar variáveis de ambiente necessárias
  - Criar .gitignore (incluir .env, __pycache__/, *.pyc, ~$*, .pytest_cache/, .hypothesis/, arquivos de teste temporários)
  - Configurar logging básico do Python
  - _Requirements: 10.1, 10.2, 10.3, 10.4, 10.5_

- [x] 2. Criar data models e exceções customizadas
  - [x] 2.1 Implementar dataclasses para modelos de dados
    - Criar FileInfo, ExcelData, WordData, PowerPointData, SlideData
    - Criar AgentResponse, GeminiRequest
    - _Requirements: Design Document - Data Models_

  - [x] 2.2 Implementar exceções customizadas
    - Criar ConfigurationError, AuthenticationError, QuotaExceededError
    - Criar CorruptedFileError, ValidationError, NetworkError
    - _Requirements: Design Document - Error Handling_

- [-] 3. Implementar módulo de configuração (Config)
  - [x] 3.1 Criar classe Config com leitura de variáveis de ambiente
    - Implementar propriedades: api_key, root_path, model_name
    - Adicionar validação de API key obrigatória
    - Usar exceção ConfigurationError (definida na Task 2)
    - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5_

  - [ ]* 3.2 Escrever testes unitários para Config
    - Testar leitura de variáveis de ambiente
    - Testar exceção quando API key não está configurada
    - Testar valores padrão
    - _Requirements: 8.1, 8.2, 8.3, 8.4, 8.5_

  - [ ]* 3.3 Escrever teste de propriedade para leitura de variáveis de ambiente
    - **Property 22: Config Environment Variable Reading**
    - **Validates: Requirements 8.4**
    - Gerar API keys aleatórias em variáveis de ambiente e verificar leitura correta
    - _Requirements: 8.4_

- [x] 4. Implementar File Scanner
  - [x] 4.1 Criar classe FileScanner com varredura recursiva
    - Implementar scan_office_files() para varrer diretórios
    - Implementar filtros para extensões .xlsx, .docx, .pptx
    - Implementar filtro para arquivos temporários (~$)
    - Adicionar tratamento para pasta raiz inexistente
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5_

  - [ ]* 4.2 Escrever testes unitários para FileScanner
    - Testar varredura em estrutura de diretórios mock
    - Testar filtro de extensões
    - Testar exclusão de arquivos temporários
    - Testar comportamento com pasta inexistente
    - _Requirements: 2.1, 2.2, 2.3, 2.4, 2.5_

  - [ ]* 4.3 Escrever teste de propriedade para descoberta recursiva
    - **Property 3: Recursive File Discovery**
    - **Validates: Requirements 2.1, 2.2, 2.4**
    - Gerar estruturas de diretórios aleatórias e verificar que todos os arquivos Office são descobertos
    - _Requirements: 2.1, 2.2, 2.4_

  - [ ]* 4.4 Escrever teste de propriedade para caminhos completos
    - **Property 4: Complete File Paths**
    - **Validates: Requirements 2.3**
    - Verificar que todos os caminhos retornados são válidos e completos para operações de arquivo
    - _Requirements: 2.3_

- [x] 5. Implementar Excel Tool
  - [x] 5.1 Criar classe ExcelTool com operações básicas
    - Implementar read_excel() usando openpyxl
    - Implementar create_excel() para criar novos arquivos
    - Implementar update_cell() para modificar células
    - Implementar add_sheet() para adicionar planilhas
    - Usar exceções customizadas (CorruptedFileError, FileNotFoundError)
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6_

  - [ ]* 5.2 Escrever testes unitários para ExcelTool
    - Testar leitura de arquivo Excel com dados variados
    - Testar criação de novo arquivo
    - Testar atualização de células
    - Testar adição de planilhas
    - Testar tratamento de arquivo corrompido
    - _Requirements: 4.1, 4.2, 4.3, 4.4, 4.5, 4.6_

  - [ ]* 5.3 Escrever teste de propriedade para round trip Excel
    - **Property 7: Excel Round Trip Preservation**
    - **Validates: Requirements 4.1, 4.2**
    - Gerar dados aleatórios, criar arquivo, ler de volta e verificar equivalência
    - _Requirements: 4.1, 4.2_

  - [ ]* 5.4 Escrever teste de propriedade para atualização de células
    - **Property 8: Excel Cell Update**
    - **Validates: Requirements 4.3**
    - Gerar coordenadas e valores aleatórios, atualizar célula e verificar leitura
    - _Requirements: 4.3_

  - [ ]* 5.5 Escrever teste de propriedade para adição de planilhas
    - **Property 9: Excel Sheet Addition**
    - **Validates: Requirements 4.4**
    - Gerar dados de planilha aleatórios, adicionar planilha e verificar que existe com dados corretos
    - _Requirements: 4.4_

- [x] 6. Implementar Word Tool
  - [x] 6.1 Criar classe WordTool com operações básicas
    - Implementar read_word() usando python-docx
    - Implementar create_word() para criar novos documentos
    - Implementar add_paragraph() para adicionar parágrafos
    - Implementar update_paragraph() para modificar parágrafos
    - Implementar extract_tables() para extrair tabelas
    - Usar exceções customizadas (CorruptedFileError, FileNotFoundError)
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6_

  - [ ]* 6.2 Escrever testes unitários para WordTool
    - Testar leitura de documento Word
    - Testar criação de novo documento
    - Testar adição de parágrafos
    - Testar atualização de parágrafos
    - Testar extração de tabelas
    - Testar tratamento de arquivo corrompido
    - _Requirements: 5.1, 5.2, 5.3, 5.4, 5.5, 5.6_

  - [ ]* 6.3 Escrever teste de propriedade para round trip Word
    - **Property 10: Word Round Trip Preservation**
    - **Validates: Requirements 5.1, 5.2**
    - Gerar texto aleatório, criar documento, ler de volta e verificar equivalência
    - _Requirements: 5.1, 5.2_

  - [ ]* 6.4 Escrever teste de propriedade para adição de parágrafos
    - **Property 11: Word Paragraph Addition**
    - **Validates: Requirements 5.3**
    - Gerar parágrafos aleatórios, adicionar ao documento e verificar que existem quando lidos
    - _Requirements: 5.3_

  - [ ]* 6.5 Escrever teste de propriedade para atualização de parágrafos
    - **Property 12: Word Paragraph Update**
    - **Validates: Requirements 5.4**
    - Gerar índices e textos aleatórios, atualizar parágrafo e verificar leitura
    - _Requirements: 5.4_

- [x] 7. Implementar PowerPoint Tool
  - [x] 7.1 Criar classe PowerPointTool com operações básicas
    - Implementar read_powerpoint() usando python-pptx
    - Implementar create_powerpoint() para criar novas apresentações
    - Implementar add_slide() para adicionar slides
    - Implementar update_slide() para modificar slides
    - Implementar extract_text() para extrair texto
    - Usar exceções customizadas (CorruptedFileError, FileNotFoundError)
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6_

  - [ ]* 7.2 Escrever testes unitários para PowerPointTool
    - Testar leitura de apresentação PowerPoint
    - Testar criação de nova apresentação
    - Testar adição de slides
    - Testar atualização de slides
    - Testar extração de texto
    - Testar tratamento de arquivo corrompido
    - _Requirements: 6.1, 6.2, 6.3, 6.4, 6.5, 6.6_

  - [ ]* 7.3 Escrever teste de propriedade para round trip PowerPoint
    - **Property 14: PowerPoint Round Trip Preservation**
    - **Validates: Requirements 6.1, 6.2**
    - Gerar dados de slides aleatórios, criar apresentação, ler de volta e verificar equivalência
    - _Requirements: 6.1, 6.2_

  - [ ]* 7.4 Escrever teste de propriedade para adição de slides
    - **Property 15: PowerPoint Slide Addition**
    - **Validates: Requirements 6.3**
    - Gerar dados de slide aleatórios, adicionar slide e verificar que existe quando lido
    - _Requirements: 6.3_

  - [ ]* 7.5 Escrever teste de propriedade para atualização de slides
    - **Property 16: PowerPoint Slide Update**
    - **Validates: Requirements 6.4**
    - Gerar índices e conteúdos aleatórios, atualizar slide e verificar leitura
    - _Requirements: 6.4_

- [x] 8. Checkpoint - Verificar ferramentas de manipulação de arquivos
  - Executar todos os testes das ferramentas (Excel, Word, PowerPoint)
  - Verificar que todas as operações básicas funcionam corretamente
  - Perguntar ao usuário se há dúvidas ou ajustes necessários

- [x] 9. Implementar Gemini Client
  - [x] 9.1 Criar classe GeminiClient com integração à API
    - Implementar inicialização com API key e model name
    - Implementar generate_response() com timeout de 30 segundos
    - Usar exceções customizadas (AuthenticationError, QuotaExceededError, TimeoutError, NetworkError)
    - Adicionar logging de chamadas à API com timestamps
    - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 9.3_

  - [ ]* 9.2 Escrever testes unitários para GeminiClient
    - Testar inicialização com API key válida
    - Testar envio de prompt e recebimento de resposta (usando mock)
    - Testar tratamento de erro de autenticação
    - Testar tratamento de erro de quota
    - Testar timeout
    - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6_

  - [ ]* 9.3 Escrever teste de propriedade para processamento de requisições
    - **Property 6: Gemini Request Processing**
    - **Validates: Requirements 3.2, 3.3**
    - Gerar prompts aleatórios e verificar que requisições são feitas e respostas retornadas
    - _Requirements: 3.2, 3.3_

  - [ ]* 9.4 Escrever teste de propriedade para inicialização
    - **Property 5: Gemini Client Initialization**
    - **Validates: Requirements 3.1**
    - Verificar que para qualquer API key válida, o cliente inicializa com sucesso
    - _Requirements: 3.1_

- [x] 10. Criar templates de prompt e estratégia de parsing
  - [x] 10.1 Criar módulo prompt_templates.py
    - Criar template de system prompt que instrui o Gemini sobre capacidades disponíveis
    - Criar template para incluir contexto de arquivos no prompt
    - Definir formato de resposta estruturado (JSON) que o Gemini deve retornar
    - Incluir exemplos de respostas válidas no prompt
    - _Requirements: 7.3, 7.4, 7.5_

  - [x] 10.2 Criar módulo response_parser.py
    - Implementar função para parsear resposta JSON do Gemini
    - Implementar fallback para parsing de texto livre se JSON falhar
    - Validar estrutura da resposta (ações, arquivos alvo, parâmetros)
    - Usar exceção ValidationError para respostas inválidas
    - _Requirements: 7.5, 7.6_

  - [ ]* 10.3 Escrever testes para parsing de respostas
    - Testar parsing de respostas JSON válidas
    - Testar fallback para texto livre
    - Testar validação de estrutura
    - Testar tratamento de respostas malformadas
    - _Requirements: 7.5, 7.6_

- [-] 11. Implementar validação de segurança
  - [x] 11.1 Criar módulo security_validator.py
    - Implementar validação de caminhos de arquivo (prevenir path traversal)
    - Implementar whitelist de operações permitidas
    - Implementar verificação de que arquivos alvo estão dentro do root_path
    - Implementar sanitização de nomes de arquivo
    - _Requirements: Design Document - Security Considerations_

  - [ ]* 11.2 Escrever testes para validação de segurança
    - Testar detecção de path traversal (../, absolute paths)
    - Testar rejeição de arquivos fora do root_path
    - Testar sanitização de nomes de arquivo
    - _Requirements: Design Document - Security Considerations_

- [x] 12. Implementar Agent (orquestração)
  - [x] 12.1 Criar classe Agent com coordenação de workflow
    - Implementar process_user_request() como método principal
    - Implementar _discover_files() usando FileScanner
    - Implementar _filter_relevant_files() para selecionar arquivos relevantes baseado no prompt
    - Implementar _read_file_content() com seleção de ferramenta por extensão
    - Implementar _build_context_prompt() usando prompt_templates
    - Implementar _execute_actions() usando response_parser e validação de segurança
    - Adicionar logging de início e fim de operações
    - Retornar AgentResponse com status, mensagem e arquivos modificados
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7, 7.8, 9.2_

  - [ ]* 12.2 Escrever testes unitários para Agent
    - Testar descoberta e leitura de arquivos
    - Testar filtragem de arquivos relevantes
    - Testar construção de prompt contextualizado
    - Testar seleção de ferramenta correta por tipo de arquivo
    - Testar retorno de AgentResponse
    - Testar tratamento de erros
    - _Requirements: 7.1, 7.2, 7.3, 7.4, 7.5, 7.6, 7.7, 7.8_

  - [ ]* 12.3 Escrever teste de propriedade para seleção de ferramentas
    - **Property 20: Agent Tool Selection**
    - **Validates: Requirements 7.6**
    - Gerar arquivos de tipos aleatórios e verificar que a ferramenta correta é usada
    - _Requirements: 7.6_

  - [ ]* 12.4 Escrever teste de propriedade para descoberta e leitura
    - **Property 18: Agent File Discovery and Reading**
    - **Validates: Requirements 7.1, 7.2**
    - Verificar que para qualquer prompt, arquivos relevantes são identificados e lidos com ferramenta apropriada
    - _Requirements: 7.1, 7.2_

  - [ ]* 12.5 Escrever teste de propriedade para construção de contexto
    - **Property 19: Agent Context Building and Gemini Invocation**
    - **Validates: Requirements 7.3, 7.4**
    - Verificar que prompt contextualizado contém user intent e file content, e é enviado ao Gemini
    - _Requirements: 7.3, 7.4_

- [-] 13. Criar factory para inicialização do sistema
  - [x] 13.1 Criar módulo factory.py
    - Implementar função create_agent() que instancia todos os componentes
    - Instanciar Config e validar configuração
    - Instanciar FileScanner com root_path do Config
    - Instanciar GeminiClient com api_key e model_name do Config
    - Instanciar todas as ferramentas (ExcelTool, WordTool, PowerPointTool)
    - Injetar dependências no Agent
    - Retornar Agent configurado e pronto para uso
    - _Requirements: Design Document - Components and Interfaces_

  - [ ]* 13.2 Escrever testes para factory
    - Testar criação completa do Agent
    - Testar que todas as dependências são injetadas corretamente
    - Testar tratamento de erro de configuração
    - _Requirements: Design Document - Components and Interfaces_

- [x] 14. Implementar interface Streamlit
  - [x] 14.1 Criar aplicação Streamlit com interface de usuário
    - Implementar main() como entry point usando factory.create_agent()
    - Implementar campo de entrada de texto para User_Prompt
    - Implementar botão de submit
    - Implementar área de exibição de resultados
    - Implementar área de exibição de erros com st.error()
    - Adicionar spinner (st.spinner) durante processamento
    - Adicionar mensagens de status para cada etapa (scanning, reading, calling Gemini, executing)
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5_

  - [x] 14.2 Adicionar seleção de arquivos na interface
    - Exibir lista de arquivos descobertos pelo FileScanner
    - Adicionar checkboxes para usuário selecionar arquivos relevantes
    - Passar apenas arquivos selecionados para o Agent
    - Adicionar opção "Selecionar todos" e "Limpar seleção"
    - _Requirements: Design Document - Usability_

  - [x] 14.3 Adicionar histórico de conversação
    - Usar st.session_state para armazenar histórico
    - Exibir histórico de prompts e resultados anteriores
    - Adicionar botão para limpar histórico
    - _Requirements: 1.3, 1.4_

  - [ ]* 14.4 Adicionar confirmação para operações destrutivas
    - Detectar quando Agent vai sobrescrever ou deletar arquivos
    - Exibir modal de confirmação com st.dialog() ou st.warning()
    - Só executar ação após confirmação do usuário
    - _Requirements: Design Document - Security_

  - [ ]* 14.5 Escrever testes para interface Streamlit
    - Testar renderização de componentes (usando streamlit testing)
    - Testar integração com Agent
    - Testar exibição de resultados e erros
    - Testar seleção de arquivos
    - _Requirements: 1.1, 1.2, 1.3, 1.4, 1.5_

- [x] 15. Implementar logging e tratamento de erros global
  - [x] 15.1 Configurar sistema de logging completo
    - Configurar formato de log com timestamps
    - Configurar níveis de log (INFO, WARNING, ERROR)
    - Adicionar logging de stack traces para erros
    - Adicionar logging de operações de arquivo
    - Configurar arquivo de log (logs/agent.log)
    - _Requirements: 9.1, 9.2, 9.4, 9.5_

  - [ ]* 15.2 Escrever testes para logging
    - Testar que erros são logados com stack trace
    - Testar que operações são logadas com timestamps
    - Testar níveis de log apropriados
    - _Requirements: 9.1, 9.2, 9.4, 9.5_

  - [ ]* 15.3 Escrever teste de propriedade para logging de erros
    - **Property 23: Error Logging with Stack Trace**
    - **Validates: Requirements 9.1, 9.5**
    - Verificar que para qualquer erro, o sistema loga com ERROR level e stack trace completo
    - _Requirements: 9.1, 9.5_

  - [ ]* 15.4 Escrever teste de propriedade para logging de operações
    - **Property 24: Operation Lifecycle Logging**
    - **Validates: Requirements 9.2**
    - Verificar que para qualquer operação principal, início e fim são logados com timestamps
    - _Requirements: 9.2_

- [ ] 16. Criar testes de integração end-to-end
  - [ ] 16.1 Escrever teste de integração completo
    - Criar diretório de teste com arquivos Office de exemplo
    - Simular prompt do usuário (ex: "Adicione uma coluna 'Total' na planilha vendas.xlsx")
    - Verificar que Agent processa corretamente
    - Verificar que arquivo é modificado conforme esperado
    - Verificar que AgentResponse contém informações corretas
    - Limpar arquivos de teste após execução
    - _Requirements: Design Document - Testing Strategy_

  - [ ]* 16.2 Escrever testes de integração para cada tipo de arquivo
    - Teste end-to-end para operação Excel
    - Teste end-to-end para operação Word
    - Teste end-to-end para operação PowerPoint
    - Teste end-to-end para operação multi-arquivo
    - _Requirements: Design Document - Testing Strategy_

  - [ ]* 16.3 Escrever teste de propriedade para resposta do Agent
    - **Property 21: Agent Response Return**
    - **Validates: Requirements 7.7**
    - Verificar que para qualquer operação completa, AgentResponse contém success, message e files_modified
    - _Requirements: 7.7_

- [ ] 17. Criar documentação e exemplos
  - [x] 17.1 Criar README.md com instruções de instalação e uso
    - Documentar requisitos (Python 3.8+)
    - Documentar instalação de dependências (pip install -r requirements.txt)
    - Documentar configuração de variáveis de ambiente (.env)
    - Incluir exemplos de comandos em linguagem natural
    - Incluir seção de troubleshooting
    - Incluir seção de limitações conhecidas
    - _Requirements: 10.5_

  - [ ] 17.2 Atualizar .env.example com documentação completa
    - Documentar GEMINI_API_KEY com link para obter chave
    - Documentar ROOT_PATH com exemplo de caminho
    - Documentar MODEL_NAME com valor padrão (gemini-2.5-flash-lite)
    - _Requirements: 8.1, 8.2, 8.3_

  - [ ] 17.3 Criar script de exemplo para teste rápido
    - Criar example.py com uso programático do Agent
    - Criar exemplos de criação de arquivos Office
    - Incluir comentários explicativos
    - _Requirements: Design Document - Overview_

  - [ ]* 17.4 Criar guia de contribuição
    - Documentar estrutura do código
    - Documentar como adicionar novas ferramentas
    - Documentar padrões de código e testes
    - _Requirements: Design Document - Overview_

## Notes

- Tasks marcadas com `*` são opcionais e podem ser puladas para um MVP mais rápido
- Cada task referencia requisitos específicos para rastreabilidade
- Checkpoints garantem validação incremental
- Testes de propriedade validam propriedades universais de correção
- Testes unitários validam exemplos específicos e casos extremos
- A implementação usa Python 3.8+ com type hints
- O sistema deve funcionar sem LangChain ou pywin32 (dependências mínimas)
- Task 2 (data models) vem antes das implementações pois outros componentes dependem dessas definições
- Task 10 (prompt templates) é crítica para o Agent funcionar corretamente
- Task 11 (segurança) previne operações perigosas como path traversal
- Task 13 (factory) garante inicialização correta de todos os componentes
