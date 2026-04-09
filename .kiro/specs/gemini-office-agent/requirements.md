# Requirements Document

## Introduction

O Gemini Office Agent é um sistema Python que permite aos usuários manipular arquivos Office (Excel, Word, PowerPoint) através de comandos em linguagem natural. O sistema utiliza a API do Gemini 2.5 Flash-Lite para interpretar as intenções do usuário e executar operações de leitura, criação e edição em arquivos locais através de uma interface Streamlit.

## Glossary

- **Agent**: O componente principal que coordena o fluxo de trabalho entre a interface do usuário, o Gemini API e as ferramentas de manipulação de arquivos
- **Gemini_Client**: O wrapper que encapsula as chamadas à API do Google Generative AI
- **File_Scanner**: O componente responsável por varrer pastas locais e identificar arquivos Office disponíveis
- **Excel_Tool**: O módulo que manipula arquivos .xlsx usando openpyxl
- **Word_Tool**: O módulo que manipula arquivos .docx usando python-docx
- **PowerPoint_Tool**: O módulo que manipula arquivos .pptx usando python-pptx
- **User_Prompt**: A descrição em linguagem natural fornecida pelo usuário sobre a operação desejada
- **Streamlit_Interface**: A interface web que permite ao usuário interagir com o Agent

## Requirements

### Requirement 1: Interface de Usuário Streamlit

**User Story:** Como usuário, eu quero descrever minhas intenções em linguagem natural através de uma interface web, para que eu possa manipular arquivos Office sem conhecer programação.

#### Acceptance Criteria

1. THE Streamlit_Interface SHALL exibir um campo de texto para entrada do User_Prompt
2. WHEN o usuário submete um User_Prompt, THE Streamlit_Interface SHALL enviar o prompt para o Agent
3. THE Streamlit_Interface SHALL exibir o progresso da operação em tempo real
4. WHEN o Agent completa uma operação, THE Streamlit_Interface SHALL exibir o resultado ou mensagem de confirmação
5. IF ocorrer um erro durante a execução, THEN THE Streamlit_Interface SHALL exibir uma mensagem de erro descritiva

### Requirement 2: Varredura de Arquivos Locais

**User Story:** Como usuário, eu quero que o sistema identifique automaticamente os arquivos Office disponíveis na pasta configurada, para que eu possa referenciá-los nas minhas solicitações.

#### Acceptance Criteria

1. THE File_Scanner SHALL varrer a pasta raiz configurada recursivamente
2. THE File_Scanner SHALL identificar todos os arquivos com extensões .xlsx, .docx e .pptx
3. WHEN solicitado, THE File_Scanner SHALL retornar uma lista com os caminhos completos dos arquivos encontrados
4. THE File_Scanner SHALL ignorar arquivos temporários do Office (que começam com ~$)
5. IF a pasta raiz não existir, THEN THE File_Scanner SHALL retornar uma lista vazia e registrar um aviso

### Requirement 3: Integração com Gemini API

**User Story:** Como desenvolvedor, eu quero um wrapper simples para a API do Gemini, para que o Agent possa enviar prompts e receber respostas estruturadas.

#### Acceptance Criteria

1. THE Gemini_Client SHALL inicializar com a API key configurada
2. WHEN recebe um prompt, THE Gemini_Client SHALL enviar a requisição para o modelo Gemini 2.5 Flash-Lite
3. THE Gemini_Client SHALL retornar a resposta do modelo como texto
4. IF a API retornar erro de autenticação, THEN THE Gemini_Client SHALL lançar uma exceção descritiva
5. IF a API retornar erro de quota ou rate limit, THEN THE Gemini_Client SHALL lançar uma exceção descritiva
6. THE Gemini_Client SHALL incluir timeout de 30 segundos nas requisições

### Requirement 4: Manipulação de Arquivos Excel

**User Story:** Como usuário, eu quero ler, criar e editar planilhas Excel, para que eu possa automatizar tarefas com dados tabulares.

#### Acceptance Criteria

1. THE Excel_Tool SHALL ler o conteúdo de arquivos .xlsx e retornar os dados em formato estruturado
2. THE Excel_Tool SHALL criar novos arquivos .xlsx com dados fornecidos
3. THE Excel_Tool SHALL modificar células específicas em planilhas existentes
4. THE Excel_Tool SHALL adicionar novas planilhas em arquivos existentes
5. WHEN salva modificações, THE Excel_Tool SHALL preservar a formatação existente quando possível
6. IF um arquivo .xlsx estiver corrompido, THEN THE Excel_Tool SHALL retornar uma mensagem de erro descritiva

### Requirement 5: Manipulação de Arquivos Word

**User Story:** Como usuário, eu quero ler, criar e editar documentos Word, para que eu possa automatizar a criação e modificação de documentos de texto.

#### Acceptance Criteria

1. THE Word_Tool SHALL ler o conteúdo de arquivos .docx e retornar o texto completo
2. THE Word_Tool SHALL criar novos arquivos .docx com conteúdo fornecido
3. THE Word_Tool SHALL adicionar parágrafos a documentos existentes
4. THE Word_Tool SHALL modificar parágrafos específicos em documentos existentes
5. THE Word_Tool SHALL extrair e retornar informações sobre tabelas presentes no documento
6. IF um arquivo .docx estiver corrompido, THEN THE Word_Tool SHALL retornar uma mensagem de erro descritiva

### Requirement 6: Manipulação de Arquivos PowerPoint

**User Story:** Como usuário, eu quero ler, criar e editar apresentações PowerPoint, para que eu possa automatizar a criação de slides.

#### Acceptance Criteria

1. THE PowerPoint_Tool SHALL ler o conteúdo de arquivos .pptx e retornar informações sobre os slides
2. THE PowerPoint_Tool SHALL criar novos arquivos .pptx com slides fornecidos
3. THE PowerPoint_Tool SHALL adicionar novos slides a apresentações existentes
4. THE PowerPoint_Tool SHALL modificar o conteúdo de slides específicos
5. THE PowerPoint_Tool SHALL extrair texto de todos os slides
6. IF um arquivo .pptx estiver corrompido, THEN THE PowerPoint_Tool SHALL retornar uma mensagem de erro descritiva

### Requirement 7: Loop Principal do Agent

**User Story:** Como desenvolvedor, eu quero um componente central que coordene o fluxo de trabalho, para que o sistema processe as solicitações do usuário de forma estruturada.

#### Acceptance Criteria

1. WHEN recebe um User_Prompt, THE Agent SHALL identificar os arquivos relevantes usando o File_Scanner
2. THE Agent SHALL ler o conteúdo dos arquivos identificados usando as ferramentas apropriadas
3. THE Agent SHALL construir um prompt contextualizado incluindo o User_Prompt e o conteúdo dos arquivos
4. THE Agent SHALL enviar o prompt para o Gemini_Client
5. THE Agent SHALL interpretar a resposta do Gemini_Client e determinar as ações necessárias
6. THE Agent SHALL executar as ações usando as ferramentas apropriadas (Excel_Tool, Word_Tool, PowerPoint_Tool)
7. THE Agent SHALL retornar o resultado da operação para a interface
8. IF o Gemini_Client retornar uma resposta ambígua, THEN THE Agent SHALL solicitar clarificação ao usuário

### Requirement 8: Gerenciamento de Configuração

**User Story:** Como usuário, eu quero configurar a API key e a pasta raiz em um único local, para que eu possa personalizar o comportamento do sistema facilmente.

#### Acceptance Criteria

1. THE Config SHALL armazenar a API key do Gemini
2. THE Config SHALL armazenar o caminho da pasta raiz para varredura de arquivos
3. THE Config SHALL armazenar o nome do modelo Gemini a ser utilizado
4. THE Config SHALL permitir leitura de variáveis de ambiente para a API key
5. IF a API key não estiver configurada, THEN THE Config SHALL lançar uma exceção ao inicializar o sistema

### Requirement 9: Tratamento de Erros e Logging

**User Story:** Como desenvolvedor, eu quero que o sistema registre erros e operações importantes, para que eu possa diagnosticar problemas facilmente.

#### Acceptance Criteria

1. WHEN ocorre um erro em qualquer componente, THE System SHALL registrar o erro com stack trace completo
2. THE System SHALL registrar o início e fim de cada operação principal
3. THE System SHALL registrar as chamadas à API do Gemini com timestamps
4. IF um arquivo não puder ser acessado, THEN THE System SHALL registrar o caminho e a razão da falha
5. THE System SHALL usar níveis de log apropriados (INFO, WARNING, ERROR)

### Requirement 10: Dependências Mínimas

**User Story:** Como desenvolvedor, eu quero manter o projeto com dependências mínimas, para que a instalação e manutenção sejam simples.

#### Acceptance Criteria

1. THE System SHALL utilizar apenas as bibliotecas: google-generativeai, openpyxl, python-docx, python-pptx, streamlit
2. THE System SHALL NOT utilizar LangChain
3. THE System SHALL NOT utilizar pywin32
4. THE System SHALL NOT incluir dependências desnecessárias no requirements.txt
5. THE System SHALL funcionar com Python 3.8 ou superior
