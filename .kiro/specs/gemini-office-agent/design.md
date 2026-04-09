# Design Document: Gemini Office Agent

## Overview

O Gemini Office Agent é um sistema Python que permite manipulação de arquivos Microsoft Office (Excel, Word, PowerPoint) através de comandos em linguagem natural. O sistema utiliza a API do Google Gemini 2.5 Flash-Lite para interpretar intenções do usuário e executar operações de leitura, criação e edição em arquivos locais.

### Objetivos Principais

1. Fornecer uma interface web intuitiva (Streamlit) para interação em linguagem natural
2. Integrar com a API do Gemini para interpretação de comandos
3. Manipular arquivos Office (.xlsx, .docx, .pptx) de forma programática
4. Manter arquitetura simples e modular com dependências mínimas

### Escopo

O sistema cobre:
- Interface web para entrada de comandos em linguagem natural
- Varredura automática de arquivos Office em diretórios locais
- Leitura, criação e edição de arquivos Excel, Word e PowerPoint
- Integração com Gemini API para processamento de linguagem natural
- Gerenciamento de configuração e logging

O sistema NÃO cobre:
- Edição avançada de formatação (fontes, cores, estilos complexos)
- Integração com Office 365 ou serviços cloud
- Autenticação de múltiplos usuários
- Versionamento de arquivos

## Architecture

### Visão Geral da Arquitetura

O sistema segue uma arquitetura em camadas com separação clara de responsabilidades:

```
┌─────────────────────────────────────────┐
│     Streamlit Interface Layer           │
│  (User Input/Output, UI Components)     │
└──────────────┬──────────────────────────┘
               │
┌──────────────▼──────────────────────────┐
│         Agent Orchestration Layer       │
│   (Workflow Coordination, Decision)     │
└──────┬───────────────────────┬──────────┘
       │                       │
┌──────▼────────┐    ┌────────▼──────────┐
│  Gemini API   │    │   File Scanner    │
│    Client     │    │                   │
└───────────────┘    └───────────────────┘
                              │
       ┌──────────────────────┼──────────────────────┐
       │                      │                      │
┌──────▼────────┐   ┌────────▼────────┐   ┌────────▼────────┐
│  Excel Tool   │   │   Word Tool     │   │ PowerPoint Tool │
│  (openpyxl)   │   │ (python-docx)   │   │  (python-pptx)  │
└───────────────┘   └─────────────────┘   └─────────────────┘
       │                      │                      │
       └──────────────────────┼──────────────────────┘
                              │
                    ┌─────────▼─────────┐
                    │   File System     │
                    └───────────────────┘
```

### Fluxo de Dados Principal

1. **User Input**: Usuário insere comando em linguagem natural na interface Streamlit
2. **File Discovery**: Agent solicita ao File_Scanner a lista de arquivos disponíveis
3. **Context Building**: Agent lê conteúdo relevante dos arquivos usando as ferramentas apropriadas
4. **LLM Processing**: Agent envia prompt contextualizado para Gemini API
5. **Action Interpretation**: Agent interpreta resposta do Gemini e determina ações
6. **File Operations**: Agent executa operações nos arquivos usando as ferramentas
7. **Result Display**: Interface exibe resultado ou confirmação ao usuário

### Decisões Arquiteturais

**1. Streamlit como Interface**
- Justificativa: Desenvolvimento rápido, interface web nativa, ideal para protótipos
- Trade-off: Menos controle sobre UI comparado a frameworks web completos

**2. Arquitetura Baseada em Ferramentas**
- Justificativa: Cada tipo de arquivo tem sua própria ferramenta especializada
- Trade-off: Facilita manutenção e testes, mas requer coordenação no Agent

**3. Gemini API Direta (sem LangChain)**
- Justificativa: Requisito explícito de dependências mínimas
- Trade-off: Menos abstrações prontas, mas maior controle e simplicidade

**4. Varredura de Arquivos sob Demanda**
- Justificativa: Garante lista atualizada de arquivos a cada operação
- Trade-off: Pequeno overhead, mas evita cache desatualizado

## Components and Interfaces

### 1. Streamlit Interface

**Responsabilidade**: Fornecer interface web para interação do usuário

**Interface Pública**:
```python
def main() -> None:
    """Entry point da aplicação Streamlit"""
    
def display_prompt_input() -> str:
    """Exibe campo de entrada e retorna o prompt do usuário"""
    
def display_progress(message: str) -> None:
    """Exibe mensagem de progresso"""
    
def display_result(result: str) -> None:
    """Exibe resultado da operação"""
    
def display_error(error: str) -> None:
    """Exibe mensagem de erro"""
```

**Dependências**: Agent, Streamlit

**Estado**: Session state do Streamlit para histórico de conversação

### 2. Agent

**Responsabilidade**: Coordenar o fluxo de trabalho entre componentes

**Interface Pública**:
```python
class Agent:
    def __init__(self, config: Config):
        """Inicializa o agent com configuração"""
        
    def process_user_request(self, user_prompt: str) -> str:
        """Processa solicitação do usuário e retorna resultado"""
        
    def _discover_files(self) -> List[str]:
        """Descobre arquivos Office disponíveis"""
        
    def _read_file_content(self, file_path: str) -> str:
        """Lê conteúdo de arquivo usando ferramenta apropriada"""
        
    def _build_context_prompt(self, user_prompt: str, files: List[str]) -> str:
        """Constrói prompt contextualizado para o Gemini"""
        
    def _execute_actions(self, gemini_response: str) -> str:
        """Interpreta resposta do Gemini e executa ações"""
```

**Dependências**: Gemini_Client, File_Scanner, Excel_Tool, Word_Tool, PowerPoint_Tool

**Estado**: Referências aos componentes injetados

### 3. Gemini Client

**Responsabilidade**: Encapsular comunicação com Google Generative AI API

**Interface Pública**:
```python
class GeminiClient:
    def __init__(self, api_key: str, model_name: str):
        """Inicializa cliente com API key e nome do modelo"""
        
    def generate_response(self, prompt: str, timeout: int = 30) -> str:
        """Envia prompt e retorna resposta do modelo"""
```

**Dependências**: google-generativeai

**Exceções**:
- `AuthenticationError`: Falha de autenticação
- `QuotaExceededError`: Limite de quota atingido
- `TimeoutError`: Timeout na requisição

### 4. File Scanner

**Responsabilidade**: Varrer diretórios e identificar arquivos Office

**Interface Pública**:
```python
class FileScanner:
    def __init__(self, root_path: str):
        """Inicializa scanner com pasta raiz"""
        
    def scan_office_files(self) -> List[str]:
        """Retorna lista de caminhos de arquivos Office encontrados"""
        
    def _is_office_file(self, file_path: str) -> bool:
        """Verifica se arquivo é Office válido"""
        
    def _is_temp_file(self, file_name: str) -> bool:
        """Verifica se arquivo é temporário (~$)"""
```

**Dependências**: pathlib, os

**Estado**: Caminho da pasta raiz

### 5. Excel Tool

**Responsabilidade**: Manipular arquivos Excel (.xlsx)

**Interface Pública**:
```python
class ExcelTool:
    def read_excel(self, file_path: str) -> Dict[str, Any]:
        """Lê arquivo Excel e retorna dados estruturados"""
        
    def create_excel(self, file_path: str, data: Dict[str, List[List[Any]]]) -> None:
        """Cria novo arquivo Excel com dados fornecidos"""
        
    def update_cell(self, file_path: str, sheet: str, row: int, col: int, value: Any) -> None:
        """Atualiza célula específica"""
        
    def add_sheet(self, file_path: str, sheet_name: str, data: List[List[Any]]) -> None:
        """Adiciona nova planilha ao arquivo"""
```

**Dependências**: openpyxl

**Exceções**: `CorruptedFileError`, `FileNotFoundError`

### 6. Word Tool

**Responsabilidade**: Manipular arquivos Word (.docx)

**Interface Pública**:
```python
class WordTool:
    def read_word(self, file_path: str) -> str:
        """Lê documento Word e retorna texto completo"""
        
    def create_word(self, file_path: str, content: str) -> None:
        """Cria novo documento Word com conteúdo"""
        
    def add_paragraph(self, file_path: str, text: str) -> None:
        """Adiciona parágrafo ao documento"""
        
    def update_paragraph(self, file_path: str, index: int, new_text: str) -> None:
        """Atualiza parágrafo específico"""
        
    def extract_tables(self, file_path: str) -> List[List[List[str]]]:
        """Extrai informações de tabelas do documento"""
```

**Dependências**: python-docx

**Exceções**: `CorruptedFileError`, `FileNotFoundError`

### 7. PowerPoint Tool

**Responsabilidade**: Manipular arquivos PowerPoint (.pptx)

**Interface Pública**:
```python
class PowerPointTool:
    def read_powerpoint(self, file_path: str) -> List[Dict[str, Any]]:
        """Lê apresentação e retorna informações dos slides"""
        
    def create_powerpoint(self, file_path: str, slides: List[Dict[str, Any]]) -> None:
        """Cria nova apresentação com slides fornecidos"""
        
    def add_slide(self, file_path: str, slide_data: Dict[str, Any]) -> None:
        """Adiciona novo slide à apresentação"""
        
    def update_slide(self, file_path: str, slide_index: int, new_content: Dict[str, Any]) -> None:
        """Atualiza conteúdo de slide específico"""
        
    def extract_text(self, file_path: str) -> List[str]:
        """Extrai texto de todos os slides"""
```

**Dependências**: python-pptx

**Exceções**: `CorruptedFileError`, `FileNotFoundError`

### 8. Config

**Responsabilidade**: Gerenciar configurações do sistema

**Interface Pública**:
```python
class Config:
    def __init__(self):
        """Inicializa configuração lendo variáveis de ambiente"""
        
    @property
    def api_key(self) -> str:
        """Retorna API key do Gemini"""
        
    @property
    def root_path(self) -> str:
        """Retorna caminho da pasta raiz"""
        
    @property
    def model_name(self) -> str:
        """Retorna nome do modelo Gemini"""
```

**Dependências**: os

**Exceções**: `ConfigurationError` se API key não estiver configurada

## Data Models

### FileInfo

Representa informações sobre um arquivo Office descoberto:

```python
@dataclass
class FileInfo:
    path: str              # Caminho completo do arquivo
    name: str              # Nome do arquivo
    extension: str         # Extensão (.xlsx, .docx, .pptx)
    size: int              # Tamanho em bytes
    modified_time: float   # Timestamp da última modificação
```

### ExcelData

Representa dados de um arquivo Excel:

```python
@dataclass
class ExcelData:
    sheets: Dict[str, List[List[Any]]]  # Nome da planilha -> dados
    metadata: Dict[str, Any]             # Metadados (autor, data criação, etc)
```

### WordData

Representa dados de um documento Word:

```python
@dataclass
class WordData:
    paragraphs: List[str]                # Lista de parágrafos
    tables: List[List[List[str]]]        # Lista de tabelas
    metadata: Dict[str, Any]             # Metadados
```

### PowerPointData

Representa dados de uma apresentação PowerPoint:

```python
@dataclass
class SlideData:
    index: int                           # Índice do slide
    title: str                           # Título do slide
    content: List[str]                   # Conteúdo textual
    notes: str                           # Notas do apresentador

@dataclass
class PowerPointData:
    slides: List[SlideData]              # Lista de slides
    metadata: Dict[str, Any]             # Metadados
```

### AgentResponse

Representa a resposta do Agent após processar uma solicitação:

```python
@dataclass
class AgentResponse:
    success: bool                        # Se operação foi bem-sucedida
    message: str                         # Mensagem descritiva
    files_modified: List[str]            # Arquivos que foram modificados
    error: Optional[str]                 # Mensagem de erro, se houver
```

### GeminiRequest

Representa uma requisição para o Gemini API:

```python
@dataclass
class GeminiRequest:
    prompt: str                          # Prompt completo
    user_intent: str                     # Intenção original do usuário
    context_files: List[str]             # Arquivos incluídos no contexto
    timestamp: float                     # Timestamp da requisição
```


## Correctness Properties

*A property is a characteristic or behavior that should hold true across all valid executions of a system-essentially, a formal statement about what the system should do. Properties serve as the bridge between human-readable specifications and machine-verifiable correctness guarantees.*

### Property 1: Interface Prompt Submission

*For any* user prompt submitted through the Streamlit interface, the prompt should be sent to the Agent for processing.

**Validates: Requirements 1.2**

### Property 2: Interface Result Display

*For any* operation result (success or error), the Streamlit interface should display an appropriate message to the user.

**Validates: Requirements 1.4, 1.5**

### Property 3: Recursive File Discovery

*For any* directory structure under the configured root path, the File_Scanner should discover all Office files (.xlsx, .docx, .pptx) at all directory levels, excluding temporary files starting with ~$.

**Validates: Requirements 2.1, 2.2, 2.4**

### Property 4: Complete File Paths

*For any* file discovered by the File_Scanner, the returned path should be complete and valid for file system operations.

**Validates: Requirements 2.3**

### Property 5: Gemini Client Initialization

*For any* valid API key, the Gemini_Client should initialize successfully and be ready to process requests.

**Validates: Requirements 3.1**

### Property 6: Gemini Request Processing

*For any* prompt sent to the Gemini_Client, a request should be made to the Gemini 2.5 Flash-Lite model and a text response should be returned.

**Validates: Requirements 3.2, 3.3**

### Property 7: Excel Round Trip Preservation

*For any* valid Excel data structure, creating a file with that data and then reading it back should produce equivalent data.

**Validates: Requirements 4.1, 4.2**

### Property 8: Excel Cell Update

*For any* Excel file, sheet name, cell coordinates, and value, updating the cell should result in the cell containing the new value when read back.

**Validates: Requirements 4.3**

### Property 9: Excel Sheet Addition

*For any* Excel file and new sheet data, adding a sheet should result in the file containing the new sheet with the correct data.

**Validates: Requirements 4.4**

### Property 10: Word Round Trip Preservation

*For any* text content, creating a Word document with that content and then reading it back should produce equivalent text.

**Validates: Requirements 5.1, 5.2**

### Property 11: Word Paragraph Addition

*For any* Word document and paragraph text, adding the paragraph should result in the document containing the new paragraph when read back.

**Validates: Requirements 5.3**

### Property 12: Word Paragraph Update

*For any* Word document, paragraph index, and new text, updating the paragraph should result in that paragraph containing the new text when read back.

**Validates: Requirements 5.4**

### Property 13: Word Table Extraction

*For any* Word document containing tables, extracting tables should return all tables with their complete structure and content.

**Validates: Requirements 5.5**

### Property 14: PowerPoint Round Trip Preservation

*For any* valid slide data structure, creating a PowerPoint file with those slides and then reading it back should produce equivalent slide information.

**Validates: Requirements 6.1, 6.2**

### Property 15: PowerPoint Slide Addition

*For any* PowerPoint file and new slide data, adding the slide should result in the presentation containing the new slide when read back.

**Validates: Requirements 6.3**

### Property 16: PowerPoint Slide Update

*For any* PowerPoint file, slide index, and new content, updating the slide should result in that slide containing the new content when read back.

**Validates: Requirements 6.4**

### Property 17: PowerPoint Text Extraction

*For any* PowerPoint file, extracting text should return all text content from all slides.

**Validates: Requirements 6.5**

### Property 18: Agent File Discovery and Reading

*For any* user prompt, the Agent should identify relevant files using the File_Scanner and read their content using the appropriate tool based on file extension.

**Validates: Requirements 7.1, 7.2**

### Property 19: Agent Context Building and Gemini Invocation

*For any* user prompt and identified files, the Agent should build a contextualized prompt containing both the user's intent and file content, then send it to the Gemini_Client.

**Validates: Requirements 7.3, 7.4**

### Property 20: Agent Tool Selection

*For any* action determined by the Agent, the appropriate tool should be used based on the target file type (Excel_Tool for .xlsx, Word_Tool for .docx, PowerPoint_Tool for .pptx).

**Validates: Requirements 7.6**

### Property 21: Agent Response Return

*For any* completed operation, the Agent should return an AgentResponse object containing success status, message, and modified files list.

**Validates: Requirements 7.7**

### Property 22: Config Environment Variable Reading

*For any* API key set in environment variables, the Config should successfully read and provide access to that key.

**Validates: Requirements 8.4**

### Property 23: Error Logging with Stack Trace

*For any* error occurring in any component, the system should log the error with ERROR level and include the complete stack trace.

**Validates: Requirements 9.1, 9.5**

### Property 24: Operation Lifecycle Logging

*For any* principal operation, the system should log both the start and completion of the operation with appropriate timestamps.

**Validates: Requirements 9.2**

### Property 25: Gemini API Call Logging

*For any* call to the Gemini API, the system should log the call with a timestamp.

**Validates: Requirements 9.3**

### Property 26: Appropriate Log Levels

*For any* logged event, the log level should match the event type (ERROR for errors, WARNING for warnings, INFO for informational messages).

**Validates: Requirements 9.5**

## Error Handling

### Error Categories

**1. Configuration Errors**
- Missing API key: Raise `ConfigurationError` with clear message about setting GEMINI_API_KEY
- Invalid root path: Log warning and return empty file list
- Invalid model name: Raise `ConfigurationError` with list of valid models

**2. API Errors**
- Authentication failure: Raise `AuthenticationError` with message to check API key
- Quota exceeded: Raise `QuotaExceededError` with message about rate limits
- Timeout: Raise `TimeoutError` after 30 seconds with retry suggestion
- Network errors: Raise `NetworkError` with underlying error details

**3. File Operation Errors**
- File not found: Raise `FileNotFoundError` with file path
- Corrupted file: Raise `CorruptedFileError` with file path and error details
- Permission denied: Raise `PermissionError` with file path
- Disk full: Raise `IOError` with disk space information

**4. Data Validation Errors**
- Invalid Excel data structure: Raise `ValidationError` with details about expected format
- Invalid Word content: Raise `ValidationError` with content requirements
- Invalid PowerPoint slide data: Raise `ValidationError` with slide structure requirements

### Error Handling Strategy

**Fail Fast Principle**
- Configuration errors should prevent system startup
- API authentication errors should be caught at initialization
- File corruption should be detected before attempting operations

**Graceful Degradation**
- If a file cannot be read, continue with other files and report the error
- If Gemini API is unavailable, provide clear error message but don't crash
- If a single operation fails, allow user to retry without restarting

**Error Context**
- All errors should include relevant context (file paths, operation type, input data)
- Stack traces should be logged for debugging
- User-facing errors should be clear and actionable

**Retry Logic**
- Gemini API calls: No automatic retry (user can retry manually)
- File operations: No automatic retry (may indicate permission or corruption issues)
- File scanning: Retry once if initial scan fails

### Error Propagation

```
File Tool Error → Agent catches → Logs error → Returns AgentResponse with error
                                              ↓
                                    Streamlit displays error message

Gemini API Error → Agent catches → Logs error → Returns AgentResponse with error
                                              ↓
                                    Streamlit displays error message

Config Error → System initialization fails → Clear error message to user
```

## Testing Strategy

### Dual Testing Approach

The testing strategy employs both unit tests and property-based tests to ensure comprehensive coverage:

**Unit Tests**: Focus on specific examples, edge cases, and integration points
- Specific file format examples (Excel with formulas, Word with tables, PowerPoint with images)
- Edge cases (empty files, single cell, maximum size files)
- Error conditions (corrupted files, missing API key, network failures)
- Integration between components (Agent → Tools, Interface → Agent)

**Property-Based Tests**: Verify universal properties across all inputs
- Round trip properties for all file operations
- File discovery across random directory structures
- API client behavior with random prompts
- Agent workflow with random user requests

Together, these approaches provide comprehensive coverage where unit tests catch concrete bugs and property-based tests verify general correctness.

### Property-Based Testing Configuration

**Library Selection**: Use `hypothesis` for Python property-based testing

**Test Configuration**:
- Minimum 100 iterations per property test (due to randomization)
- Each property test must reference its design document property
- Tag format: `# Feature: gemini-office-agent, Property {number}: {property_text}`

**Example Property Test Structure**:

```python
from hypothesis import given, strategies as st
import hypothesis

# Feature: gemini-office-agent, Property 7: Excel Round Trip Preservation
@given(st.lists(st.lists(st.one_of(st.integers(), st.text(), st.floats()))))
@hypothesis.settings(max_examples=100)
def test_excel_round_trip(data):
    """For any valid Excel data structure, creating a file with that data 
    and then reading it back should produce equivalent data."""
    # Create Excel file with random data
    excel_tool = ExcelTool()
    temp_file = "test_temp.xlsx"
    excel_tool.create_excel(temp_file, {"Sheet1": data})
    
    # Read back and verify
    result = excel_tool.read_excel(temp_file)
    assert result["sheets"]["Sheet1"] == data
```

### Unit Test Coverage

**Component-Level Tests**:
- `test_streamlit_interface.py`: UI component rendering and interaction
- `test_agent.py`: Workflow coordination and decision logic
- `test_gemini_client.py`: API communication and error handling
- `test_file_scanner.py`: Directory traversal and file filtering
- `test_excel_tool.py`: Excel operations and edge cases
- `test_word_tool.py`: Word operations and edge cases
- `test_powerpoint_tool.py`: PowerPoint operations and edge cases
- `test_config.py`: Configuration loading and validation

**Integration Tests**:
- `test_agent_integration.py`: End-to-end workflow from prompt to file modification
- `test_file_tools_integration.py`: Cross-tool operations (e.g., reading Excel data to create Word report)

**Edge Case Tests**:
- Empty files (0 bytes)
- Maximum size files (test performance limits)
- Files with special characters in names
- Deeply nested directory structures (10+ levels)
- Concurrent file access scenarios
- API rate limiting and quota scenarios
- Network timeout scenarios

### Test Data Management

**Fixtures**:
- Sample Office files (Excel with various data types, Word with tables, PowerPoint with multiple slides)
- Corrupted file samples for error testing
- Directory structures for scanner testing

**Mocking Strategy**:
- Mock Gemini API for unit tests (use `unittest.mock`)
- Mock file system for scanner tests when needed
- Real file operations for integration tests

### Continuous Testing

**Pre-commit Hooks**:
- Run unit tests before allowing commits
- Run linting and type checking

**CI/CD Pipeline**:
- Run full test suite (unit + property tests) on every push
- Test on multiple Python versions (3.8, 3.9, 3.10, 3.11, 3.12)
- Generate coverage reports (target: 80%+ coverage)

### Performance Testing

While not the primary focus, basic performance benchmarks should be established:
- File scanning: < 1 second for 1000 files
- Excel read: < 2 seconds for 10MB file
- Word read: < 1 second for 100 pages
- PowerPoint read: < 2 seconds for 50 slides
- Gemini API call: < 5 seconds (excluding network latency)

