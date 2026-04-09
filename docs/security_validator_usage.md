# Security Validator - Guia de Uso

## Visão Geral

O módulo `security_validator.py` fornece validação de segurança para operações de arquivo no Gemini Office Agent. Ele previne ataques de path traversal, garante que operações ocorram apenas dentro do diretório raiz configurado, e sanitiza nomes de arquivo.

## Funcionalidades

### 1. Validação de Caminhos de Arquivo

Previne path traversal e garante que arquivos estão dentro do `root_path`:

```python
from src.security_validator import SecurityValidator

validator = SecurityValidator(root_path="/home/user/documents")

# Caminho válido dentro do root_path
safe_path = validator.validate_file_path("reports/sales.xlsx")
# Retorna: /home/user/documents/reports/sales.xlsx

# Tentativa de path traversal - lança ValidationError
try:
    validator.validate_file_path("../../../etc/passwd")
except ValidationError as e:
    print(f"Bloqueado: {e}")
```

### 2. Whitelist de Operações

Valida que apenas operações permitidas são executadas:

```python
# Operações permitidas
ALLOWED_OPERATIONS = {
    'read', 'create', 'update', 'add_sheet', 'add_paragraph',
    'add_slide', 'update_cell', 'update_paragraph', 'update_slide',
    'extract_tables', 'extract_text'
}

# Validar operação
validator.validate_operation("read")  # OK
validator.validate_operation("delete")  # Lança ValidationError
```

### 3. Sanitização de Nomes de Arquivo

Remove caracteres perigosos de nomes de arquivo:

```python
# Remove caracteres especiais e path separators
clean_name = validator.sanitize_filename("test@file#.xlsx")
# Retorna: "testfile.xlsx"

# Remove tentativas de path traversal
clean_name = validator.sanitize_filename("../../../evil.xlsx")
# Retorna: "evil.xlsx"

# Remove pontos no início (arquivos ocultos)
clean_name = validator.sanitize_filename("...hidden.xlsx")
# Retorna: "hidden.xlsx"
```

### 4. Validação de Extensões

Verifica que arquivos têm extensões permitidas:

```python
allowed_extensions = {'.xlsx', '.docx', '.pptx'}

# Extensão válida
validator.validate_file_extension("report.xlsx", allowed_extensions)  # OK

# Extensão inválida
validator.validate_file_extension("script.py", allowed_extensions)  # ValidationError
```

## Integração com Ferramentas

### Exemplo: Excel Tool com Validação de Segurança

```python
from src.excel_tool import ExcelTool
from src.security_validator import SecurityValidator
from config import Config

# Inicializar
config = Config()
validator = SecurityValidator(config.root_path)
excel_tool = ExcelTool()

# Validar caminho antes de ler
user_input = "reports/sales.xlsx"
safe_path = validator.validate_file_path(user_input)

# Validar operação
validator.validate_operation("read")

# Executar operação com caminho validado
data = excel_tool.read_excel(str(safe_path))
```

### Exemplo: Agent com Validação de Segurança

```python
class Agent:
    def __init__(self, config: Config):
        self.config = config
        self.validator = SecurityValidator(config.root_path)
        self.excel_tool = ExcelTool()
        # ... outras ferramentas
    
    def execute_file_operation(self, operation: str, file_path: str):
        # Validar operação
        self.validator.validate_operation(operation)
        
        # Validar caminho
        safe_path = self.validator.validate_file_path(file_path)
        
        # Validar extensão
        allowed_extensions = {'.xlsx', '.docx', '.pptx'}
        self.validator.validate_file_extension(str(safe_path), allowed_extensions)
        
        # Executar operação com caminho seguro
        if operation == "read":
            return self.excel_tool.read_excel(str(safe_path))
        # ... outras operações
```

## Funções Auxiliares

Para uso rápido, o módulo fornece funções auxiliares:

```python
from src.security_validator import (
    validate_file_path,
    validate_operation,
    sanitize_filename
)

# Validar caminho
safe_path = validate_file_path("test.xlsx", "/home/user/documents")

# Validar operação
validate_operation("read")

# Sanitizar nome
clean_name = sanitize_filename("test@file.xlsx")
```

## Tratamento de Erros

Todas as validações lançam `ValidationError` quando detectam problemas:

```python
from src.exceptions import ValidationError

try:
    validator.validate_file_path("../../../etc/passwd")
except ValidationError as e:
    logger.error(f"Validação de segurança falhou: {e}")
    # Informar usuário sobre tentativa de acesso não autorizado
```

## Boas Práticas

1. **Sempre valide antes de executar**: Valide caminhos e operações antes de qualquer operação de arquivo
2. **Use caminhos resolvidos**: Use o `Path` retornado por `validate_file_path()` para operações
3. **Sanitize user input**: Sempre sanitize nomes de arquivo fornecidos pelo usuário
4. **Log security events**: Registre tentativas de acesso não autorizado para auditoria
5. **Fail securely**: Em caso de dúvida, rejeite a operação

## Exemplos de Ataques Prevenidos

### Path Traversal
```python
# Tentativas bloqueadas:
validator.validate_file_path("../../../etc/passwd")
validator.validate_file_path("..\\..\\..\\windows\\system32")
validator.validate_file_path("/etc/passwd")
validator.validate_file_path("C:\\Windows\\System32\\config")
```

### Nomes de Arquivo Maliciosos
```python
# Sanitizados automaticamente:
sanitize_filename("test/../../evil.xlsx")  # -> "testevil.xlsx"
sanitize_filename("test\\..\\..\\evil.xlsx")  # -> "testevil.xlsx"
sanitize_filename("<script>alert('xss')</script>.xlsx")  # -> "scriptalertxssscript.xlsx"
```

### Operações Não Autorizadas
```python
# Bloqueadas:
validator.validate_operation("delete")  # ValidationError
validator.validate_operation("execute")  # ValidationError
validator.validate_operation("chmod")  # ValidationError
```
