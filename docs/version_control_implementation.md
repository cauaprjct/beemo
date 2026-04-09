# Implementação de Versionamento e Undo/Redo

## Visão Geral

Sistema de versionamento automático que permite desfazer e refazer operações em arquivos Office, mantendo um histórico completo de modificações.

## Arquitetura

### Componentes Principais

1. **VersionManager** (`src/version_manager.py`)
   - Gerencia backups e histórico de versões
   - Armazena versões em `.versions/` dentro do ROOT_PATH
   - Mantém metadata em JSON para rastreamento

2. **Integração com Agent** (`src/agent.py`)
   - Cria backups automaticamente antes de modificações
   - Salva estado pós-operação para suportar redo
   - Expõe métodos `undo_file()` e `redo_file()`

3. **Interface Streamlit** (`app.py`)
   - Seção "Histórico de Versões" na sidebar
   - Botões de undo/redo por arquivo
   - Visualização de histórico com timestamps

### Estrutura de Dados

```python
Version:
  - version_id: UUID único
  - timestamp: ISO 8601
  - operation: tipo de operação (create, update, add)
  - backup_path: caminho do backup (estado ANTES)
  - after_path: caminho do estado APÓS (para redo)
  - user_prompt: prompt que gerou a operação
  - file_size: tamanho do arquivo

FileHistory:
  - file_path: caminho relativo do arquivo
  - versions: lista de Version
  - current_index: índice da versão atual (-1 = estado mais recente)
```

### Armazenamento

```
ROOT_PATH/
├── .versions/
│   ├── metadata.json          # Metadados de todas as versões
│   ├── file1_uuid_before.xlsx # Backup antes da operação
│   ├── file1_uuid_after.xlsx  # Estado após a operação
│   └── ...
└── [arquivos do usuário]
```

## Fluxo de Operação

### Criação de Backup

1. Usuário solicita operação (ex: "Atualizar célula A1")
2. Agent valida operação
3. VersionManager cria backup do estado atual (se arquivo existe)
4. Agent executa operação
5. VersionManager salva estado pós-operação
6. Metadata é atualizado e persistido

### Undo

1. Usuário clica em "Desfazer"
2. VersionManager verifica se há versões disponíveis
3. Restaura arquivo do backup (estado ANTES da operação)
4. Decrementa current_index
5. Metadata é atualizado

### Redo

1. Usuário clica em "Refazer"
2. VersionManager verifica se há versões futuras
3. Restaura arquivo do estado APÓS a operação
4. Incrementa current_index
5. Metadata é atualizado

### Descarte de Versões Futuras

Quando uma nova operação é realizada após um undo:
1. Todas as versões "futuras" são descartadas
2. Arquivos de backup são removidos
3. Nova versão é adicionada ao histórico

## Configuração

### Variável de Ambiente

```env
MAX_VERSIONS=10  # Número máximo de versões por arquivo
```

### Limites

- Padrão: 10 versões por arquivo
- Versões mais antigas são automaticamente removidas
- Arquivos de backup são deletados quando versões são descartadas

## Uso

### Programático

```python
from src.factory import create_agent

agent = create_agent()

# Operação normal (backup automático)
response = agent.process_user_request("Atualizar célula A1 para 100")

# Desfazer
undo_response = agent.undo_file("/path/to/file.xlsx")

# Refazer
redo_response = agent.redo_file("/path/to/file.xlsx")

# Verificar disponibilidade
can_undo = agent.version_manager.can_undo("/path/to/file.xlsx")
can_redo = agent.version_manager.can_redo("/path/to/file.xlsx")

# Obter histórico
history = agent.version_manager.get_history("/path/to/file.xlsx")
```

### Interface Streamlit

1. Realize operações normalmente
2. Na sidebar, veja "Histórico de Versões"
3. Expanda um arquivo para ver opções
4. Clique em "↩️ Desfazer" ou "↪️ Refazer"
5. Veja o histórico completo com timestamps

## Testes

```bash
# Testes do VersionManager
pytest tests/test_version_manager.py -v

# Testes de integração
pytest tests/test_agent.py tests/test_factory.py -v
```

### Cobertura de Testes

- ✅ Inicialização do VersionManager
- ✅ Criação de backups
- ✅ Operações de undo/redo
- ✅ Verificação de disponibilidade (can_undo/can_redo)
- ✅ Múltiplas versões
- ✅ Limpeza de versões antigas
- ✅ Descarte de versões futuras
- ✅ Salvamento de estado pós-operação
- ✅ Integração com Agent

## Limitações

1. **Operações de leitura**: Não criam backups (apenas modificações)
2. **Arquivos novos**: Primeira operação (create) não tem estado anterior
3. **Espaço em disco**: Versões consomem espaço proporcional ao tamanho dos arquivos
4. **Concorrência**: Não há proteção contra modificações externas simultâneas
5. **Compressão**: Backups não são comprimidos (pode ser adicionado futuramente)

## Melhorias Futuras

- [ ] Compressão de backups para economizar espaço
- [ ] Detecção de modificações externas
- [ ] Limite de tamanho total para versões
- [ ] Exportar/importar histórico de versões
- [ ] Comparação visual entre versões (diff)
- [ ] Restauração de versões específicas (não apenas undo/redo sequencial)
- [ ] Proteção contra corrupção de metadata.json
- [ ] Suporte a tags/labels para versões importantes

## Arquivos Modificados

### Novos Arquivos
- `src/version_manager.py` - Implementação do gerenciador de versões
- `tests/test_version_manager.py` - Testes unitários
- `docs/version_control_implementation.md` - Esta documentação

### Arquivos Modificados
- `config/__init__.py` - Adicionado `max_versions` property
- `src/factory.py` - Inicialização do VersionManager
- `src/agent.py` - Integração com versionamento + métodos undo/redo
- `app.py` - UI de histórico de versões na sidebar
- `.env.example` - Adicionado MAX_VERSIONS
- `README.md` - Documentação da funcionalidade

## Conclusão

O sistema de versionamento está totalmente funcional e integrado ao Gemini Office Agent. Todos os testes passam e a funcionalidade está disponível tanto programaticamente quanto através da interface Streamlit.
