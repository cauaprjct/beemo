---
inclusion: auto
---

# Regras de Desenvolvimento

**Todo código entregue deve ser 100% funcional e pronto para produção.**

## Proibições Absolutas

**Nunca use:**
- Placeholders: `TODO`, `FIXME`, `HACK`, `pass`, `NotImplementedError`, `"..."`, `[PLACEHOLDER]`
- Nomes genéricos: `foo`, `bar`, `temp`, `data2`, `x2`
- Código morto: `if False:`, funções vazias sem propósito
- Exceções genéricas: `except:` ou `except Exception: pass`

**Se faltar contexto, pergunte. Nunca assuma.**

## Implementação Completa

Toda função/classe deve incluir:
- Imports corretos
- Type hints (parâmetros e retorno)
- Tratamento de erros específico com ação real
- Validação de entrada
- Logging operacional (não `print`)
- Docstrings (Google/NumPy style)

```python
# ✅ Correto
def process_file(path: str, encoding: str = "utf-8") -> dict[str, Any]:
    """Process file and return parsed data.
    
    Args:
        path: File path to process
        encoding: File encoding (default: utf-8)
        
    Returns:
        Parsed data dictionary
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format is invalid
    """
    try:
        with open(path, encoding=encoding) as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        logger.error("Invalid JSON in %s: %s", path, e)
        raise ValueError(f"Invalid JSON format: {e}") from e

# ❌ Errado
def process_file(path, encoding="utf-8"):
    # TODO: implement this
    pass
```

## Segurança

- **Nunca** hardcode credenciais (use variáveis de ambiente)
- Valide e sanitize toda entrada de usuário
- Valide paths para prevenir path traversal
- Use parâmetros preparados em SQL (nunca f-strings)

## Qualidade

**Nomenclatura:** Nomes descritivos que revelam intenção. Funções fazem uma coisa só.

**Estrutura:** Funções >40 linhas precisam ser quebradas. Não repita código (DRY).

**Comentários:** Explique o **porquê**, não o **quê**. Código auto-explicativo > comentários.

```python
# ✅ Bom: explica decisão não óbvia
# API retorna 429 mesmo em erros de auth, checamos o body
if response.status_code == 429 and "auth" in response.text:
    raise AuthenticationError(...)

# ❌ Ruído
# Incrementa contador
counter += 1
```

## Testes

Escreva testes unitários (pytest) **a menos que explicitamente dispensado**.
- Cubra: caminho feliz, entradas inválidas, edge cases
- Mocks realistas (não esconda comportamento real)

## Dependências

- Verifique se já existe em `requirements.txt`
- Justifique novas dependências
- Prefira stdlib quando possível

## Checklist Pré-Entrega

Antes de entregar código, confirme:

1. ✅ Nenhum placeholder ou TODO
2. ✅ Imports completos
3. ✅ Type hints em todas as funções
4. ✅ Tratamento de erro específico
5. ✅ Nenhuma credencial hardcoded
6. ✅ Docstrings em funções/classes públicas
7. ✅ Código executável sem edição manual

**Se não souber algo, pergunte explicitamente. Não invente.**
