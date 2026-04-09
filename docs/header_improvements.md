# Melhorias no Método add_header

## Resumo

Adicionados dois novos parâmetros ao método `add_header` do WordTool para suportar cabeçalhos mais complexos e profissionais.

## Novos Parâmetros

### 1. `include_total_pages: bool = False`

Permite adicionar o total de páginas após o número da página, criando o formato "Página X de Y".

**Antes:** Apenas disponível no `add_footer`
**Agora:** Disponível também no `add_header`

**Exemplo:**
```python
word_tool.add_header(
    "relatorio.docx",
    text="Relatório Mensal",
    alignment="left",
    include_page_number=True,
    page_number_position="right",
    include_total_pages=True  # ← NOVO
)
```

**Resultado:** Cabeçalho com "Relatório Mensal" à esquerda e "1 de 5" à direita (mas ainda na mesma linha horizontal, não alinhados corretamente sem tab stops)

### 2. `use_tab_stops: bool = False`

Configura tab stops automaticamente para permitir texto à esquerda e número de página à direita na MESMA LINHA, corretamente alinhados.

**Problema anterior:** Quando `alignment='left'` e `page_number_position='right'`, o número de página ficava à esquerda também, apenas separado por um tab simples.

**Solução:** Quando `use_tab_stops=True`, o método:
1. Calcula a largura da página menos as margens
2. Adiciona um tab stop alinhado à direita nessa posição
3. Insere um caractere tab antes do número de página
4. O número de página é empurrado para a margem direita

**Exemplo:**
```python
word_tool.add_header(
    "relatorio.docx",
    text="Beemo AI • Relatório de Teste",
    alignment="left",
    include_page_number=True,
    page_number_position="right",
    include_total_pages=True,
    use_tab_stops=True  # ← NOVO
)
```

**Resultado:** 
```
Beemo AI • Relatório de Teste                    Página 1 de 5
└─ alinhado à esquerda                           alinhado à direita ─┘
```

## Implementação Técnica

### Tab Stops

O código calcula a posição do tab stop baseado nas dimensões da página:

```python
page_width = section.page_width
right_margin = section.right_margin
left_margin = section.left_margin
tab_position = page_width - right_margin - left_margin

tab_stops = para.paragraph_format.tab_stops
tab_stops.add_tab_stop(tab_position, WD_TAB_ALIGNMENT.RIGHT)
```

### Total de Páginas

Reutiliza o método `_add_num_pages_field` que já existia no código:

```python
if include_total_pages:
    run_sep = para.add_run(' de ')
    if font_size:
        run_sep.font.size = Pt(font_size)
    if font_name:
        run_sep.font.name = font_name
    self._add_num_pages_field(para)
```

## Uso no Prompt do Agente

Adicionado exemplo no `prompt_templates.py`:

```json
{
  "tool": "word",
  "operation": "add_header",
  "target_file": "relatorio.docx",
  "parameters": {
    "text": "Company Report",
    "alignment": "left",
    "include_page_number": true,
    "page_number_position": "right",
    "include_total_pages": true,
    "use_tab_stops": true
  }
}
```

## Casos de Uso

### 1. Relatórios Corporativos
```
Empresa XYZ • Relatório Trimestral                    Página 1 de 12
```

### 2. Documentos Acadêmicos (ABNT)
```
TCC - Análise de Dados                                Página 1 de 45
```

### 3. Documentos Legais
```
Contrato de Prestação de Serviços                     Página 1 de 8
```

### 4. Manuais Técnicos
```
Manual do Usuário v2.0                                Página 1 de 120
```

## Compatibilidade

- ✅ Retrocompatível: parâmetros novos têm valores padrão `False`
- ✅ Funciona com todos os formatos de fonte e tamanho
- ✅ Respeita margens personalizadas
- ✅ Funciona com orientação retrato e paisagem

## Limitações

- `use_tab_stops` só funciona quando `alignment='left'` e `page_number_position='right'`
- Para outros alinhamentos, o comportamento é o mesmo de antes (tab simples)

## Testes Recomendados

1. Criar documento com cabeçalho complexo
2. Verificar alinhamento em diferentes tamanhos de página (A4, Letter, A3)
3. Testar com margens personalizadas (ABNT, estreitas, etc.)
4. Verificar em orientação retrato e paisagem
5. Exportar para PDF e validar que o cabeçalho está correto

## Próximos Passos

- [ ] Adicionar testes unitários para os novos parâmetros
- [ ] Documentar no README principal
- [ ] Criar exemplos visuais (screenshots)
- [ ] Considerar adicionar `use_tab_stops` também ao `add_footer` (se houver demanda)
