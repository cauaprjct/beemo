"""Script para auditar o prompt enviado ao Gemini."""

from src.factory import create_agent
from src.excel_tool import ExcelTool
from src.prompt_templates import PromptTemplates

# Ler o arquivo vendas
excel_tool = ExcelTool()
content = excel_tool.read_excel("C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx")

# Ver quantas linhas existem
for sheet_name, rows in content['sheets'].items():
    print(f"\n{'='*80}")
    print(f"Sheet: {sheet_name}")
    print(f"Total de linhas: {len(rows)}")
    print(f"Primeiras 3 linhas:")
    for i, row in enumerate(rows[:3], 1):
        print(f"  Row {i}: {row}")
    print(f"...")
    print(f"Últimas 2 linhas:")
    for i, row in enumerate(rows[-2:], len(rows)-1):
        print(f"  Row {i}: {row}")

# Formatar como seria enviado ao Gemini
formatted = PromptTemplates.format_file_content(
    "vendas_2025.xlsx",
    "excel",
    content
)

print(f"\n{'='*80}")
print(f"TAMANHO DO CONTEÚDO FORMATADO:")
print(f"  Caracteres: {len(formatted['content'])}")
print(f"  Linhas: {formatted['content'].count(chr(10))}")

# Mostrar primeiras e últimas linhas do conteúdo formatado
lines = formatted['content'].split('\n')
print(f"\n{'='*80}")
print(f"PRIMEIRAS 10 LINHAS DO CONTEÚDO FORMATADO:")
print('\n'.join(lines[:10]))
print(f"\n... ({len(lines) - 20} linhas omitidas) ...\n")
print(f"ÚLTIMAS 10 LINHAS:")
print('\n'.join(lines[-10:]))

# Construir prompt completo
file_contexts = [formatted]
full_prompt = PromptTemplates.build_context_prompt(
    "analise o conteudo e crie um grafico de pizza",
    file_contexts
)

print(f"\n{'='*80}")
print(f"TAMANHO DO PROMPT COMPLETO:")
print(f"  Caracteres: {len(full_prompt)}")
print(f"  Tokens estimados: {len(full_prompt) // 4}")
print(f"  Linhas: {full_prompt.count(chr(10))}")

# Verificar se as instruções CRITICAL estão presentes
if "CRITICAL INSTRUCTION - MULTIPLE ACTIONS" in full_prompt:
    print(f"\n✅ Instrução sobre múltiplas ações encontrada")
else:
    print(f"\n❌ Instrução sobre múltiplas ações NÃO encontrada")

if "CRITICAL - PIE CHART REQUIREMENTS" in full_prompt:
    print(f"✅ Instrução sobre gráficos de pizza encontrada")
else:
    print(f"❌ Instrução sobre gráficos de pizza NÃO encontrada")

# Calcular proporção de instruções vs dados
system_prompt = PromptTemplates.get_system_prompt()
data_size = len(formatted['content'])
instructions_size = len(system_prompt)
total_size = len(full_prompt)

print(f"\n{'='*80}")
print(f"PROPORÇÃO DO PROMPT:")
print(f"  Instruções do sistema: {instructions_size} chars ({instructions_size/total_size*100:.1f}%)")
print(f"  Dados do arquivo: {data_size} chars ({data_size/total_size*100:.1f}%)")
print(f"  Outros (request, etc): {total_size - instructions_size - data_size} chars")
print(f"\n⚠️  Se os dados ocupam >50% do prompt, as instruções podem ser 'perdidas'")
