"""Script para testar se o prompt atualizado resolve o problema de múltiplas ações."""

from src.prompt_templates import PromptTemplates

# Verificar se a instrução foi adicionada
system_prompt = PromptTemplates.get_system_prompt()

print("=" * 80)
print("VERIFICANDO SE A INSTRUÇÃO SOBRE MÚLTIPLAS AÇÕES FOI ADICIONADA")
print("=" * 80)

if "CRITICAL INSTRUCTION - MULTIPLE ACTIONS" in system_prompt:
    print("✅ Instrução sobre múltiplas ações encontrada!")
    
    # Extrair e mostrar a seção relevante
    start = system_prompt.find("CRITICAL INSTRUCTION")
    end = system_prompt.find("EXAMPLES OF VALID RESPONSES", start)
    
    if end == -1:
        end = start + 500
    
    print("\nSeção adicionada:")
    print("-" * 80)
    print(system_prompt[start:end])
    print("-" * 80)
else:
    print("❌ Instrução NÃO encontrada no prompt!")

# Verificar se o exemplo 10 foi adicionado
if "Example 10 - Analyze and create chart" in system_prompt:
    print("\n✅ Exemplo 10 (análise + gráfico) encontrado!")
else:
    print("\n❌ Exemplo 10 NÃO encontrado!")

print("\n" + "=" * 80)
print("PRÓXIMOS PASSOS:")
print("=" * 80)
print("1. Reinicie o Streamlit (Ctrl+C e depois 'streamlit run app.py')")
print("2. Teste novamente com: 'analise o conteudo e crie um grafico de pizza'")
print("3. O Gemini agora deve executar AMBAS as ações (read + add_chart)")
