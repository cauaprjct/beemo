"""Script para limpar cache e testar as correções."""

import os
import shutil
from pathlib import Path

print("="*80)
print("LIMPANDO CACHE E TESTANDO CORREÇÕES")
print("="*80)

# 1. Limpar cache
cache_dir = Path("C:/Users/ngb/Documents/GeminiOfficeFiles/.cache")
if cache_dir.exists():
    try:
        shutil.rmtree(cache_dir)
        print(f"\n✅ Cache removido: {cache_dir}")
    except Exception as e:
        print(f"\n⚠️  Erro ao remover cache: {e}")
else:
    print(f"\n⚠️  Cache não encontrado em: {cache_dir}")

# 2. Verificar se as mudanças foram aplicadas
print("\n" + "="*80)
print("VERIFICANDO MUDANÇAS NO PROMPT")
print("="*80)

with open("src/prompt_templates.py", "r", encoding="utf-8") as f:
    content = f.read()
    
    checks = [
        ("CRITICAL RULE: You MUST ALWAYS respond with valid JSON", "Instrução crítica no início"),
        ("CRITICAL - YOUR RESPONSE FORMAT:", "Instrução explícita no final"),
        ("DO NOT respond with free text", "Proibição de texto livre"),
        ("CRITICAL - ANALYZING DATA BEFORE CREATING CHARTS", "Instrução de análise de dados"),
    ]
    
    all_ok = True
    for check_text, description in checks:
        if check_text in content:
            print(f"✅ {description}")
        else:
            print(f"❌ {description} - NÃO ENCONTRADA!")
            all_ok = False
    
    if all_ok:
        print("\n✅ Todas as mudanças estão presentes no código")
    else:
        print("\n❌ Algumas mudanças estão faltando!")

print("\n" + "="*80)
print("PRÓXIMOS PASSOS")
print("="*80)
print("""
1. REINICIE O STREAMLIT:
   - Pressione Ctrl+C no terminal onde está rodando
   - Execute: streamlit run app.py

2. TESTE COM COMANDOS SIMPLES:
   - "crie um grafico de pizza"
   - "analise o conteudo e crie um grafico"
   - "remova todos os graficos e crie um novo"

3. VERIFIQUE O LOG:
   - Deve mostrar "Successfully parsed response as JSON"
   - NÃO deve mostrar "JSON parsing failed"
   - NÃO deve mostrar "I need more information"

4. SE AINDA FALHAR:
   - Desabilite o cache em .env: CACHE_ENABLED=false
   - Tente modelo diferente: MODEL_NAME=gemini-2.5-flash
   - Verifique se o Streamlit realmente reiniciou (veja timestamp no log)
""")
