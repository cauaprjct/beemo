"""Teste final para validar as correções implementadas."""

import sys
import json
from src.factory import create_agent

# Configurar encoding para UTF-8
sys.stdout.reconfigure(encoding='utf-8')

print("="*80)
print("TESTE DE VALIDAÇÃO FINAL - Correções de Gráficos")
print("="*80)

agent = create_agent()

# Teste 1: Comando genérico (deve analisar e usar ranges corretos)
print("\n" + "="*80)
print("TESTE 1: Comando genérico")
print("Prompt: 'analise o conteudo e crie um grafico de pizza'")
print("="*80)
print("\nEsperado:")
print("  - 2 ações: read + add_chart")
print("  - categories: C2:C101 (coluna Produto)")
print("  - values: F2:F101 (coluna Total)")
print("  - position: H2 ou similar (longe dos dados)")

try:
    response1 = agent.process_user_request(
        "analise o conteudo e crie um grafico de pizza",
        selected_files=["C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx"]
    )
    
    print(f"\nResultado:")
    print(f"  Sucesso: {response1.success}")
    if response1.batch_result:
        print(f"  Total ações: {response1.batch_result.total_actions}")
        print(f"  Ações bem-sucedidas: {response1.batch_result.successful_actions}")
        
        if response1.batch_result.total_actions == 2:
            print("  ✅ Executou 2 ações (read + add_chart)")
        else:
            print(f"  ❌ Executou {response1.batch_result.total_actions} ações (esperado: 2)")
    
    # Verificar no log se os ranges estão corretos
    with open("logs/agent.log", "r", encoding="utf-8") as f:
        log_content = f.read()
        if "C2:C101" in log_content and "F2:F101" in log_content:
            print("  ✅ Ranges corretos encontrados no log (C2:C101, F2:F101)")
        else:
            print("  ❌ Ranges incorretos ou não encontrados no log")
            
except Exception as e:
    print(f"  ❌ Erro: {e}")

# Teste 2: Comando com múltiplas ações
print("\n" + "="*80)
print("TESTE 2: Múltiplas ações")
print("Prompt: 'remova o grafico atual e crie um novo'")
print("="*80)
print("\nEsperado:")
print("  - 2 ou 3 ações: list_charts/delete_chart + add_chart")
print("  - Não deve executar apenas list_charts e parar")

try:
    response2 = agent.process_user_request(
        "remova o grafico atual e crie um novo",
        selected_files=["C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx"]
    )
    
    print(f"\nResultado:")
    print(f"  Sucesso: {response2.success}")
    if response2.batch_result:
        print(f"  Total ações: {response2.batch_result.total_actions}")
        
        if response2.batch_result.total_actions >= 2:
            print("  ✅ Executou múltiplas ações (não parou no list_charts)")
        else:
            print(f"  ❌ Executou apenas {response2.batch_result.total_actions} ação")
            
except Exception as e:
    print(f"  ❌ Erro: {e}")

# Teste 3: Top 10 produtos
print("\n" + "="*80)
print("TESTE 3: Subset de dados")
print("Prompt: 'crie um grafico de pizza com os top 10 produtos'")
print("="*80)
print("\nEsperado:")
print("  - categories: C2:C11 (primeiros 10 produtos)")
print("  - values: F2:F11 (primeiros 10 totais)")

try:
    response3 = agent.process_user_request(
        "crie um grafico de pizza com os top 10 produtos",
        selected_files=["C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx"]
    )
    
    print(f"\nResultado:")
    print(f"  Sucesso: {response3.success}")
    
    # Verificar no log
    with open("logs/agent.log", "r", encoding="utf-8") as f:
        log_content = f.read()
        if "C2:C11" in log_content and "F2:F11" in log_content:
            print("  ✅ Ranges corretos para top 10 (C2:C11, F2:F11)")
        elif "C2:C101" in log_content:
            print("  ⚠️  Usou todos os dados (C2:C101) em vez de top 10")
        else:
            print("  ❌ Ranges não encontrados ou incorretos")
            
except Exception as e:
    print(f"  ❌ Erro: {e}")

print("\n" + "="*80)
print("RESUMO DA VALIDAÇÃO")
print("="*80)
print("\nSe todos os testes passaram:")
print("  ✅ O Gemini está analisando os dados corretamente")
print("  ✅ O Gemini está usando ranges apropriados")
print("  ✅ O Gemini está executando múltiplas ações quando necessário")
print("\nSe algum teste falhou:")
print("  ❌ Verifique o log em logs/agent.log")
print("  ❌ Verifique se o Streamlit foi reiniciado após as mudanças")
print("  ❌ Considere desabilitar o cache temporariamente")
