"""Teste para capturar a resposta exata do Gemini."""

import json
from src.factory import create_agent

# Criar agent
agent = create_agent()

# Testar com o prompt problemático
print("="*80)
print("TESTE 1: 'analise o conteudo e crie um grafico de pizza'")
print("="*80)

response = agent.process_user_request(
    "analise o conteudo e crie um grafico de pizza",
    selected_files=["C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx"]
)

print(f"\nSucesso: {response.success}")
print(f"Mensagem: {response.message}")
print(f"Arquivos modificados: {response.files_modified}")

if response.batch_result:
    print(f"\nBatch result:")
    print(f"  Total ações: {response.batch_result.total_actions}")
    print(f"  Sucesso: {response.batch_result.successful_actions}")
    print(f"  Falhas: {response.batch_result.failed_actions}")
    
    for i, action_result in enumerate(response.batch_result.action_results, 1):
        print(f"\n  Ação {i}:")
        print(f"    Sucesso: {action_result.success}")
        print(f"    Mensagem: {action_result.message}")

print("\n" + "="*80)
print("TESTE 2: 'crie um grafico de pizza mostrando vendas por produto, use os dados da coluna C para categorias e coluna F para valores, posicione em H2'")
print("="*80)

response2 = agent.process_user_request(
    "crie um grafico de pizza mostrando vendas por produto, use os dados da coluna C para categorias e coluna F para valores, posicione em H2",
    selected_files=["C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx"]
)

print(f"\nSucesso: {response2.success}")
print(f"Mensagem: {response2.message}")

print("\n" + "="*80)
print("ANÁLISE:")
print("="*80)
print("Se o TESTE 1 falhar mas o TESTE 2 funcionar, significa que o Gemini")
print("precisa de instruções MUITO EXPLÍCITAS sobre quais colunas usar.")
print("\nIsso indica que o Gemini não está 'analisando' os dados, apenas")
print("seguindo templates genéricos.")
