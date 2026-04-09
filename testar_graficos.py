"""
Script de teste para as funcionalidades de gerenciamento de gráficos.

Execute este script para testar as três implementações:
1. Validação de posição de gráficos
2. Listagem de gráficos
3. Deleção de gráficos

Uso:
    python testar_graficos.py
"""

from src.excel_tool import ExcelTool
from pathlib import Path

# Configuração
ARQUIVO_TESTE = r"C:\Users\ngb\Documents\GeminiOfficeFiles\vendas_2025.xlsx"
SHEET = "Sheet1"

def print_separator(title):
    """Imprime um separador visual."""
    print("\n" + "="*70)
    print(f"  {title}")
    print("="*70 + "\n")

def test_list_charts(tool):
    """Testa a listagem de gráficos."""
    print_separator("TESTE 1: Listar Gráficos")
    
    try:
        result = tool.list_charts(ARQUIVO_TESTE, SHEET)
        
        print(f"✅ Total de gráficos encontrados: {result['total_count']}")
        print(f"✅ Sheets analisadas: {', '.join(result['sheets_analyzed'])}\n")
        
        if result['total_count'] > 0:
            print("Gráficos encontrados:")
            for i, chart in enumerate(result['charts'], 1):
                print(f"  {i}. {chart['title']} ({chart['type']})")
                print(f"     Posição: {chart['position']}, Índice: {chart['index']}")
        else:
            print("⚠️  Nenhum gráfico encontrado no arquivo.")
        
        return result
        
    except Exception as e:
        print(f"❌ Erro ao listar gráficos: {e}")
        return None

def test_add_chart_validation(tool):
    """Testa a validação de posição ao adicionar gráficos."""
    print_separator("TESTE 2: Validação de Posição")
    
    # Primeiro, criar um gráfico em H2
    print("Tentando criar gráfico na posição H2...")
    chart_config_1 = {
        'type': 'column',
        'title': 'Teste - Vendas por Produto',
        'categories': 'C2:C10',
        'values': 'F2:F10',
        'position': 'H2'
    }
    
    try:
        tool.add_chart(ARQUIVO_TESTE, SHEET, chart_config_1)
        print("✅ Gráfico criado com sucesso na posição H2\n")
    except Exception as e:
        print(f"⚠️  Gráfico já existe ou erro: {e}\n")
    
    # Tentar criar outro gráfico na mesma posição (deve falhar)
    print("Tentando criar outro gráfico na MESMA posição H2...")
    chart_config_2 = {
        'type': 'pie',
        'title': 'Teste - Distribuição',
        'categories': 'C2:C10',
        'values': 'F2:F10',
        'position': 'H2'
    }
    
    try:
        tool.add_chart(ARQUIVO_TESTE, SHEET, chart_config_2)
        print("❌ ERRO: Gráfico foi criado (não deveria!)")
    except ValueError as e:
        print(f"✅ Validação funcionou! Erro esperado: {str(e)[:100]}...")
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")
    
    # Criar em posição diferente (deve funcionar)
    print("\nTentando criar gráfico em posição DIFERENTE (K2)...")
    chart_config_3 = {
        'type': 'pie',
        'title': 'Teste - Distribuição',
        'categories': 'C2:C10',
        'values': 'F2:F10',
        'position': 'K2'
    }
    
    try:
        tool.add_chart(ARQUIVO_TESTE, SHEET, chart_config_3)
        print("✅ Gráfico criado com sucesso na posição K2")
    except Exception as e:
        print(f"⚠️  Gráfico já existe ou erro: {e}")

def test_delete_chart(tool, charts_info):
    """Testa a deleção de gráficos."""
    print_separator("TESTE 3: Deletar Gráficos")
    
    if not charts_info or charts_info['total_count'] == 0:
        print("⚠️  Nenhum gráfico para deletar. Execute o teste 2 primeiro.")
        return
    
    # Teste 3.1: Deletar por índice
    print("Teste 3.1: Deletar gráfico por ÍNDICE (índice 0)...")
    try:
        tool.delete_chart(ARQUIVO_TESTE, SHEET, 0)
        print("✅ Gráfico deletado com sucesso por índice\n")
    except Exception as e:
        print(f"❌ Erro ao deletar por índice: {e}\n")
    
    # Verificar quantos gráficos restam
    result = tool.list_charts(ARQUIVO_TESTE, SHEET)
    print(f"Gráficos restantes: {result['total_count']}\n")
    
    if result['total_count'] > 0:
        # Teste 3.2: Deletar por título
        chart_title = result['charts'][0]['title']
        print(f"Teste 3.2: Deletar gráfico por TÍTULO ('{chart_title}')...")
        try:
            tool.delete_chart(ARQUIVO_TESTE, SHEET, chart_title)
            print("✅ Gráfico deletado com sucesso por título\n")
        except Exception as e:
            print(f"❌ Erro ao deletar por título: {e}\n")
    
    # Teste 3.3: Tentar deletar gráfico inexistente (deve falhar)
    print("Teste 3.3: Tentar deletar gráfico INEXISTENTE...")
    try:
        tool.delete_chart(ARQUIVO_TESTE, SHEET, "Gráfico Inexistente")
        print("❌ ERRO: Deveria ter falhado!")
    except ValueError as e:
        print(f"✅ Validação funcionou! Erro esperado: {str(e)[:100]}...")
    except Exception as e:
        print(f"❌ Erro inesperado: {e}")

def test_workflow_complete(tool):
    """Testa o workflow completo: listar → deletar → criar."""
    print_separator("TESTE 4: Workflow Completo")
    
    print("Passo 1: Listar gráficos existentes...")
    result = tool.list_charts(ARQUIVO_TESTE, SHEET)
    print(f"✅ Encontrados {result['total_count']} gráficos\n")
    
    print("Passo 2: Deletar todos os gráficos...")
    deleted_count = 0
    while True:
        result = tool.list_charts(ARQUIVO_TESTE, SHEET)
        if result['total_count'] == 0:
            break
        try:
            tool.delete_chart(ARQUIVO_TESTE, SHEET, 0)
            deleted_count += 1
        except Exception as e:
            print(f"❌ Erro ao deletar: {e}")
            break
    print(f"✅ Deletados {deleted_count} gráficos\n")
    
    print("Passo 3: Criar novos gráficos...")
    charts_to_create = [
        {
            'type': 'column',
            'title': 'Vendas por Produto',
            'categories': 'C2:C10',
            'values': 'F2:F10',
            'position': 'H2'
        },
        {
            'type': 'pie',
            'title': 'Distribuição de Vendas',
            'categories': 'C2:C10',
            'values': 'F2:F10',
            'position': 'K2'
        }
    ]
    
    created_count = 0
    for chart_config in charts_to_create:
        try:
            tool.add_chart(ARQUIVO_TESTE, SHEET, chart_config)
            created_count += 1
            print(f"  ✅ Criado: {chart_config['title']} em {chart_config['position']}")
        except Exception as e:
            print(f"  ❌ Erro ao criar {chart_config['title']}: {e}")
    
    print(f"\n✅ Criados {created_count} novos gráficos\n")
    
    print("Passo 4: Verificar resultado final...")
    result = tool.list_charts(ARQUIVO_TESTE, SHEET)
    print(f"✅ Total de gráficos no arquivo: {result['total_count']}")
    for chart in result['charts']:
        print(f"  - {chart['title']} ({chart['type']}) em {chart['position']}")

def main():
    """Função principal."""
    print("\n" + "🧪 TESTE DAS FUNCIONALIDADES DE GRÁFICOS ".center(70, "="))
    print(f"\nArquivo de teste: {ARQUIVO_TESTE}")
    print(f"Sheet: {SHEET}\n")
    
    # Verificar se o arquivo existe
    if not Path(ARQUIVO_TESTE).exists():
        print(f"❌ ERRO: Arquivo não encontrado: {ARQUIVO_TESTE}")
        print("\nPor favor, verifique o caminho do arquivo e tente novamente.")
        return
    
    print("⚠️  IMPORTANTE: Feche o Excel se o arquivo estiver aberto!\n")
    input("Pressione ENTER para continuar...")
    
    # Criar instância do ExcelTool
    tool = ExcelTool()
    
    # Executar testes
    try:
        # Teste 1: Listar gráficos
        charts_info = test_list_charts(tool)
        
        # Teste 2: Validação de posição
        test_add_chart_validation(tool)
        
        # Atualizar lista de gráficos
        charts_info = tool.list_charts(ARQUIVO_TESTE, SHEET)
        
        # Teste 3: Deletar gráficos
        test_delete_chart(tool, charts_info)
        
        # Teste 4: Workflow completo
        test_workflow_complete(tool)
        
        print_separator("TESTES CONCLUÍDOS")
        print("✅ Todos os testes foram executados!")
        print("\nVocê pode abrir o arquivo no Excel para verificar visualmente:")
        print(f"  {ARQUIVO_TESTE}\n")
        
    except Exception as e:
        print(f"\n❌ ERRO GERAL: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
