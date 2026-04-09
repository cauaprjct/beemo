from src.excel_tool import ExcelTool

tool = ExcelTool()
result = tool.list_charts('C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx')

print(f'\n=== ESTADO ATUAL DOS GRAFICOS ===\n')
print(f'Total de graficos: {result["total_count"]}')
print(f'Sheets analisadas: {", ".join(result["sheets_analyzed"])}\n')

if result['total_count'] > 0:
    print('Graficos encontrados:')
    for i, chart in enumerate(result['charts'], 1):
        print(f'  {i}. {chart["title"]} ({chart["type"]})')
        print(f'     Posicao: {chart["position"]}, Indice: {chart["index"]}')
else:
    print('Nenhum grafico encontrado.')
