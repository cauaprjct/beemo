from src.excel_tool import ExcelTool

tool = ExcelTool()
data = tool.read_excel('C:/Users/ngb/Documents/GeminiOfficeFiles/vendas_2025.xlsx')

sheet_name = list(data['sheets'].keys())[0]
rows = data['sheets'][sheet_name]

print('=' * 80)
print('VERIFICAÇÃO DE MODIFICAÇÃO - vendas_2025.xlsx')
print('=' * 80)

print('\n📍 CÉLULA B2 (Data da primeira venda):')
print(f'   Linha 2 completa: {rows[1]}')
print(f'   Valor da célula B2: {rows[1][1]}')

print('\n📊 CONTEXTO (primeiras 5 linhas):')
for i, row in enumerate(rows[:5], start=1):
    if i == 1:
        print(f'   Linha {i} (Cabeçalho): {row}')
    else:
        print(f'   Linha {i}: ID={row[0]}, Data={row[1]}, Produto={row[2]}')

print('\n' + '=' * 80)
print('RESULTADO:')
print('=' * 80)

valor_b2 = rows[1][1]
print(f'\n✅ Valor atual em B2: {valor_b2}')
print(f'📅 Valor esperado: 2025-01-15')

if str(valor_b2) == '2025-01-15':
    print('\n🎉 SUCESSO! A modificação foi aplicada corretamente!')
elif str(valor_b2) == '2025-01-01':
    print('\n❌ FALHA! O valor ainda é o original (2025-01-01)')
else:
    print(f'\n⚠️ ATENÇÃO! O valor é diferente do esperado: {valor_b2}')
