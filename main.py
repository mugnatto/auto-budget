import pandas as pd

print("======================================================================================================")
print("Este programa carrega um arquivo em XLSX contendo duas planilhas")
print('Ambas planilhas devem estar no mesmo arquivo excel e mesmo diretório do programa para funcionar.')
print("****")
print("MODO DE USO")
print("Criar uma tabela execel e após nomea-la, alterar a linha 27 do código para o nome utilizado")
print("A primeira planilha deve ser a do cliente, e a segunda do fabricante. NÃO INVERTER!")
print('No fim salva o resultado em um arquivo chamado Orçamento.xlsx')
print("======================================================================================================")

# planilhas
empresa = input("Digite o nome do cliente? ")
file = input('digite o nome da planliha do pedido+tabela sem a extensão (.xlsx, .xls): ')  # substituir file1
# orcamento = '' # planilha que sera inserida os dados processados
arquivo = file + ".xlsx"
print('CARREGANDO DADOS...')

print('*****************')
print('PEDIDO DO CLIENTE')
print('*****************')

# Essa parte do código LÊ o arquivo do pedido

data1 = pd.read_excel(arquivo, sheet_name=0)
pedido = data1[["Descr. Produto", 'Quantidade']]
produtos_pedidos = data1["Descr. Produto"]
quantidade = data1["Quantidade"]
print(pedido)
print(pedido.shape)
# data1.count(axis='columns')

print('*****************')
print('CARREGANDO A TABELA DO FABRICANTE')
print('*****************')

# Essa parte do código LÊ o arquivo da fábrica
data2 = pd.read_excel(arquivo, sheet_name=1)
itens_do_fabricante = data2[['Cód.', "Descr. Produto", 'FINAL']]
valor_unitario = data2['FINAL']
fabricante = itens_do_fabricante.to_dict('index')

# Transforma a planilha do fabricante em lista
for produtos_pedidos in fabricante.items():
    (indice, *resultados) = produtos_pedidos  # faz o unpack da tupla para retornar apenas o dict
    print(resultados)
"""
#Essa parte faz o parsing
busca_itens = data1["Descr. Produto"]
parsing_itens = resultados.filter=[busca_itens]
print(parsing_itens)

parsing = fabricante.compare(resultados, keep_shape=True)
print(parsing)
"""

"""
# Essa parte cria o arquivo final com o orçamento resultante
item = data1["Descr. Produto"] # primeira coluna com o nome do item
["FINAL"] # segunda coluna com valor unitário referente ao produto!
final = (valor_do_produto) * (quantidade)
orcamento = pd.DataFrame({ 'Item' : item, 'Quantidade' : quantidade, 'Valor Unitário': valor_unitario, 'Final': final})
orcamento.to_excel('Orçamento %s.xlsx' % empresa)
print('==================================================')
print('Arquivo salvo como: Orçamento %s.xlsx' % empresa)
"""