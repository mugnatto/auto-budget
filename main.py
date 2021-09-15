import pandas as pd

print("======================================================================================================")
print("Este programa carrega um arquivo em XLSX contendo duas planilhas")
print('Ambas planilhas devem estar no mesmo arquivo excel e mesmo diretório do programa para funcionar.')
print('Formatar ambas planilhas para que tenham as colunas: ')
print("Cliente 'Descr. Produto','Quantidade'")
print("Fabricante 'Cód.','Descr. Produto','FINAL','Quantidade'")
print("****")
print("MODO DE USO")
print("Criar uma tabela execel e após nomeá-la deverá inserir quando solicitado o nome exatamente como digitou")
print("A primeira planilha deve ser a do cliente, e a segunda do fabricante. NÃO INVERTER!")
print('No fim, será salvo o resultado em um arquivo chamado Orçamento NOME DO CLIENTE.xlsx')
print("======================================================================================================")

# planilhas
empresa = input("Digite o nome do cliente? ")
file = input('Digite o nome da planliha do pedido+tabela sem a extensão (.xlsx, .xls): ')  # substituir file1
# orcamento = '' # planilha que sera inserida os dados processados
arquivo = file + ".xlsx"
print('CARREGANDO DADOS...')

print('*****************')
print('PEDIDO DO CLIENTE')
print('*****************')

# Essa parte do código LÊ o arquivo do pedido

data1 = pd.read_excel(arquivo, sheet_name=0)
pedido = data1[["Descr. Produto", 'Valor Unitario', 'Quantidade']]
print(pedido)

# data1.count(axis='columns')

print('*****************')
print('CARREGANDO A TABELA DO FABRICANTE')
print('*****************')

# Essa parte do código LÊ o arquivo da fábrica
data2 = pd.read_excel(arquivo, sheet_name=1)
itens_do_fabricante = data2[['Cód.',"Descr. Produto",'Valor Unitario','Quantidade']]
fabricante = itens_do_fabricante.to_dict('index')
print(itens_do_fabricante)


comum = pd.merge(pedido, itens_do_fabricante, on=['Descr. Produto'])
print(comum)
comum2 = pd.merge(comum, pedido, on=['Descr. Produto'])
print(comum2)
comum2['Valor Final'] = comum2['Valor Unitario'] * comum2['Quantidade']
comum2['Valor Total'] =

df = comum2[['Cód.',"Descr. Produto",'Valor Unitario','Quantidade', 'Valor Final']]
df.to_excel('Orçamento %s.xlsx' % empresa)
print('==================================================')
print('Arquivo salvo como: Orçamento %s.xlsx' % empresa)
