import pandas as pd

# Colunas da tabela
receita = []
dia = []
valor = []

# Path do arquivo Excel
caminho_arquivo = 'realizado_02.xls'

# Lendo arquivo
tabela = pd.read_excel(caminho_arquivo)

for indice, linha in tabela.iterrows():
    nome_primeira_coluna = None
    for nome_coluna, valor_celula in linha.items():
        if isinstance(nome_coluna, str) and (nome_coluna == 'Plano Financeiro - Receitas' or nome_coluna == 'Plano Financeiro - Despesas'):
            nome_primeira_coluna = valor_celula
        else:
            nome_receita = str(nome_primeira_coluna).replace('            ','')
            nome_receita = str(nome_primeira_coluna).replace('    ','')
            nome_receita = nome_receita.strip()
            receita.append(nome_receita)

            adiciona_dia = str(nome_coluna).replace('Dia   ','')
            adiciona_dia = str(adiciona_dia) + '/01/2023'
            adiciona_dia = adiciona_dia.strip()
            
            dia.append(adiciona_dia)
            valor.append(valor_celula)

# Cria um DataFrame pandas com base nas listas
dados = pd.DataFrame({'Receita': receita, 'Dia': dia, 'valor': valor})

# Escreve os dados no arquivo Excel
dados.to_excel('realizado_despesa_2023_02.xlsx', index=False)

print('Terminou!')