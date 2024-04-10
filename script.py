import pandas as pd

# Carregando a planilha. Coloque o arquivo na mesma pasta do script.
df = pd.read_excel('Metas.xlsx', sheet_name=1) # 'Sheet_name' refere em qual aba o arquivo está. No caso, a segunda aba da planilha.

# Realizando a substituição de valor não-numéricos
df.replace(['FÉRIAS', 'DESLIGADA', 'LIC. MATER', 'TREINAMENTO'], [0, 0, 0, 0], inplace=True)

# Selecionando as primeiras 3 colunas
primeiras_3_colunas = df.iloc[:, :3]

# Gerando as datas repetidas
datas_repetidas = pd.date_range(start='2024-04-01', periods=30)  # ALTERE A DATA INICIAL E O PERÍODO!!!

# Transformar o DatetimeIndex em uma Series
datas_repetidas = pd.Series(datas_repetidas)

# Selecionando as demais colunas
colunas_restantes = df.iloc[:, 3:]

# Criando uma lista vazia para armazenar as linhas empilhadas
linhas_empilhadas = []

# Iterando sobre cada linha do dataframe
for index, row in df.iterrows():
    # Repetindo as primeiras 3 colunas (Alterar a quantidade de acordo com os dias do mês)
    primeiras_3_colunas_repetidas = pd.concat([primeiras_3_colunas.iloc[[index]]]*30, ignore_index=True) # ALTERE OS DIAS DE ACORDO COM O MÊS!!!!!
    # Adicionando a coluna de datas
    primeiras_3_colunas_repetidas['Data'] = datas_repetidas.tolist()
    # Adicionando as primeiras 3 colunas repetidas à lista de linhas empilhadas
    linhas_empilhadas.extend(primeiras_3_colunas_repetidas.values.tolist())

# Criando um dataframe com as linhas empilhadas
df_empilhado = pd.DataFrame(linhas_empilhadas, columns=list(primeiras_3_colunas.columns) + ['Data'])

# Empilhando as demais colunas em uma única coluna
demais_colunas_empilhadas = colunas_restantes.stack().reset_index(drop=True)

# Adicionando a coluna 'Meta' ao dataframe
df_empilhado['Meta'] = demais_colunas_empilhadas

# Adicionando a coluna 'Cargo' com o valor 'Vendedor' em todas as linhas
df_empilhado['Cargo'] = 'Vendedor'

# Reordenando as colunas
colunas = ['Código', 'Loja', 'Consultor', 'Meta', 'Data', 'Cargo']
resultado_final = df_empilhado[colunas]

# Salvando o resultado em uma nova planilha
resultado_final.to_excel('nova_planilha.xlsx', index=False)
