import tkinter as tk
import pandas as pd
from datetime import datetime
from tkcalendar import Calendar

# Criando uma função para escolher a data no calendário
def selecionar_data():
    def obter_data_selecionada():
        data = cal.get_date()
        data_inicial.set(data)
        top.destroy()

    top = tk.Toplevel(root)
    cal = Calendar(top, selectmode='day', day=datetime.now().day, month=datetime.now().month, year=datetime.now().year)
    cal.pack()
    tk.Button(top, text='Selecionar Data', command=obter_data_selecionada).pack() #Ao clicar no botão de selecionar data, a janela é fechada e a data selecionada é armazenada na variável data_inicial

# Criando uma função para gerar a planilha
def gerar_planilha():
    # Carregando a planilha. Coloque o arquivo na mesma pasta do script.
    df = pd.read_excel('Metas Consultores.xlsx', sheet_name=1) # 'Sheet_name' refere em qual aba o arquivo está. No caso, a segunda aba da planilha.

    # Substituindo valores não-numéricos do DataFrame
    pd.set_option('future.no_silent_downcasting', True)
    df.replace(['FÉRIAS', 'DESLIGADA', 'LIC. MATER', 'TREINAMENTO'], [0, 0, 0, 0], inplace=True)

    # Selecionando as primeiras 3 colunas
    primeiras_3_colunas = df.iloc[:, :3]
    
    # Gerando as datas repetidas
    data_inicial_datetime = data_inicial.get()
    datas_repetidas = pd.date_range(start=data_inicial_datetime, periods=int(dias_do_mes.get()))  # O período é definido de acordo com a opção selecionada
    datas_repetidas = pd.Series(datas_repetidas).dt.strftime('%d/%m/%Y') # Transformar o DatetimeIndex em uma Series e formatar as datas para excluir as informações de hora
    datas_repetidas = pd.Series(datas_repetidas) # Transformar o DatetimeIndex em uma Series

    # Selecionando as demais colunas
    colunas_restantes = df.iloc[:, 3:]

    # Criando uma lista vazia para armazenar as linhas empilhadas
    linhas_empilhadas = []

    # Iterando sobre cada linha do dataframe
    for index, row in df.iterrows():
        # Repetindo as primeiras 3 colunas (Alterar a quantidade de acordo com os dias do mês)
        primeiras_3_colunas_repetidas = pd.concat([primeiras_3_colunas.iloc[[index]]]*int(dias_do_mes.get()), ignore_index=True) # A quantidade de colunas é alterada de acordo com a opção selecionada
        primeiras_3_colunas_repetidas['Dia'] = datas_repetidas.tolist() # Adicionando a coluna de dias
        linhas_empilhadas.extend(primeiras_3_colunas_repetidas.values.tolist()) # Adicionando as primeiras 3 colunas repetidas à lista de linhas empilhadas

    # Criando um dataframe com as linhas empilhadas
    df_empilhado = pd.DataFrame(linhas_empilhadas, columns=list(primeiras_3_colunas.columns) + ['Dia'])

    # Empilhando as demais colunas em uma única coluna
    demais_colunas_empilhadas = colunas_restantes.stack().reset_index(drop=True)

    # Adicionando a coluna 'Meta' ao dataframe
    df_empilhado['Meta'] = demais_colunas_empilhadas

    # Adicionando a coluna 'Cargo' com o valor 'Vendedor' em todas as linhas
    df_empilhado['Cargo'] = 'Vendedor'

    # Reordenando as colunas
    colunas = ['Cód.', 'Loja', 'Consultores', 'Meta', 'Dia', 'Cargo']
    resultado_final = df_empilhado[colunas]

    # Escrevendo o DataFrame em um arquivo Excel e definindo a largura das colunas
    with pd.ExcelWriter('planilha_para_BI.xlsx', engine='xlsxwriter') as writer:
        # Escreve o DataFrame empilhado na aba 'Metas Consultores'
        resultado_final.to_excel(writer, index=False, sheet_name='Metas Consultores')
        worksheet_empilhado = writer.sheets['Metas Consultores']  # Acessa a aba 'Metas Consultores'

        # Define a largura das colunas do DataFrame empilhado
        for i, col in enumerate(df_empilhado.columns):
            column_len = max(df_empilhado[col].astype(str).map(len).max(), len(str(col))) + 2  # Adiciona uma margem de 2 caracteres
            worksheet_empilhado.set_column(i, i, column_len)
    root.destroy() # Fecha a janela de interação após clicar em "Gerar Planilha"
    pass 

# Criando uma janela para interação
root = tk.Tk()
root.title('Gerador de Planilha')

# cria a variável para alterar a quantidade de dias no mês, dentro das variáveis 'datas_repetidas' e 'primeiras_3_colunas_repetidas'
dias_do_mes = tk.IntVar()

tk.Label(root, text='Quantos dias tem  mês?').pack()

# Criando 3 opções de botão para que o usuario selecione a quantidade de dias no mes correspondente
tk.Radiobutton(root, text='28', variable=dias_do_mes, value=28).pack()
tk.Radiobutton(root, text='30', variable=dias_do_mes, value=30).pack()
tk.Radiobutton(root, text='31', variable=dias_do_mes, value=31).pack()

# Criando uma variável para armazenar a data inicial
data_inicial = tk.StringVar()
data_inicial.set(datetime.now().strftime('%d/%m/%Y'))

tk.Button(root, text='Selecionar Data', command=selecionar_data).pack()  # Botão para selecionar a data inicial
tk.Button(root, text='Gerar Planilha', command=gerar_planilha).pack() # Botão para gerar a planilha

root.mainloop()