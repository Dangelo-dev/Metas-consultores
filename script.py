import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
from tkcalendar import Calendar

# Variável global para armazenar o DataFrame após a seleção do arquivo
df = None

def selecionar_data():
    def obter_data_selecionada():
        data = cal.get_date()
        data_inicial.set(data)
        top.destroy()

    top = tk.Toplevel(root)
    cal = Calendar(top, selectmode='day', day=datetime.now().day, month=datetime.now().month, year=datetime.now().year)
    cal.pack()
    tk.Button(top, text='Selecionar Data', command=obter_data_selecionada).pack()

# Função para selecionar o arquivo
def selecionar_arquivo():
    global df
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
    if arquivo:
        try:
            df = pd.read_excel(arquivo, sheet_name=1)
        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", f"Ocorreu um erro ao abrir o arquivo:\n{str(e)}")

# Função para gerar a planilha
def gerar_planilha():
    global df
    if df is not None:
        # Substituindo valores não-numéricos do DataFrame
        pd.set_option('future.no_silent_downcasting', True)
        df.replace(['FÉRIAS', 'DESLIGADA', 'LIC. MATER', 'TREINAMENTO'], [0,0,0,0], inplace=True)
        df.fillna(0, inplace=True) # Converte célula vazia para '0'

        primeiras_3_colunas = df.iloc[:, :3]
        data_inicial_datetime = data_inicial.get()
        datas_repetidas = pd.date_range(start=data_inicial_datetime, periods=int(dias_do_mes.get()))
        datas_repetidas = pd.Series(datas_repetidas).dt.strftime('%d/%m/%Y')
        datas_repetidas = pd.Series(datas_repetidas)

        colunas_restantes = df.iloc[:, 3:]
        linhas_empilhadas = []

        for index, row  in df.iterrows():
            primeiras_3_colunas_repetidas = pd.concat([primeiras_3_colunas.iloc[[index]]]*int(dias_do_mes.get()), ignore_index=True)
            primeiras_3_colunas_repetidas['Dia'] = datas_repetidas.tolist()
            linhas_empilhadas.extend(primeiras_3_colunas_repetidas.values.tolist())

        df_empilhado = pd.DataFrame(linhas_empilhadas, columns=list(primeiras_3_colunas.columns) + ['Dia'])
        demais_colunas_empilhadas = colunas_restantes.stack().reset_index(drop=True)
        df_empilhado['Meta'] = demais_colunas_empilhadas
        df_empilhado['Cargo'] = 'Vendedor'
        colunas = ['Cód.', 'Loja', 'Consultores', 'Meta', 'Dia', 'Cargo']
        resultado_final = df_empilhado[colunas]
        
        try:
            with pd.ExcelWriter('Meta para BI.xlsx', engine='xlsxwriter') as writer:
                resultado_final.to_excel(writer, index=False, sheet_name='Metas Consultores')
                worksheet_empilhado = writer.sheets['Metas Consultores']
                for i, col in enumerate(df_empilhado.columns):
                    column_len = max(df_empilhado[col].astype(str).map(len).max(), len(str(col))) + 2
                    worksheet_empilhado.set_column(i, i, column_len)
            root.destroy()
            messagebox.showinfo("Sucesso", "Planilha gerada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro ao gerar planilha", f"Ocorreu um erro ao gerar a planilha:\n{str(e)}")

root = tk.Tk()
root.title('Gerador de Planilha')

dias_do_mes = tk.IntVar()

tk.Button(root, text='Selecionar arquivo', command=selecionar_arquivo).pack() 

tk.Label(root, text='Quantos dias tem o mês?').pack()

tk.Radiobutton(root, text='28', variable=dias_do_mes, value=28).pack()
tk.Radiobutton(root, text='30', variable=dias_do_mes, value=30).pack()
tk.Radiobutton(root, text='31', variable=dias_do_mes, value=31).pack()

data_inicial = tk.StringVar()
data_inicial.set(datetime.now().strftime('%d/%m/%Y'))

tk.Button(root, text='Selecionar data inicial', command=selecionar_data).pack()
tk.Button(root, text='Gerar Planilha', command=gerar_planilha).pack()

root.mainloop()
