from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
from datetime import datetime
from PIL import Image

# Variável global para armazenar o DataFrame após a seleção do arquivo
df = None
data_inicial = None
list_range = None


def definir_data_inicial():
    global data_inicial
    data_atual = datetime.now()
    data_inicial = datetime(data_atual.year, data_atual.month, 1) # Definindo o dia padrão como 01, e o mês e ano de acordo com o atual
    data_inicial.strftime('%d/%m/%Y') # Formatando a data para dia/mês/ano

def selecionar_arquivo():
    global df
    global list_range
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
    if arquivo:
        try:
            list_range = dias_do_mes.get() + 5 # variavel para adicionar 5 colunas de acordo com o período selecionado
            df = pd.read_excel(arquivo, sheet_name=0, usecols=[0,1,3] + list(range(5, list_range))) # usando somente as colunas que interessam
        except Exception as e:
            messagebox.showerror("Erro ao abrir arquivo", f"Ocorreu um erro ao abrir o arquivo:\n{str(e)}")

def gerar_planilha():
    global df
    if df is not None:
        definir_data_inicial()
        # Substituindo valores não-numéricos do DataFrame
        pd.set_option('future.no_silent_downcasting', True)
        df.replace(['FÉRIAS', 'DESLIGADA', 'LIC. MATER', 'TREINAMENTO', 'AFASTADA'], [0,0,0,0,0], inplace=True)
        df.fillna(0, inplace=True) # Converte célula vazia para '0'

        primeiras_3_colunas = df.iloc[:, :3]
        datas_repetidas = pd.date_range(start=data_inicial, periods=int(dias_do_mes.get()))
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

# Criando a interface
root = ctk.CTk()
root.geometry("350x250")
root.title('Conversor de metas consultores')

# Aplicando fundo ao aplicativo
fundo = Image.open("fundo.jpg")
background_image = ctk.CTkImage(fundo, size=(350, 250))
bg_label = ctk.CTkLabel(root, text="", image=background_image)
bg_label.place(x=0, y=0)


dias_do_mes = ctk.IntVar()
dias_do_mes.set(30) # definido um valor padrão para evitar problemas com valor não definido

ctk.CTkLabel(root, text='Quantos dias tem o mês?', bg_color="#70967E", text_color="black").pack(anchor='center')

ctk.CTkRadioButton(root, text='28', variable=dias_do_mes, value=28, bg_color="#70967E", text_color="black", fg_color="#007e78", hover_color="#93DA49").pack(pady=5, anchor='center')
ctk.CTkRadioButton(root, text='30', variable=dias_do_mes, value=30, bg_color="#70967E", text_color="black", fg_color="#007e78", hover_color="#93DA49").pack(pady=5, anchor='center')
ctk.CTkRadioButton(root, text='31', variable=dias_do_mes, value=31, bg_color="#70967E", text_color="black", fg_color="#007e78", hover_color="#93DA49").pack(pady=5, anchor='center')

ctk.CTkButton(root, text='Selecionar arquivo', command=selecionar_arquivo, bg_color="#007E78", fg_color="#007e78", text_color="black", hover_color="#93DA49").pack(pady=15, anchor='center') 
ctk.CTkButton(root, text='Gerar Planilha', command=gerar_planilha, bg_color="#007E78", fg_color="#007e78", text_color="black", hover_color="#93DA49").pack(pady=5, anchor='center')

root.mainloop()
