import pandas as pd
import tkinter as tk
from tkinter import filedialog, Toplevel, Label, Canvas, Scrollbar
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


def analisar_planilha2(arquivo):
    df = pd.read_excel(arquivo, engine='openpyxl')
    colunas_interesse = ['Instituição', 'Cod_IBGE', 'UF', 'População', 'Coluna', 'Conta', 'Valor']
    df = df[colunas_interesse]
    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0)

    resultado_total = df.groupby('Coluna').agg({'Valor': 'sum', 'Instituição': 'nunique'}).reset_index()
    resultado_total.columns = ['Coluna', 'Valor Total', 'Municípios']

    resultado_conta = df.groupby('Conta').agg({'Valor': 'sum', 'Instituição': 'nunique'}).reset_index()
    resultado_conta.columns = ['Conta', 'Valor Total', 'Municípios']

    gastos_por_uf_coluna = df.groupby(['UF', 'Coluna']).agg({'Valor': 'sum'}).reset_index()
    gastos_por_uf_conta = df.groupby(['UF', 'Conta']).agg({'Valor': 'sum'}).reset_index()

    return resultado_total, resultado_conta, gastos_por_uf_coluna, gastos_por_uf_conta

def abrir_arquivo2():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos do Excel", "*.xlsx;*.xls"), ("Arquivos CSV", "*.csv")])
    resultado_total, resultado_conta, gastos_por_uf_coluna, gastos_por_uf_conta = analisar_planilha2(arquivo)
    exibir_resultados(resultado_total, resultado_conta, gastos_por_uf_coluna, gastos_por_uf_conta)



def exibir_resultados(resultado_total, resultado_conta, gastos_por_uf_coluna, gastos_por_uf_conta):
    def dataframe_to_list_string(df, col_space=0):
        formatted_rows = []
        for index, row in df.iterrows():
            row_str = []
            for col in df.columns:
                if col == 'Municípios':
                    formatted_value = f"{col}: {int(row[col])}"
                elif isinstance(row[col], (int, float)):
                    formatted_value = f"{col}: {locale.currency(row[col], grouping=True)}"
                else:
                    formatted_value = f"{col}: {row[col]}"
                row_str.append(formatted_value.ljust(col_space))
            formatted_rows.append("; ".join(row_str))
        return "\n".join(formatted_rows).replace("; Coluna", "Coluna").replace("; Conta", "Conta")

    resultado_total_text = dataframe_to_list_string(resultado_total, col_space=16)
    resultado_conta_text = dataframe_to_list_string(resultado_conta, col_space=16)
    gastos_por_uf_coluna_text = dataframe_to_list_string(gastos_por_uf_coluna, col_space=16)

    resultado_window = Toplevel()
    resultado_window.title("Resultados")
    resultado_window.geometry("1300x800")

    canvas = Canvas(resultado_window, bg="white")
    scrollbar_y = Scrollbar(resultado_window, orient="vertical", command=canvas.yview)
    scrollbar_x = Scrollbar(resultado_window, orient="horizontal", command=canvas.xview)
    frame_resultados = tk.Frame(canvas, bg="white")

    canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x.pack(side="bottom", fill="x")

    canvas.create_window((0, 0), window=frame_resultados, anchor="nw")
    frame_resultados.bind("<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all")))

    titulo_label = Label(frame_resultados, text="Relatório - Segundo Exercício de Contabilidade", font=("Arial", 16, "bold"), justify="left", bg="white")
    titulo_label.pack(padx=10, pady=10)

    resultado_total_label = Label(frame_resultados, text="Valor Total por Categoria de despesas(somente para as subfunções de saúde)", font=("Arial", 14, "bold"), justify="left", bg="white")
    resultado_total_label.pack(padx=10, pady=2)
    resultado_total_text_label = Label(frame_resultados, text=resultado_total_text, font=("Arial", 14), justify="left", bg="white")
    resultado_total_text_label.pack(padx=10, pady=2)

    resultado_conta_label = Label(frame_resultados, text="Valor Total por Funções e Quantidade de Municípios:", font=("Arial", 14, "bold"), justify="left", bg="white")
    resultado_conta_label.pack(padx=10, pady=2)
    resultado_conta_text_label = Label(frame_resultados, text=resultado_conta_text, font=("Arial", 14), justify="left", bg="white")
    resultado_conta_text_label.pack(padx=10, pady=2)

    gastos_por_uf_coluna_label = Label(frame_resultados, text="Gastos por Estado de Federação por Despesa:", font=("Arial", 14, "bold"), justify="left", bg="white")
    gastos_por_uf_coluna_label.pack(padx=10, pady=2)
    gastos_por_uf_coluna_text_label = Label(frame_resultados, text=gastos_por_uf_coluna_text, font=("Arial", 14), justify="left", bg="white")
    gastos_por_uf_coluna_text_label.pack(padx=10, pady=2)

    gastos_por_uf_conta_label = Label(frame_resultados,
                                      text="Gasto por Estado de Federação por Função:",
                                      font=("Arial", 14, "bold"), justify="left", bg="white")
    gastos_por_uf_conta_label.pack(padx=10, pady=2)
    gastos_por_uf_conta_text = dataframe_to_list_string(gastos_por_uf_conta, col_space=16)
    gastos_por_uf_conta_text_label = Label(frame_resultados, text=gastos_por_uf_conta_text, font=("Arial", 14),
                                           justify="left", bg="white")
    gastos_por_uf_conta_text_label.pack(padx=10, pady=2)


root = tk.Tk()
root.title("Relatório - Segundo Exercício de Contabilidade")

root.geometry("1200x800")

container = tk.Frame(root)
container.pack(side="left", padx=10, pady=10)

abrir_arquivo_button2 = tk.Button(container, text="Abrir Planilha", command=abrir_arquivo2, anchor="w", width=11, height=2, font=("Arial", 14))
abrir_arquivo_button2.pack(pady=10)


root.mainloop()
