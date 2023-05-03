import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, Toplevel, Label, Canvas, Scrollbar
import locale
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')



def analisar_planilha1(arquivo):
    df = pd.read_excel(arquivo, engine='openpyxl')
    colunas_interesse = ['Instituição', 'Cod_IBGE', 'UF', 'População', 'Coluna', 'Conta', 'Valor']
    df = df[colunas_interesse]
    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce').fillna(0)

    return df


def abrir_arquivo():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos do Excel", "*.xlsx;*.xls"), ("Arquivos CSV", "*.csv")])
    gastos_por_uf_conta = analisar_planilha1(arquivo)
    # Selecionar estado e cidade antes de exibir os resultados
    estados = sorted(gastos_por_uf_conta['UF'].unique())

    def estado_selecionado(*args):
        uf = var_estado.get()
        municipios_disponiveis = sorted(gastos_por_uf_conta[gastos_por_uf_conta['UF'] == uf]['Instituição'].unique())
        var_cidade.set('')
        opcao_cidade['values'] = municipios_disponiveis

    def exibir_cidade_selecionada(*args):
        uf = var_estado.get()
        cidade = var_cidade.get()
        if uf and cidade:
            exibir_resultados(uf, cidade, gastos_por_uf_conta)

    janela_selecao = Toplevel()
    janela_selecao.title("Seleção de Estado e Cidade")

    var_estado = tk.StringVar(janela_selecao)
    var_estado.trace('w', estado_selecionado)
    opcao_estado = ttk.Combobox(janela_selecao, textvariable=var_estado, values=estados, width=35)
    opcao_estado.grid(row=0, column=0, padx=10, pady=10)
    estado_label = tk.Label(janela_selecao, text="Selecione o estado:", font=("Arial", 14))
    estado_label.grid(row=0, column=1, padx=10, pady=10)

    var_cidade = tk.StringVar(janela_selecao)
    var_cidade.trace('w', exibir_cidade_selecionada)
    opcao_cidade = ttk.Combobox(janela_selecao, textvariable=var_cidade, width=35)
    opcao_cidade.grid(row=1, column=0, padx=10, pady=10)
    cidade_label = tk.Label(janela_selecao, text="Selecione a cidade:", font=("Arial", 14))
    cidade_label.grid(row=1, column=1, padx=10, pady=10)



def exibir_resultados(estado, municipio, gastos_por_uf_conta):
    janela_resultados = Toplevel()
    janela_resultados.title(f"Resultados para {municipio} - {estado}")
    janela_resultados.geometry("800x600")

    # Filtrar os dados com base no município e estado selecionados
    dados_municipio = gastos_por_uf_conta[(gastos_por_uf_conta['UF'] == estado) &
                                          (gastos_por_uf_conta['Instituição'] == municipio)]

    # Criar um DataFrame com os resultados consolidados por Coluna e Conta
    resultado_por_coluna_conta = dados_municipio.groupby(['Coluna', 'Conta'])['Valor'].sum().reset_index()

    # Criar um canvas com barra de rolagem para acomodar os dados
    canvas = Canvas(janela_resultados)
    canvas.pack(side="left", fill="both", expand=True)

    scrollbar = Scrollbar(janela_resultados, orient="vertical", command=canvas.yview)
    scrollbar.pack(side="left", fill="y")

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    # Adicionar os resultados no canvas
    frame_resultados = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame_resultados, anchor="nw")

    current_row = 0
    current_coluna = None

    for index, row in resultado_por_coluna_conta.iterrows():
        if current_coluna != row['Coluna']:
            current_coluna = row['Coluna']
            coluna_title = tk.Label(frame_resultados, text=row['Coluna'], font=("Arial", 16), pady=10)
            coluna_title.grid(row=current_row, column=0, columnspan=3)
            current_row += 1

            conta_label = tk.Label(frame_resultados, text="Conta", font=("Arial", 14))
            conta_label.grid(row=current_row, column=0, padx=10, pady=10)

            valor_label = tk.Label(frame_resultados, text="Valor", font=("Arial", 14))
            valor_label.grid(row=current_row, column=1, padx=10, pady=10)

            current_row += 1

        conta = tk.Label(frame_resultados, text=row['Conta'], font=("Arial", 12))
        valor = tk.Label(frame_resultados, text=locale.currency(row['Valor'], grouping=True), font=("Arial", 12))

        conta.grid(row=current_row, column=0, padx=10, pady=5)
        valor.grid(row=current_row, column=1, padx=10, pady=5)

        current_row += 1

root = tk.Tk()
root.title("Relatório - Segundo Exercício de Contabilidade")

root.geometry("1200x800")

container = tk.Frame(root)
container.pack(side="left", padx=10, pady=10)

abrir_arquivo_button = tk.Button(container, text="Abrir Planilha", command=abrir_arquivo, anchor="w", width=11, height=2, font=("Arial", 14))
abrir_arquivo_button.pack(pady=10)

root.mainloop()

