import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import pyexcel as p
import csv
import chardet

class PlanilhaUnificadoraApp:
    def __init__(self, master):
        self.master = master
        master.title("UNIFICADOR DE PLANILHAS PARA INVENTARIO - Versao 1.0")
        master.geometry('1024x800')

        self.tab_control = ttk.Notebook(master)
        
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab2 = ttk.Frame(self.tab_control)
        self.tab3 = ttk.Frame(self.tab_control)
        self.tab4 = ttk.Frame(self.tab_control)

        self.tab_control.add(self.tab1, text='Processar')
        self.tab_control.add(self.tab2, text='Auditoria')
        self.tab_control.add(self.tab3, text='Unificação Manual')
        self.tab_control.add(self.tab4, text='Tratamento CSV')
        self.tab_control.pack(expand=1, fill="both")
        
        # Tab 1: Processar
        self.label = tk.Label(self.tab1, text="Processar Produtos", font=("Helvetica", 16, "bold"))
        self.label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
        
        self.anexar_button = tk.Button(self.tab1, text="Anexar", command=self.anexar_arquivos, width=15, height=2)
        self.anexar_button.grid(row=1, column=2, padx=(5, 20), pady=5)
        
        self.caminho_button = tk.Button(self.tab1, text="Caminho", command=self.definir_caminho, width=15, height=2)
        self.caminho_button.grid(row=2, column=2, padx=(5, 20), pady=5)
        
        self.processar_button = tk.Button(self.tab1, text="Processar", command=self.processar, width=15, height=2)
        self.processar_button.grid(row=3, column=1, columnspan=2, padx=(5, 20), pady=(20, 5))
        
        self.anexar_entry = tk.Entry(self.tab1, width=40)
        self.anexar_entry.grid(row=1, column=1, padx=(20, 5), pady=5)
        
        self.caminho_entry = tk.Entry(self.tab1, width=40)
        self.caminho_entry.grid(row=2, column=1, padx=(20, 5), pady=5)
        
        self.arquivos = []
        self.caminho_exportacao = ""
        self.dados_unificados = pd.DataFrame()
        self.arquivos_convertidos = []

        # Tab 2: Auditoria
        self.audit_label = tk.Label(self.tab2, text="Auditoria de Produtos", font=("Helvetica", 16, "bold"))
        self.audit_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))

        self.codigo_entry = tk.Entry(self.tab2, width=40)
        self.codigo_entry.grid(row=1, column=1, padx=(20, 5), pady=5)
        self.codigo_label = tk.Label(self.tab2, text="Código do Produto:")
        self.codigo_label.grid(row=1, column=0, padx=(20, 5), pady=5)

        self.auditar_button = tk.Button(self.tab2, text="Auditar", command=self.auditar, width=15, height=2)
        self.auditar_button.grid(row=1, column=2, padx=(5, 20), pady=5)

        self.result_tree = ttk.Treeview(self.tab2, columns=("codigo", "quantidade", "arquivo"), show='headings')
        self.result_tree.heading("codigo", text="Código")
        self.result_tree.heading("quantidade", text="Quantidade")
        self.result_tree.heading("arquivo", text="Arquivo")
        self.result_tree.grid(row=2, column=0, columnspan=3, padx=(20, 20), pady=(20, 20), sticky='nsew')

        self.tab2.grid_rowconfigure(2, weight=1)
        self.tab2.grid_columnconfigure(1, weight=1)

        # Tab 3: Unificação Manual
        self.manual_label = tk.Label(self.tab3, text="Unificação Manual de Produtos", font=("Helvetica", 16, "bold"))
        self.manual_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
        
        self.anexar_manual_button = tk.Button(self.tab3, text="Anexar", command=self.anexar_arquivos_manual, width=15, height=2)
        self.anexar_manual_button.grid(row=1, column=2, padx=(5, 20), pady=5)
        
        self.caminho_manual_button = tk.Button(self.tab3, text="Caminho", command=self.definir_caminho_manual, width=15, height=2)
        self.caminho_manual_button.grid(row=2, column=2, padx=(5, 20), pady=5)
        
        self.unificar_manual_button = tk.Button(self.tab3, text="Unificar", command=self.unificar_manual, width=15, height=2)
        self.unificar_manual_button.grid(row=3, column=1, columnspan=2, padx=(5, 20), pady=(20, 5))
        
        self.anexar_manual_entry = tk.Entry(self.tab3, width=40)
        self.anexar_manual_entry.grid(row=1, column=1, padx=(20, 5), pady=5)
        
        self.caminho_manual_entry = tk.Entry(self.tab3, width=40)
        self.caminho_manual_entry.grid(row=2, column=1, padx=(20, 5), pady=5)

        self.arquivos_manual = []
        self.caminho_manual_exportacao = ""
        self.dados_unificados_manual = pd.DataFrame()

        # Tab 4: Tratamento CSV
        self.tratamento_label = tk.Label(self.tab4, text="Tratamento de CSV", font=("Helvetica", 16, "bold"))
        self.tratamento_label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
        
        self.anexar_tratamento_button = tk.Button(self.tab4, text="Anexar CSV", command=self.anexar_arquivos_tratamento, width=15, height=2)
        self.anexar_tratamento_button.grid(row=1, column=2, padx=(5, 20), pady=5)
        
        self.caminho_tratamento_button = tk.Button(self.tab4, text="Caminho CSV", command=self.definir_caminho_tratamento, width=15, height=2)
        self.caminho_tratamento_button.grid(row=2, column=2, padx=(5, 20), pady=5)
        
        self.tratar_tratamento_button = tk.Button(self.tab4, text="Tratar CSV", command=self.tratar_csv, width=15, height=2)
        self.tratar_tratamento_button.grid(row=3, column=1, columnspan=2, padx=(5, 20), pady=(20, 5))
        
        self.anexar_tratamento_entry = tk.Entry(self.tab4, width=40)
        self.anexar_tratamento_entry.grid(row=1, column=1, padx=(20, 5), pady=5)
        
        self.caminho_tratamento_entry = tk.Entry(self.tab4, width=40)
        self.caminho_tratamento_entry.grid(row=2, column=1, padx=(20, 5), pady=5)

        self.arquivos_tratamento = []
        self.caminho_tratamento_exportacao = ""

    def anexar_arquivos(self):
        self.arquivos = filedialog.askopenfilenames(filetypes=[("Excel and CSV files", "*.xls *.xlsx *.xlsb *.csv")])
        self.anexar_entry.delete(0, tk.END)
        self.anexar_entry.insert(0, ', '.join(self.arquivos))

    def definir_caminho(self):
        self.caminho_exportacao = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        self.caminho_entry.delete(0, tk.END)
        self.caminho_entry.insert(0, self.caminho_exportacao)

    def converter_xls_para_xlsx(self, caminho_xls, pasta_destino):
        nome_arquivo = os.path.basename(caminho_xls).replace('.xls', '.xlsx')
        caminho_xlsx = os.path.join(pasta_destino, nome_arquivo)
        p.save_book_as(file_name=caminho_xls, dest_file_name=caminho_xlsx)
        return caminho_xlsx

    def criar_pasta_convertidos(self):
        pasta_convertidos = os.path.join(os.path.dirname(self.arquivos[0]), 'convertidos_xlsx')
        if not os.path.exists(pasta_convertidos):
            os.makedirs(pasta_convertidos)
        return pasta_convertidos

    def processar(self):
        if not self.arquivos:
            messagebox.showerror("Erro", "Nenhum arquivo anexado!")
            return
        
        pasta_convertidos = self.criar_pasta_convertidos()
        self.arquivos_convertidos = []

        for arquivo in self.arquivos:
            try:
                if arquivo.endswith('.xls'):
                    arquivo = self.converter_xls_para_xlsx(arquivo, pasta_convertidos)
                elif arquivo.endswith('.xlsb'):
                    # Se precisar lidar com arquivos .xlsb, pode adicionar uma função similar para converter .xlsb para .xlsx
                    pass
                elif arquivo.endswith('.csv'):
                    self.arquivos_convertidos.append(arquivo)
                    continue
                else:
                    novo_caminho = os.path.join(pasta_convertidos, os.path.basename(arquivo))
                    if not os.path.exists(novo_caminho):
                        os.rename(arquivo, novo_caminho)
                    arquivo = novo_caminho

                self.arquivos_convertidos.append(arquivo)
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao converter o arquivo {arquivo}: {str(e)}")
                return

        self.unificar_planilhas(self.arquivos_convertidos)
        self.exportar()

    def unificar_planilhas(self, arquivos):
        self.dados_unificados = pd.DataFrame(columns=['Quantidade', 'Codigo'])
        
        for arquivo in arquivos:
            try:
                if arquivo.endswith('.csv'):
                    df = pd.read_csv(arquivo, delimiter=';', header=None, names=['Codigo', 'Quantidade'])
                else:
                    df = pd.read_excel(arquivo, engine='openpyxl', header=None)
                    df.columns = ['Quantidade', 'Codigo', 'Ignorar']
                    df = df[['Quantidade', 'Codigo']]

                # Filtrar linhas onde 'Codigo' contém apenas dígitos
                df = df[df['Codigo'].apply(lambda x: str(x).isdigit())]
                
                if 'Quantidade' not in df.columns or 'Codigo' not in df.columns:
                    messagebox.showerror("Erro", f"Arquivo {arquivo} não contém as colunas necessárias.")
                    return
                
                self.dados_unificados = pd.concat([self.dados_unificados, df], ignore_index=True)
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao ler o arquivo {arquivo}: {str(e)}")
                return
        
        self.dados_unificados = self.dados_unificados.groupby('Codigo', as_index=False).sum()
        
        # Formatar as colunas conforme solicitado
        self.dados_unificados['Codigo'] = self.dados_unificados['Codigo'].apply(lambda x: f"{int(x):014}")
        self.dados_unificados['Quantidade'] = self.dados_unificados['Quantidade'].apply(lambda x: f"{x:.3f}".replace('.', ','))

        messagebox.showinfo("Sucesso", "Conversão realizada com sucesso!")

    def exportar(self):
        if self.caminho_exportacao and not self.dados_unificados.empty:
            with open(self.caminho_exportacao, 'w', newline='') as file:
                writer = csv.writer(file, delimiter=';', quotechar='', quoting=csv.QUOTE_NONE, escapechar='\\')
                for index, row in self.dados_unificados.iterrows():
                    writer.writerow([row['Codigo'], row['Quantidade']])
            messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso para {self.caminho_exportacao}!")
        else:
            messagebox.showerror("Erro", "Defina o caminho de exportação e realize a conversão antes de exportar!")

    def auditar(self):
        self.result_tree.delete(*self.result_tree.get_children())
        
        codigo_produto = self.codigo_entry.get().strip()
        if not codigo_produto:
            messagebox.showerror("Erro", "Por favor, insira um código de produto para auditar.")
            return
        
        codigo_produto = str(codigo_produto).zfill(14)  # Normalizar para 14 dígitos
        resultados = []

        for arquivo in self.arquivos_convertidos:
            try:
                if arquivo.endswith('.csv'):
                    df = pd.read_csv(arquivo, delimiter=';', header=None, names=['Codigo', 'Quantidade'])
                else:
                    df = pd.read_excel(arquivo, engine='openpyxl', header=None)
                    df.columns = ['Quantidade', 'Codigo', 'Ignorar']
                    df['Codigo'] = df['Codigo'].apply(lambda x: str(x).zfill(14))
                    
                if codigo_produto in df['Codigo'].values:
                    quantidade = df[df['Codigo'] == codigo_produto]['Quantidade'].sum()
                    resultados.append((codigo_produto, quantidade, os.path.basename(arquivo)))
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao ler o arquivo {arquivo}: {str(e)}")
                return
        
        quantidade_unificada = 0
        if codigo_produto in self.dados_unificados['Codigo'].values:
            quantidade_unificada = self.dados_unificados[self.dados_unificados['Codigo'] == codigo_produto]['Quantidade'].values[0]
        
        for resultado in resultados:
            self.result_tree.insert('', 'end', values=resultado)
        
        self.result_tree.insert('', 'end', values=(codigo_produto, quantidade_unificada, 'Total Unificado'))

    def anexar_arquivos_manual(self):
        self.arquivos_manual = filedialog.askopenfilenames(filetypes=[("Excel and CSV files", "*.xls *.xlsx *.xlsb *.csv")])
        self.anexar_manual_entry.delete(0, tk.END)
        self.anexar_manual_entry.insert(0, ', '.join(self.arquivos_manual))

    def definir_caminho_manual(self):
        self.caminho_manual_exportacao = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        self.caminho_manual_entry.delete(0, tk.END)
        self.caminho_manual_entry.insert(0, self.caminho_manual_exportacao)

    def unificar_manual(self):
        if not self.arquivos_manual:
            messagebox.showerror("Erro", "Nenhum arquivo anexado!")
            return

        self.dados_unificados_manual = pd.DataFrame(columns=['Quantidade', 'Codigo'])

        for arquivo in self.arquivos_manual:
            try:
                if arquivo.endswith('.csv'):
                    df = pd.read_csv(arquivo, delimiter=';', header=None, names=['Codigo', 'Quantidade'])
                else:
                    df = pd.read_excel(arquivo, engine='openpyxl', header=None)
                    df.columns = ['Quantidade', 'Codigo', 'Ignorar']
                    df = df[['Quantidade', 'Codigo']]

                # Filtrar linhas onde 'Codigo' contém apenas dígitos
                df = df[df['Codigo'].apply(lambda x: str(x).isdigit())]
                
                if 'Quantidade' not in df.columns or 'Codigo' not in df.columns:
                    messagebox.showerror("Erro", f"Arquivo {arquivo} não contém as colunas necessárias.")
                    return
                
                self.dados_unificados_manual = pd.concat([self.dados_unificados_manual, df], ignore_index=True)
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao ler o arquivo {arquivo}: {str(e)}")
                return
        
        self.dados_unificados_manual = self.dados_unificados_manual.groupby('Codigo', as_index=False).sum()
        
        # Formatar as colunas conforme solicitado
        self.dados_unificados_manual['Codigo'] = self.dados_unificados_manual['Codigo'].apply(lambda x: f"{int(x):014}")
        self.dados_unificados_manual['Quantidade'] = self.dados_unificados_manual['Quantidade'].apply(lambda x: f"{x:.3f}".replace('.', ','))

        if self.caminho_manual_exportacao and not self.dados_unificados_manual.empty:
            with open(self.caminho_manual_exportacao, 'w', newline='') as file:
                writer = csv.writer(file, delimiter=';', quotechar='', quoting=csv.QUOTE_NONE, escapechar='\\')
                for index, row in self.dados_unificados_manual.iterrows():
                    writer.writerow([row['Codigo'], row['Quantidade']])
            messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso para {self.caminho_manual_exportacao}!")
        else:
            messagebox.showerror("Erro", "Defina o caminho de exportação e realize a unificação antes de exportar!")

    def anexar_arquivos_tratamento(self):
        self.arquivos_tratamento = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
        self.anexar_tratamento_entry.delete(0, tk.END)
        self.anexar_tratamento_entry.insert(0, ', '.join(self.arquivos_tratamento))

    def definir_caminho_tratamento(self):
        self.caminho_tratamento_exportacao = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        self.caminho_tratamento_entry.delete(0, tk.END)
        self.caminho_tratamento_entry.insert(0, self.caminho_tratamento_exportacao)

    def tratar_csv(self):
        if not self.arquivos_tratamento:
            messagebox.showerror("Erro", "Nenhum arquivo CSV anexado!")
            return

        for arquivo in self.arquivos_tratamento:
            try:
                with open(arquivo, 'rb') as f:
                    result = chardet.detect(f.read())
                encoding = result['encoding']

                df = pd.read_csv(arquivo, delimiter=';', encoding=encoding)

                # Selecionar apenas as colunas necessárias
                if 'Código do Produto' not in df.columns or 'Digitado' not in df.columns:
                    messagebox.showerror("Erro", f"Arquivo {arquivo} não contém as colunas necessárias.")
                    return

                df_selecionado = df[['Código do Produto', 'Digitado']]

                # Substituir vírgulas por pontos para a conversão para float
                df_selecionado['Digitado'] = df_selecionado['Digitado'].str.replace(',', '.')

                # Formatar a coluna Código do Produto para ter 14 dígitos e a coluna Digitado para ter 3 casas decimais
                df_selecionado['Código do Produto'] = df_selecionado['Código do Produto'].apply(lambda x: f"{int(x):014}")
                df_selecionado['Digitado'] = df_selecionado['Digitado'].apply(lambda x: f"{float(x):.3f}".replace('.', ','))

                # Salvar o resultado em um novo arquivo CSV
                df_selecionado.to_csv(self.caminho_tratamento_exportacao, sep=';', index=False, header=False, quoting=csv.QUOTE_NONE, quotechar='', escapechar='\\')
                
                messagebox.showinfo("Sucesso", f"Arquivo tratado exportado para {self.caminho_tratamento_exportacao}!")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao ler ou tratar o arquivo {arquivo}: {str(e)}")
                return

if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaUnificadoraApp(root)
    root.mainloop()


       
