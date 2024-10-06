import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pyexcel as p
from openpyxl import Workbook

class PlanilhaUnificadoraApp:
    def __init__(self, master):
        self.master = master
        master.title("UNIFICADOR DE PLANILHAS PARA INVENTARIO - Versao 1.0")
        
        self.label = tk.Label(master, text="UNIFICADOR DE PLANILHAS PARA INVENTARIO - Versao 1.0", font=("Helvetica", 16, "bold"))
        self.label.grid(row=0, column=0, columnspan=3, pady=(10, 20))
        
        self.anexar_button = tk.Button(master, text="Anexar", command=self.anexar_arquivos, width=15, height=2)
        self.anexar_button.grid(row=1, column=2, padx=(5, 20), pady=5)
        
        self.caminho_button = tk.Button(master, text="Caminho", command=self.definir_caminho, width=15, height=2)
        self.caminho_button.grid(row=2, column=2, padx=(5, 20), pady=5)
        
        self.converter_button = tk.Button(master, text="Converter", command=self.converter, width=15, height=2)
        self.converter_button.grid(row=3, column=1, padx=5, pady=(20, 5))
        
        self.exportar_button = tk.Button(master, text="Exportar", command=self.exportar, width=15, height=2)
        self.exportar_button.grid(row=3, column=2, padx=(5, 20), pady=(20, 5))
        
        self.anexar_entry = tk.Entry(master, width=40)
        self.anexar_entry.grid(row=1, column=1, padx=(20, 5), pady=5)
        
        self.caminho_entry = tk.Entry(master, width=40)
        self.caminho_entry.grid(row=2, column=1, padx=(20, 5), pady=5)
        
        self.arquivos = []
        self.caminho_exportacao = ""

    def anexar_arquivos(self):
        self.arquivos = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xls *.xlsx *.xlsb")])
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

    def converter(self):
        if not self.arquivos:
            messagebox.showerror("Erro", "Nenhum arquivo anexado!")
            return
        
        pasta_convertidos = self.criar_pasta_convertidos()
        arquivos_convertidos = []

        for arquivo in self.arquivos:
            if arquivo.endswith('.xls'):
                arquivo = self.converter_xls_para_xlsx(arquivo, pasta_convertidos)
            elif arquivo.endswith('.xlsb'):
                # Se precisar lidar com arquivos .xlsb, pode adicionar uma função similar para converter .xlsb para .xlsx
                pass
            else:
                novo_caminho = os.path.join(pasta_convertidos, os.path.basename(arquivo))
                if not os.path.exists(novo_caminho):
                    os.rename(arquivo, novo_caminho)
                arquivo = novo_caminho

            arquivos_convertidos.append(arquivo)

        self.unificar_planilhas(arquivos_convertidos)

    def unificar_planilhas(self, arquivos):
        self.dados_unificados = pd.DataFrame(columns=['Quantidade', 'Codigo'])
        
        for arquivo in arquivos:
            df = pd.read_excel(arquivo, engine='openpyxl', header=None)
            df.columns = ['Quantidade', 'Codigo', 'Ignorar']
            
            # Filtrar linhas onde 'Codigo' contém apenas dígitos
            df = df[df['Codigo'].apply(lambda x: str(x).isdigit())]
            
            if 'Quantidade' not in df.columns or 'Codigo' not in df.columns:
                messagebox.showerror("Erro", f"Arquivo {arquivo} não contém as colunas necessárias.")
                return
            
            df = df[['Quantidade', 'Codigo']]
            self.dados_unificados = pd.concat([self.dados_unificados, df], ignore_index=True)
        
        self.dados_unificados = self.dados_unificados.groupby('Codigo', as_index=False).sum()
        
        # Formatar as colunas conforme solicitado
        self.dados_unificados['Codigo'] = self.dados_unificados['Codigo'].apply(lambda x: f"{int(x):014}")
        self.dados_unificados['Quantidade'] = self.dados_unificados['Quantidade'].apply(lambda x: f"{x:.3f}".replace('.', ','))
        
        messagebox.showinfo("Sucesso", "Conversão realizada com sucesso!")

    def exportar(self):
        if self.caminho_exportacao and hasattr(self, 'dados_unificados'):
            self.dados_unificados.to_csv(self.caminho_exportacao, sep=';', index=False, header=False)
            messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso para {self.caminho_exportacao}!")
        else:
            messagebox.showerror("Erro", "Defina o caminho de exportação e realize a conversão antes de exportar!")

if __name__ == "__main__":
    root = tk.Tk()
    app = PlanilhaUnificadoraApp(root)
    root.mainloop()
