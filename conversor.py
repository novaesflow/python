import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de XLS/XLSX para CSV - Solidcon")
        
        self.create_widgets()
        
    def create_widgets(self):
        self.input_file_label = tk.Label(self.root, text="Arquivo Origem:")
        self.input_file_label.grid(row=0, column=0, padx=10, pady=10, sticky='e')
        
        self.input_file_entry = tk.Entry(self.root, width=50)
        self.input_file_entry.grid(row=0, column=1, padx=10, pady=10)
        
        self.input_file_button = tk.Button(self.root, text="Importar", command=self.import_file)
        self.input_file_button.grid(row=0, column=2, padx=10, pady=10)
        
        self.output_file_label = tk.Label(self.root, text="Arquivo Destino:")
        self.output_file_label.grid(row=1, column=0, padx=10, pady=10, sticky='e')
        
        self.output_file_entry = tk.Entry(self.root, width=50)
        self.output_file_entry.grid(row=1, column=1, padx=10, pady=10)
        
        self.output_file_button = tk.Button(self.root, text="Caminho", command=self.export_file)
        self.output_file_button.grid(row=1, column=2, padx=10, pady=10)
        
        self.convert_button = tk.Button(self.root, text="Converter", command=self.convert_file)
        self.convert_button.grid(row=2, column=1, padx=10, pady=10, sticky='e')
        
        self.cancel_button = tk.Button(self.root, text="Cancelar", command=self.cancel)
        self.cancel_button.grid(row=2, column=2, padx=10, pady=10, sticky='w')
    
    def import_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        self.input_file_entry.delete(0, tk.END)
        self.input_file_entry.insert(0, file_path)
    
    def export_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        self.output_file_entry.delete(0, tk.END)
        self.output_file_entry.insert(0, file_path)
    
    def convert_file(self):
        input_path = self.input_file_entry.get()
        output_path = self.output_file_entry.get()
        if not input_path or not output_path:
            messagebox.showerror("Erro", "Os campos de caminho do arquivo não podem estar vazios.")
            return
        try:
            df = pd.read_excel(input_path, engine='openpyxl' if input_path.endswith('.xlsx') else None)
            df.to_csv(output_path, index=False)
            messagebox.showinfo("Sucesso", "Arquivo convertido com sucesso!")
        except ImportError:
            messagebox.showerror("Erro", "Dependência faltando: openpyxl. Use pip para instalar openpyxl.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao converter arquivo: {e}")
    
    def cancel(self):
        self.root.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
