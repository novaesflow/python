"""
========================================
Nome do Script: FlowList
Autor: Tiego Novaes Santana
Data de Criação: 01/09/2024
Versão: 1.0
Descrição: Este script é utilizado para listar produtos e preços de um banco de dados SQL Server.
Git Repository: https://github.com/novaesflow/python.git
========================================
"""

import sys
import pyodbc
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QTextEdit, QLabel, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem, QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QMimeData
from PyQt5.QtWidgets import QHeaderView  # Importação adicional

class FlowListApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("FlowList")
        self.setGeometry(200, 200, 800, 600)

        # Definindo o ícone do programa
        self.setWindowIcon(QIcon('C:\\Python\\Biblioteca de icones\\Flow List\\icone_programa.png'))

        # Tabs
        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        # Criação da aba de listagem de preços
        self.create_price_listing_tab()

        # Inicializando conexão SQL
        self.conn = None
        self.connect_to_database()

    def create_price_listing_tab(self):
        price_tab = QWidget()
        layout = QVBoxLayout()

        # Tabela para exibir os produtos e preços
        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Descrição do Produto", "Preço (R$)"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)  # Corrigido aqui
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Corrigido aqui
        layout.addWidget(self.table)

        # Layout horizontal para botões
        button_layout = QHBoxLayout()

        # Botão para gerar os dados
        generate_button = QPushButton("Gerar")
        generate_button.setIcon(QIcon('C:\\Python\\Biblioteca de icones\\Flow List\\botao_gerar.png'))
        generate_button.clicked.connect(self.load_data)
        button_layout.addWidget(generate_button)

        # Botão para copiar para a área de transferência
        copy_button = QPushButton("Copiar")
        copy_button.setIcon(QIcon('C:\\Python\\Biblioteca de icones\\Flow List\\botao_copiar.png'))
        copy_button.clicked.connect(self.copy_to_clipboard)
        button_layout.addWidget(copy_button)

        # Adicionar layout de botões ao layout principal
        layout.addLayout(button_layout)

        price_tab.setLayout(layout)
        self.tabs.addTab(price_tab, "Listagem de Preço")

    def connect_to_database(self):
        connection_string = 'DRIVER={SQL Server};SERVER=localhost;DATABASE=solidcon;UID=Solidcon;PWD=Nocdilos'

        try:
            self.conn = pyodbc.connect(connection_string)
        except Exception as e:
            self.show_error(f"Erro ao conectar ao banco de dados: {e}")

    def load_data(self):
        if self.conn:
            cursor = self.conn.cursor()
            try:
                query = """
                    SELECT 
                        sp.nmprodutopai, 
                        spv.vlvenda
                    FROM tbsuperproduto sp
                    INNER JOIN tbsuperprodutovenda spv 
                        ON sp.cdsuperproduto = spv.cdsuperproduto
                    WHERE spv.vlvenda > 0
                """
                cursor.execute(query)
                results = cursor.fetchall()

                # Preencher a tabela com os dados obtidos
                self.table.setRowCount(len(results))
                for row_idx, (descricao, preco) in enumerate(results):
                    descricao_item = QTableWidgetItem(descricao)
                    descricao_item.setTextAlignment(Qt.AlignCenter)  # Centralizar o texto
                    self.table.setItem(row_idx, 0, descricao_item)

                    preco_item = QTableWidgetItem(f"R$ {preco:.2f}")
                    preco_item.setTextAlignment(Qt.AlignCenter)  # Centralizar o texto
                    self.table.setItem(row_idx, 1, preco_item)

            except pyodbc.Error as e:
                self.show_error(f"Erro ao executar a consulta SQL: {e}")

    def copy_to_clipboard(self):
        clipboard = QApplication.clipboard()
        mime_data = QMimeData()

        data = ""
        for row in range(self.table.rowCount()):
            descricao = self.table.item(row, 0).text()
            preco = self.table.item(row, 1).text()
            data += f"{descricao} - {preco}\n"

        mime_data.setText(data)
        clipboard.setMimeData(mime_data)
        self.show_message("Dados copiados com sucesso!")  # Adiciona pop-up de confirmação

    def show_error(self, message):
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Critical)
        error_dialog.setText(message)
        error_dialog.setWindowTitle("Erro")
        error_dialog.exec_()

    def show_message(self, message):
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setText(message)
        msg_box.setWindowTitle("Mensagem")
        msg_box.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowListApp()
    window.show()
    sys.exit(app.exec_())
