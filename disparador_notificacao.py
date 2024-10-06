"""
========================================
Nome do Script: Notificador de WhatsApp via SQL Server
Autor: Tiego Novaes Santana
Data de Criação: 03/09/2024
Versão: 1.2
Descrição: Este script é utilizado para enviar notificações via WhatsApp usando dados de um banco de dados SQL Server.
Git Repository: https://github.com/novaesflow/python.git
========================================
"""
import sys
import pyodbc
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QTextEdit, QLabel, QTabWidget, QWidget, QVBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem
from PyQt5.QtGui import QIcon  # Import necessário para trabalhar com ícones
import requests

class WhatsAppNotifierApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Flow Notify")
        self.setGeometry(200, 200, 800, 600)

        # ícone da janela
        self.setWindowIcon(QIcon('C:\\Python\\Biblioteca de icones\\Flow Notify\\aviao-de-papel.png'))

        # Tabs
        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        self.create_sql_tab()
        self.create_whatsapp_tab()
        self.create_notifications_tab()

        # Inicializando conexão SQL e configurações da API
        self.conn = None
        self.api_url = ""
        self.api_token = ""

    def create_sql_tab(self):
        sql_tab = QWidget()
        layout = QVBoxLayout()

        # Layout em grid para organizar os campos de entrada e o botão
        grid_layout = QGridLayout()

        # Inputs de conexão
        self.server_input = QLineEdit()
        self.database_input = QLineEdit()
        self.username_input = QLineEdit()
        self.password_input = QLineEdit()

        self.server_input.setPlaceholderText("Servidor")
        self.database_input.setPlaceholderText("Banco de Dados")
        self.username_input.setPlaceholderText("Usuário")
        self.password_input.setPlaceholderText("Senha")
        self.password_input.setEchoMode(QLineEdit.Password)  # Mascarar entrada de senha

        # Adicionando widgets ao layout em grid
        grid_layout.addWidget(QLabel("Servidor"), 0, 0)
        grid_layout.addWidget(self.server_input, 0, 1)
        grid_layout.addWidget(QLabel("Banco de Dados"), 1, 0)
        grid_layout.addWidget(self.database_input, 1, 1)
        grid_layout.addWidget(QLabel("Usuário"), 2, 0)
        grid_layout.addWidget(self.username_input, 2, 1)
        grid_layout.addWidget(QLabel("Senha"), 3, 0)
        grid_layout.addWidget(self.password_input, 3, 1)

        # Botão de conexão
        connect_button = QPushButton("Conectar")
        connect_button.clicked.connect(self.connect_to_database)
        grid_layout.addWidget(connect_button, 4, 0, 1, 2)  # Botão ocupa duas colunas

        # Adicionando o layout de grid ao layout vertical principal
        layout.addLayout(grid_layout)

        # Caixa de texto para mostrar status
        self.log_sql = QTextEdit()
        self.log_sql.setReadOnly(True)
        layout.addWidget(self.log_sql)

        sql_tab.setLayout(layout)
        self.tabs.addTab(sql_tab, "Conexão SQL")

    def create_whatsapp_tab(self):
        whatsapp_tab = QWidget()
        layout = QVBoxLayout()

        # Layout em grid para organizar os campos de entrada e o botão
        grid_layout = QGridLayout()

        # Inputs para as credenciais da API do WhatsApp
        self.api_url_input = QLineEdit()
        self.api_token_input = QLineEdit()
        self.api_url_input.setPlaceholderText("URL da API do WhatsApp")
        self.api_token_input.setPlaceholderText("Token de Autenticação da API")

        # Adicionando widgets ao layout em grid
        grid_layout.addWidget(QLabel("URL da API do WhatsApp"), 0, 0)
        grid_layout.addWidget(self.api_url_input, 0, 1)
        grid_layout.addWidget(QLabel("Token de Autenticação"), 1, 0)
        grid_layout.addWidget(self.api_token_input, 1, 1)

        # Botão para salvar configurações
        save_button = QPushButton("Salvar Configuração")
        save_button.clicked.connect(self.save_whatsapp_config)
        grid_layout.addWidget(save_button, 2, 0, 1, 2)

        # Adicionando o layout de grid ao layout vertical principal
        layout.addLayout(grid_layout)

        # Caixa de texto para mostrar status
        self.log_whatsapp = QTextEdit()
        self.log_whatsapp.setReadOnly(True)
        layout.addWidget(self.log_whatsapp)

        whatsapp_tab.setLayout(layout)
        self.tabs.addTab(whatsapp_tab, "WhatsApp")

    def create_notifications_tab(self):
        notifications_tab = QWidget()
        layout = QVBoxLayout()

        # Tabela para mostrar notificações
        self.notifications_table = QTableWidget()
        self.notifications_table.setColumnCount(4)
        self.notifications_table.setHorizontalHeaderLabels(['Data', 'Fornecedor', 'Número', 'Status'])
        layout.addWidget(self.notifications_table)

        notifications_tab.setLayout(layout)
        self.tabs.addTab(notifications_tab, "Notificações")

        # Disparar as notificações automaticamente ao carregar a aba
        self.tabs.currentChanged.connect(self.load_and_send_notifications)

    def connect_to_database(self):
        server = self.server_input.text()
        database = self.database_input.text()
        username = self.username_input.text()
        password = self.password_input.text()

        connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

        try:
            self.conn = pyodbc.connect(connection_string)
            self.log_sql.append("Conexão com o banco de dados realizada com sucesso.")
        except Exception as e:
            self.log_sql.append(f"Erro ao conectar ao banco de dados: {e}")

    def save_whatsapp_config(self):
        self.api_url = self.api_url_input.text()
        self.api_token = self.api_token_input.text()

        if self.api_url and self.api_token:
            self.log_whatsapp.append("Configuração da API do WhatsApp salva com sucesso.")
        else:
            self.log_whatsapp.append("Por favor, preencha todos os campos.")

    def load_and_send_notifications(self):
        if self.tabs.currentIndex() == 2:  # Verifica se a aba atual é "Notificações"
            if self.conn is None:
                self.log_whatsapp.append("Por favor, conecte-se ao banco de dados primeiro.")
                return

            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT n.dtAlteracao, p.nmPessoa, t.DDD, t.Numero
                FROM tbNegociacao n
                JOIN tbPessoa p ON p.cdPessoa = n.cdPessoaComercial
                JOIN tbTelefone t ON t.cdPessoa = n.cdPessoaComercial
                WHERE n.dtAlteracao IS NOT NULL
            """)
            negociacoes = cursor.fetchall()

            self.notifications_table.setRowCount(len(negociacoes))

            for i, negociacao in enumerate(negociacoes):
                dt_alteracao = negociacao.dtAlteracao
                fornecedor = negociacao.nmPessoa
                ddd = negociacao.DDD
                numero = negociacao.Numero
                telefone_completo = f"+{ddd}{numero}"

                # Disparo de mensagem automática
                status = self.send_whatsapp_message(telefone_completo, f"Olá {fornecedor}, há uma nova negociação com alteração em {dt_alteracao}.")

                self.notifications_table.setItem(i, 0, QTableWidgetItem(str(dt_alteracao)))
                self.notifications_table.setItem(i, 1, QTableWidgetItem(fornecedor))
                self.notifications_table.setItem(i, 2, QTableWidgetItem(telefone_completo))
                self.notifications_table.setItem(i, 3, QTableWidgetItem(status))

            # Ajustar as larguras das colunas para melhor visualização
            self.notifications_table.setColumnWidth(0, 120)  # Data
            self.notifications_table.setColumnWidth(1, 400)  # Fornecedor
            self.notifications_table.setColumnWidth(2, 150)  # Número
            self.notifications_table.setColumnWidth(3, 100)  # Status

    def send_whatsapp_message(self, to_number, message_body):
        if not self.api_url or not self.api_token:
            self.log_whatsapp.append("Configuração da API do WhatsApp não está completa.")
            return "Configuração API incompleta"

        message_data = {
            "messaging_product": "whatsapp",
            "to": to_number,
            "type": "text",
            "text": {
                "body": message_body
            }
        }

        headers = {
            "Authorization": f"Bearer {self.api_token}",
            "Content-Type": "application/json"
        }

        try:
            response = requests.post(self.api_url, json=message_data, headers=headers)
            if response.status_code == 200:
                self.log_whatsapp.append(f"Mensagem enviada para {to_number}.")
                return "Enviada"
            else:
                self.log_whatsapp.append(f"Erro ao enviar mensagem para {to_number}: {response.text}")
                return "Erro no envio"
        except Exception as e:
            self.log_whatsapp.append(f"Erro ao enviar mensagem para {to_number}: {e}")
            return "Erro de conexão"

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = WhatsAppNotifierApp()
    window.show()
    sys.exit(app.exec_())
