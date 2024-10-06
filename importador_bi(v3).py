"""
========================================
Nome do Script: Importador de Dados para SQL Server
Autor: Tiego Novaes Santana
Data de Criação: 01/09/2024
Versão: 1.3
Descrição: Este script é utilizado para importar dados de uma planilha Excel para um banco de dados SQL Server, afim de utilizar o PowerBI
Git Repository: https://github.com/novaesflow/python.git
========================================
"""
import sys
import pandas as pd
import pyodbc
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit, QLabel, QLineEdit, QTabWidget, QWidget, QVBoxLayout, QGridLayout, QHBoxLayout
from PyQt5.QtCore import Qt
import re

class ImportApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Importador de Dados para SQL Server")
        self.setGeometry(200, 200, 800, 600)

        # Tabs
        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        self.create_sql_tab()
        self.create_atendimento_tab()
        self.create_horarios_tab()

        # Inicializando conexão SQL
        self.conn = None

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

    def create_atendimento_tab(self):
        atendimento_tab = QWidget()
        layout = QVBoxLayout()

        # Botão para selecionar arquivo
        self.select_button = QPushButton("Selecionar Planilha")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        # Botão para iniciar importação
        self.import_button = QPushButton("Importar Dados")
        self.import_button.clicked.connect(self.import_data)
        layout.addWidget(self.import_button)

        # Caixa de texto para mostrar status
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        atendimento_tab.setLayout(layout)
        self.tabs.addTab(atendimento_tab, "Atendimento")

    def create_horarios_tab(self):
        horarios_tab = QWidget()
        layout = QVBoxLayout()

        # Similar interface para futura funcionalidade de "Horários"
        select_button = QPushButton("Selecionar Planilha")
        layout.addWidget(select_button)

        import_button = QPushButton("Importar Dados")
        layout.addWidget(import_button)

        log = QTextEdit()
        log.setReadOnly(True)
        layout.addWidget(log)

        horarios_tab.setLayout(layout)
        self.tabs.addTab(horarios_tab, "Horários")

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

    def select_file(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Planilha Excel", "", "Excel Files (*.xlsx *.xls)", options=options)
        if file_name:
            self.file_path = file_name
            self.log.append(f"Arquivo selecionado: {file_name}")

    def import_data(self):
        if not self.file_path:
            self.log.append("Por favor, selecione um arquivo primeiro.")
            return

        if self.conn is None:
            self.log.append("Por favor, conecte-se ao banco de dados primeiro.")
            return

        df = self.read_excel_file(self.file_path)
        if df is None:
            self.log.append("Erro ao ler a planilha Excel.")
            return

        required_columns = ['Data de Criação', 'Protocolo', 'Nome do Usuário', 'Setor', 'Tempo', 'Status']
        if not all(col in df.columns for col in required_columns):
            self.log.append("Erro: A planilha Excel não contém todas as colunas necessárias.")
            self.log.append(f"Colunas encontradas: {', '.join(df.columns)}")
            self.log.append(f"Colunas necessárias: {', '.join(required_columns)}")
            return

        for index, row in df.iterrows():
            data = row['Data de Criação']
            protocolo = row['Protocolo']
            atendente_nome = row['Nome do Usuário'].split()[0]  # Pega apenas o primeiro nome
            setor_nome = row['Setor']
            duracao = row['Tempo'] if re.match(r'^\d{2}:\d{2}:\d{2}$', str(row['Tempo'])) else None
            status = row['Status']

            atendente_id, setor_id = self.get_user_id_and_sector_id(self.conn, atendente_nome, setor_nome)

            if atendente_id is not None and setor_id is not None:
                self.insert_into_atendimento(self.conn, data, protocolo, atendente_id, setor_id, duracao, status)
                self.log.append(f"Linha {index + 2}: Dados inseridos com sucesso.")
            else:
                self.log.append(f"Linha {index + 2}: Erro - Atendente ou setor não encontrado.")

        self.log.append("Importação de dados concluída.")

    def read_excel_file(self, file_path):
        try:
            df = pd.read_excel(file_path)
            self.log.append("Dados da planilha Excel lidos com sucesso.")
            return df
        except Exception as e:
            self.log.append(f"Erro ao ler a planilha Excel: {e}")
            return None

    def get_user_id_and_sector_id(self, conn, atendente_nome, setor_nome):
        cursor = conn.cursor()
        try:
            # Buscar atendente pelo primeiro nome
            cursor.execute("SELECT id FROM Usuarios WHERE nome LIKE ?", f'{atendente_nome}%')
            user_result = cursor.fetchone()
            atendente_id = user_result[0] if user_result else None

            # Ajuste a correspondência de setores para lidar com espaços adicionais e diferenças mínimas
            cursor.execute("SELECT id, nome FROM Setores")
            setores = cursor.fetchall()
            setor_id = None
            for setor in setores:
                if re.sub(r'\s+', '', setor_nome.lower()) == re.sub(r'\s+', '', setor.nome.lower()):
                    setor_id = setor.id
                    break
            
        except pyodbc.Error as e:
            self.log.append(f"Erro ao buscar IDs de usuário ou setor: {e}")
            atendente_id = None
            setor_id = None
        
        return atendente_id, setor_id

    def insert_into_atendimento(self, conn, data, protocolo, atendente_id, setor_id, duracao, status):
        cursor = conn.cursor()
        try:
            # Certifique-se de que 'data' está no formato correto antes de usá-la
            data_formatada = pd.to_datetime(data, dayfirst=True).strftime('%Y-%m-%d %H:%M:%S')
            
            cursor.execute("""
            INSERT INTO Atendimento (data, protocolo, atendente_id, setor, duracao, status)
            VALUES (?, ?, ?, ?, ?, ?)
            """, data_formatada, protocolo, atendente_id, setor_id, duracao, status)
            conn.commit()
        except Exception as e:
            self.log.append(f"Erro ao inserir dados na tabela Atendimento (Protocolo {protocolo}): {e}")
            conn.rollback()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ImportApp()
    window.show()
    sys.exit(app.exec_())

