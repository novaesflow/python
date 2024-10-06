"""
========================================
Nome do Script: Gestor de Checklists - Admin
Autor: Tiego Novaes Santana
Data de Criação: 04/09/2024
Versão: 1.0
Descrição: Este script é utilizado para gerenciar checklists e tarefas usando dados de um banco de dados SQL Server.
Git Repository: https://github.com/novaesflow/python.git
========================================
"""

import sys
import pyodbc
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QTextEdit, QLabel, QTabWidget, QWidget, QVBoxLayout, QGridLayout, QTableWidget, QTableWidgetItem
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer

class AdminApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Gestor de Checklists - Admin")
        self.setGeometry(200, 200, 800, 600)

        # Ícone da janela
        self.setWindowIcon(QIcon('C:\\Python\\Biblioteca de icones\\Checklist\\admin-icon.png'))

        # Tabs
        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        self.create_checklist_tab()
        self.create_submissions_tab()

        # Conectar automaticamente ao servidor
        self.conn = self.connect_to_database()

        # Timer para verificar novas submissões a cada 60 segundos
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_new_submissions)
        self.timer.start(60000)  # Verifica a cada 60 segundos

    def connect_to_database(self):
        server = 'your_server_name_here'  # Substitua pelo endereço do servidor se não estiver local
        database = 'ChecklistDB'
        username = 'sa'
        password = 'N0vaes!@2019'
        connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        
        try:
            conn = pyodbc.connect(connection_string)
            print("Conexão com o banco de dados realizada com sucesso.")
            return conn
        except Exception as e:
            print(f"Erro ao conectar ao banco de dados: {e}")
            return None

    def create_checklist_tab(self):
        checklist_tab = QWidget()
        layout = QVBoxLayout()

        # Layout para criar novo checklist
        grid_layout = QGridLayout()
        self.checklist_name_input = QLineEdit()
        self.checklist_name_input.setPlaceholderText("Nome do Checklist")
        create_checklist_button = QPushButton("Criar Checklist")
        create_checklist_button.clicked.connect(self.create_checklist)
        grid_layout.addWidget(QLabel("Nome do Checklist"), 0, 0)
        grid_layout.addWidget(self.checklist_name_input, 0, 1)
        grid_layout.addWidget(create_checklist_button, 1, 0, 1, 2)
        layout.addLayout(grid_layout)

        # Caixa de texto para mostrar status
        self.log_checklist = QTextEdit()
        self.log_checklist.setReadOnly(True)
        layout.addWidget(self.log_checklist)

        checklist_tab.setLayout(layout)
        self.tabs.addTab(checklist_tab, "Checklists")

    def create_submissions_tab(self):
        submissions_tab = QWidget()
        layout = QVBoxLayout()

        # Tabela para mostrar submissões
        self.submissions_table = QTableWidget()
        self.submissions_table.setColumnCount(4)
        self.submissions_table.setHorizontalHeaderLabels(['ID', 'Checklist', 'Usuário', 'Data de Submissão'])
        layout.addWidget(self.submissions_table)

        submissions_tab.setLayout(layout)
        self.tabs.addTab(submissions_tab, "Submissões")

        self.tabs.currentChanged.connect(self.load_submissions)

    def create_checklist(self):
        if not self.conn:
            self.log_checklist.append("Erro: Não conectado ao banco de dados.")
            return

        checklist_name = self.checklist_name_input.text()
        if not checklist_name:
            self.log_checklist.append("O nome do checklist não pode estar vazio.")
            return

        cursor = self.conn.cursor()
        try:
            cursor.execute("INSERT INTO Checklists (Title, CreatedBy) VALUES (?, ?)", (checklist_name, 1))  # ID do Admin
            self.conn.commit()
            self.log_checklist.append(f"Checklist '{checklist_name}' criado com sucesso.")
        except Exception as e:
            self.log_checklist.append(f"Erro ao criar checklist: {e}")

    def load_submissions(self):
        if self.tabs.currentIndex() == 1:  # Verifica se a aba atual é "Submissões"
            if not self.conn:
                self.log_checklist.append("Erro: Não conectado ao banco de dados.")
                return

            cursor = self.conn.cursor()
            cursor.execute("""
                SELECT s.SubmissionID, c.Title, u.Username, s.SubmittedAt 
                FROM ChecklistSubmissions s
                JOIN Checklists c ON s.ChecklistID = c.ChecklistID
                JOIN Users u ON s.UserID = u.UserID
            """)
            submissions = cursor.fetchall()

            self.submissions_table.setRowCount(len(submissions))

            for i, submission in enumerate(submissions):
                self.submissions_table.setItem(i, 0, QTableWidgetItem(str(submission.SubmissionID)))
                self.submissions_table.setItem(i, 1, QTableWidgetItem(submission.Title))
                self.submissions_table.setItem(i, 2, QTableWidgetItem(submission.Username))
                self.submissions_table.setItem(i, 3, QTableWidgetItem(str(submission.SubmittedAt)))

    def check_new_submissions(self):
        if not self.conn:
            return

        cursor = self.conn.cursor()
        cursor.execute("""
            SELECT COUNT(*)
            FROM ChecklistSubmissions
            WHERE SubmittedAt > DATEADD(MINUTE, -1, GETDATE())  -- Submissões nos últimos 60 segundos
        """)
        result = cursor.fetchone()
        if result[0] > 0:
            self.log_checklist.append("Nova submissão de checklist recebida.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AdminApp()
    window.show()
    sys.exit(app.exec_())
