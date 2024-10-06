import sys
import pandas as pd
import pyodbc
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QTextEdit
from PyQt5.QtCore import Qt
import re

# Configurações de conexão com o banco de dados SQL Server
server = 'localhost'  # Atualize com seu servidor
database = 'Softm'
username = 'bi'
password = 'bi.softm'
connection_string = f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

class ImportApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Importador de Dados para SQL Server")
        self.setGeometry(200, 200, 600, 400)

        # Botão para selecionar arquivo
        self.select_button = QPushButton("Selecionar Planilha", self)
        self.select_button.setGeometry(50, 50, 200, 40)
        self.select_button.clicked.connect(self.select_file)

        # Botão para iniciar importação
        self.import_button = QPushButton("Importar Dados", self)
        self.import_button.setGeometry(350, 50, 200, 40)
        self.import_button.clicked.connect(self.import_data)

        # Caixa de texto para mostrar status
        self.log = QTextEdit(self)
        self.log.setGeometry(50, 120, 500, 200)
        self.log.setReadOnly(True)

        # Caminho do arquivo selecionado
        self.file_path = None

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

        conn = self.connect_to_database()
        if conn is None:
            self.log.append("Erro ao conectar ao banco de dados.")
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
            tempo = row['Tempo']
            status = row['Status']

            # Conversão de data para um formato aceito pelo SQL Server (YYYY-MM-DD HH:MM:SS)
            try:
                data = pd.to_datetime(data, dayfirst=True).strftime('%Y-%m-%d %H:%M:%S')
            except Exception as e:
                self.log.append(f"Linha {index + 2} (Protocolo {protocolo}): Erro de conversão de data: {e}")
                continue

            # Processar o valor de "Tempo" e ajustar o status
            duracao = None
            if isinstance(tempo, str):
                if "em atendimento" in tempo.lower():
                    duracao = None
                    status = "Em atendimento"
                else:
                    try:
                        # Verifica se o tempo é maior que 24 horas
                        h, m, s = map(int, tempo.split(':'))
                        if h >= 24:
                            # Armazena como string no formato de hora extendida
                            duracao = tempo
                        else:
                            # Armazena no formato padrão
                            duracao = f"{h:02}:{m:02}:{s:02}"
                    except ValueError:
                        self.log.append(f"Linha {index + 2} (Protocolo {protocolo}): Erro no formato do tempo: {tempo}")
                        continue
            else:
                duracao = tempo

            atendente_id, setor_id = self.get_user_id_and_sector_id(conn, atendente_nome, setor_nome)

            if atendente_id is not None and setor_id is not None:
                self.insert_into_atendimento(conn, data, protocolo, atendente_id, setor_id, duracao, status)
                self.log.append(f"Linha {index + 2} (Protocolo {protocolo}): Dados inseridos com sucesso.")
            else:
                self.log.append(f"Linha {index + 2} (Protocolo {protocolo}): Erro - Atendente ou setor não encontrado.")

        conn.close()
        self.log.append("Importação de dados concluída.")

    def connect_to_database(self):
        try:
            conn = pyodbc.connect(connection_string)
            self.log.append("Conexão com o banco de dados realizada com sucesso.")
            return conn
        except Exception as e:
            self.log.append(f"Erro ao conectar ao banco de dados: {e}")
            return None

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
            cursor.execute("""
            INSERT INTO Atendimento (data, protocolo, atendente_id, setor, duracao, status)
            VALUES (?, ?, ?, ?, ?, ?)
            """, data, protocolo, atendente_id, setor_id, duracao, status)
            conn.commit()
        except Exception as e:
            self.log.append(f"Erro ao inserir dados na tabela Atendimento (Protocolo {protocolo}): {e}")
            conn.rollback()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ImportApp()
    window.show()
    sys.exit(app.exec_())
