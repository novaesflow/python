"""
========================================
Nome do Programa: Flow Utility - Subida de Vendas
Autor: Tiego Novaes Santana
Data de Criação: 10/09/2024
Versão: 1.9
Descrição: Este programa realiza a subida de vendas de um banco de dados migrado para o sistema atual. 
A inserção é feita primeiro em um banco temporário e pode ser validada antes de subir as vendas no banco final Solidcon.
Git Repository: https://github.com/novaesflow/flow-utility.git
========================================
"""

import sys
import pyodbc
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QTextEdit, QLabel, QTabWidget, QWidget, QVBoxLayout, QGridLayout, QHBoxLayout, QRadioButton
from PyQt5.QtCore import Qt
from datetime import datetime, timedelta


class FlowUtility(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Flow Utility - DBA Manager")
        self.setGeometry(200, 200, 800, 600)

        # Tabs
        self.tabs = QTabWidget(self)
        self.setCentralWidget(self.tabs)

        # Criar abas
        self.create_sql_tab()
        self.create_sales_tab()

        # Inicializando conexão SQL
        self.conn = None

        # Desativar segunda aba inicialmente
        self.tabs.setTabEnabled(1, False)

    def create_sql_tab(self):
        sql_tab = QWidget()
        layout = QVBoxLayout()

        grid_layout = QGridLayout()

        # Inputs de conexão
        self.server_input = QLineEdit()
        self.username_input = QLineEdit()
        self.password_input = QLineEdit()

        self.server_input.setPlaceholderText("Servidor")
        self.username_input.setPlaceholderText("Usuário")
        self.password_input.setPlaceholderText("Senha")
        self.password_input.setEchoMode(QLineEdit.Password)

        # Conectar o botão "Enter" para executar a conexão
        self.server_input.returnPressed.connect(self.connect_to_database)
        self.username_input.returnPressed.connect(self.connect_to_database)
        self.password_input.returnPressed.connect(self.connect_to_database)

        grid_layout.addWidget(QLabel("Servidor"), 0, 0)
        grid_layout.addWidget(self.server_input, 0, 1)
        grid_layout.addWidget(QLabel("Usuário"), 1, 0)
        grid_layout.addWidget(self.username_input, 1, 1)
        grid_layout.addWidget(QLabel("Senha"), 2, 0)
        grid_layout.addWidget(self.password_input, 2, 1)

        connect_button = QPushButton("Conectar")
        connect_button.clicked.connect(self.connect_to_database)
        grid_layout.addWidget(connect_button, 3, 0, 1, 2)

        layout.addLayout(grid_layout)
        self.log_sql = QTextEdit()
        self.log_sql.setReadOnly(True)
        layout.addWidget(self.log_sql)

        sql_tab.setLayout(layout)
        self.tabs.addTab(sql_tab, "Conexão SQL")

    def create_sales_tab(self):
        sales_tab = QWidget()
        layout = QVBoxLayout()

        grid_layout = QGridLayout()

        # Campo para selecionar o banco origem
        self.bank_input = QLineEdit()
        self.bank_input.setPlaceholderText("Banco de Origem")
        grid_layout.addWidget(QLabel("Banco de Origem"), 0, 0)
        grid_layout.addWidget(self.bank_input, 0, 1)

        # Campo para o número da filial
        self.filial_input = QLineEdit()
        self.filial_input.setPlaceholderText("Número da Filial")
        grid_layout.addWidget(QLabel("Número da Filial"), 1, 0)
        grid_layout.addWidget(self.filial_input, 1, 1)

        # Campo de data inicial
        self.start_date_input = QLineEdit()
        self.start_date_input.setPlaceholderText("Data Inicial (YYYY-MM-DD)")
        grid_layout.addWidget(QLabel("Data Inicial"), 2, 0)
        grid_layout.addWidget(self.start_date_input, 2, 1)

        # Campo de data final (preenchido automaticamente com a última data)
        self.end_date_input = QLineEdit()
        self.end_date_input.setPlaceholderText("Data Final (YYYY-MM-DD)")
        grid_layout.addWidget(QLabel("Data Final"), 3, 0)
        grid_layout.addWidget(self.end_date_input, 3, 1)

        # Botão para buscar a última data
        get_last_date_button = QPushButton("Buscar Última Data")
        get_last_date_button.clicked.connect(self.get_last_date)
        grid_layout.addWidget(get_last_date_button, 4, 0, 1, 2)

        layout.addLayout(grid_layout)

        # Pergunta para criar o banco temporário
        self.create_temp_prompt = QHBoxLayout()
        self.create_temp_prompt.addWidget(QLabel("Deseja criar o banco SolidconTemporario?"))

        self.radio_temp_yes = QRadioButton("Sim")
        self.radio_temp_no = QRadioButton("Não")
        self.radio_temp_no.setChecked(True)

        self.create_temp_prompt.addWidget(self.radio_temp_yes)
        self.create_temp_prompt.addWidget(self.radio_temp_no)
        layout.addLayout(self.create_temp_prompt)

        self.radio_temp_yes.toggled.connect(self.update_button_label)

        # Botão para criar o banco e subir vendas
        self.upload_button = QPushButton("Subir Vendas para Temporário")
        self.upload_button.setEnabled(False)  # Desabilitado até que o banco seja criado
        self.upload_button.clicked.connect(self.create_temp_and_upload_sales)
        layout.addWidget(self.upload_button)

        # Exibe quantos registros foram inseridos no banco temporário
        self.records_label = QLabel("Registros inseridos no banco temporário: 0")
        layout.addWidget(self.records_label)

        # Pergunta sobre a validação no banco temporário
        self.validation_prompt = QHBoxLayout()
        self.validation_prompt.addWidget(QLabel("Você validou a subida de vendas no banco temporário?"))

        self.radio_yes = QRadioButton("Sim")
        self.radio_no = QRadioButton("Não")
        self.radio_no.setChecked(True)

        self.validation_prompt.addWidget(self.radio_yes)
        self.validation_prompt.addWidget(self.radio_no)
        layout.addLayout(self.validation_prompt)

        self.radio_yes.toggled.connect(self.enable_solidcon_button)

        # Botão para subir vendas para o Solidcon (desativado por padrão)
        self.upload_to_solidcon_button = QPushButton("Subir Vendas para o Solidcon e Apagar Temporário")
        self.upload_to_solidcon_button.setEnabled(False)
        self.upload_to_solidcon_button.clicked.connect(self.move_to_solidcon)
        layout.addWidget(self.upload_to_solidcon_button)

        # Exibe quantos registros foram inseridos no banco Solidcon
        self.solidcon_records_label = QLabel("Registros inseridos no banco Solidcon: 0")
        layout.addWidget(self.solidcon_records_label)

        # Mensagem final
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        layout.addWidget(self.log)

        sales_tab.setLayout(layout)
        self.tabs.addTab(sales_tab, "Subida de Vendas")

    def connect_to_database(self):
        server = self.server_input.text()
        username = self.username_input.text()
        password = self.password_input.text()

        connection_string = f'DRIVER={{SQL Server}};SERVER={server};UID={username};PWD={password}'

        try:
            if self.conn is None or not self.conn:
                self.conn = pyodbc.connect(connection_string)
                self.log_sql.append("Conexão com o banco de dados realizada com sucesso.")
                self.tabs.setTabEnabled(1, True)  # Ativar segunda aba após a conexão
            else:
                self.log_sql.append("Conexão já está ativa.")
        except Exception as e:
            self.log_sql.append(f"Erro ao conectar ao banco de dados: {e}")

    def get_last_date(self):
        if self.conn is None:
            self.log.append("Por favor, conecte-se ao banco de dados primeiro.")
            return

        selected_bank = self.bank_input.text()

        # Otimização: consulta sem índice inexistente
        query = f"""
            SELECT MAX(TRNDAT)
            FROM {selected_bank}.dbo.ITEVDA
        """

        cursor = self.conn.cursor()

        try:
            self.log.append(f"Consultando última data na tabela {selected_bank}.dbo.ITEVDA...")

            cursor.execute(query)
            result = cursor.fetchone()

            if result and result[0]:
                last_date = result[0]
                # Preenchendo o campo de data final com a última data
                self.end_date_input.setText(last_date.strftime('%Y-%m-%d'))

                # Preenchendo o campo de data inicial com um ano antes da última data
                one_year_before = last_date - timedelta(days=365)
                self.start_date_input.setText(one_year_before.strftime('%Y-%m-%d'))

                self.log.append(f"Última data encontrada: {last_date}")
            else:
                self.log.append("Não foi possível encontrar uma data válida.")
        except Exception as e:
            self.log.append(f"Erro ao buscar a última data: {e}")
        finally:
            cursor.close()

    def update_button_label(self):
        if self.radio_temp_yes.isChecked():
            self.upload_button.setText("Criar banco temporário e subir vendas")
            self.upload_button.setEnabled(True)
        else:
            self.upload_button.setText("Subir Vendas para Temporário")
            self.upload_button.setEnabled(False)

    def create_temp_and_upload_sales(self):
        try:
            # Primeiro, criamos o banco de dados fora de qualquer transação
            self.log.append("Criando o banco de dados SolidconTemporario...")
            self.create_temporary_database()

            # Agora podemos prosseguir com as outras operações, como criar a tabela e inserir dados
            cursor = self.conn.cursor()

            # Criando a tabela tbVendaPDV no banco SolidconTemporario
            cursor.execute("""
            USE [SolidconTemporario];
            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tbVendaPDV')
            BEGIN
                CREATE TABLE [dbo].[tbVendaPDV](
                    [cdPessoaFilial] [int] NOT NULL,
                    [cdVendaPDV] [bigint] IDENTITY(1,1) NOT NULL,
                    [dtVenda] [datetime] NULL,
                    [cdProduto] [int] NULL,
                    [cdEmpresaProduto] [tinyint] NULL,
                    [cdVenda] [bigint] NULL,
                    [qtVenda] [decimal](14, 4) NULL,
                    [vlVenda] [money] NULL,
                    [vlCusto] [money] NULL,
                    [cdPromocao] [int] NULL,
                    [vlVerbaComercial] [decimal](18, 8) NULL,
                    [vlSellOut] [decimal](18, 8) NULL,
                    [qtDescontoPromocaoCombinada] [decimal](18, 3) NULL,
                    [vlDescontoPromocaoCombinada] [decimal](18, 4) NULL,
                    [vlCustoLogistico] [money] NULL,
                    [vlQuebra] [money] NULL,
                    [vlPIS] [money] NULL,
                    [vlCOFINS] [money] NULL,
                    [vlICMS] [money] NULL,
                    [vlMercadoria] [money] NULL,
                    [qtCupons] [int] NULL,
                    [vlICMSVenda] [money] NULL,
                    [vlPISVenda] [money] NULL,
                    [vlCOFINSVenda] [money] NULL,
                    [vlCustoOriginal] [money] NULL,
                 CONSTRAINT [PK__tbVendaPDV] PRIMARY KEY NONCLUSTERED 
                (
                    [cdVendaPDV] ASC,
                    [cdPessoaFilial] ASC
                )
                );
            END;
            """)
            self.conn.commit()  # Confirma a criação da tabela

            self.log.append("Tabela tbVendaPDV criada com sucesso no banco SolidconTemporario.")

            # Agora, subindo as vendas para o banco temporário
            self.upload_sales_temporario()

        except Exception as e:
            self.log.append(f"Erro ao criar o banco de dados temporário e subir vendas: {e}")

    def create_temporary_database(self):
        try:
            # Abrir uma nova conexão sem transações ativas para criar o banco
            connection_string = f'DRIVER={{SQL Server}};SERVER={self.server_input.text()};UID={self.username_input.text()};PWD={self.password_input.text()}'
            temp_conn = pyodbc.connect(connection_string)
            cursor = temp_conn.cursor()

            # Criando o banco de dados temporário fora de transações
            cursor.execute("""
            IF NOT EXISTS(SELECT * FROM sys.databases WHERE name = 'SolidconTemporario')
            BEGIN
                CREATE DATABASE [SolidconTemporario];
            END;
            """)
            temp_conn.commit()  # Confirma a criação do banco de dados imediatamente

            self.log.append("Banco de dados SolidconTemporario criado com sucesso.")
            cursor.close()
            temp_conn.close()
        except Exception as e:
            self.log.append(f"Erro ao criar o banco de dados SolidconTemporario: {e}")

    def upload_sales_temporario(self):
        selected_bank = self.bank_input.text()
        filial_number = self.filial_input.text()
        start_date = self.start_date_input.text()
        end_date = self.end_date_input.text()

        try:
            cursor = self.conn.cursor()

            # Garantindo que estamos no contexto correto do banco temporário
            cursor.execute("USE SolidconTemporario;")
            self.log.append("Contexto de banco de dados definido para SolidconTemporario.")

            # Verificar se o trigger existe antes de tentar desabilitá-lo
            cursor.execute("""
            IF EXISTS (SELECT * FROM sys.triggers WHERE name = 'tgmovimentaestoque')
            BEGIN
                ALTER TABLE tbvendapdv DISABLE TRIGGER tgmovimentaestoque;
            END
            """)
            self.log.append("Trigger de movimentação de estoque desabilitado.")

            # Inserção de dados no banco temporário
            query = f"""
            INSERT INTO [SolidconTemporario].[dbo].[tbVendaPDV]
                   ([cdPessoaFilial], [dtVenda], [cdProduto], [cdEmpresaProduto], [cdVenda], [qtVenda], [vlVenda], [vlCusto], [cdPromocao])
            SELECT {filial_number} AS cdpessoafilial, 
                   cast({selected_bank}.dbo.itevda.TRNDAT as date) AS dtvenda,
                   tbproduto.cdProduto, 10 as cdempresaproduto, cdProdutoVenda as cdvenda,
                   {selected_bank}.dbo.itevda.ITVQTDVDA as Qtvenda,  
                   {selected_bank}.dbo.itevda.ITVVLRUNI as vlvenda, 
                   {selected_bank}.dbo.itevda.ITVPRCCST as Vlcusto, 0 as cdpromocao
            FROM {selected_bank}.dbo.ITEVDA
            INNER JOIN {selected_bank}.dbo.tbProdutoVenda ON {selected_bank}.dbo.itevda.PROCOD = {selected_bank}.dbo.tbProdutoVenda.cdEAN
            INNER JOIN {selected_bank}.dbo.tbProduto ON {selected_bank}.dbo.tbProdutoVenda.cdProduto = {selected_bank}.dbo.tbProduto.cdProduto 
            AND {selected_bank}.dbo.tbProdutoVenda.cdEmpresa = {selected_bank}.dbo.tbProduto.cdEmpresa
            WHERE {selected_bank}.dbo.itevda.itvtip = '1' 
            AND {selected_bank}.dbo.itevda.trndat BETWEEN '{start_date}' AND '{end_date}'
            """
            cursor.execute(query)
            self.conn.commit()
            self.log.append("Vendas inseridas com sucesso no banco temporário.")

            # Contagem de registros inseridos no banco temporário
            cursor.execute("SELECT COUNT(*) FROM [SolidconTemporario].[dbo].[tbVendaPDV]")
            record_count = cursor.fetchone()[0]
            self.records_label.setText(f"Registros inseridos no banco temporário: {record_count}")
            self.log.append(f"{record_count} registros inseridos no banco temporário.")

        except Exception as e:
            self.log.append(f"Erro ao realizar a subida de vendas: {e}")
            if cursor.messages:
                for message in cursor.messages:
                    self.log.append(f"Mensagem SQL: {message}")
            self.conn.rollback()

        finally:
            cursor.close()

    def enable_solidcon_button(self):
        if self.radio_yes.isChecked():
            self.upload_to_solidcon_button.setEnabled(True)
        else:
            self.upload_to_solidcon_button.setEnabled(False)

    def move_to_solidcon(self):
        query_move = """
        USE [Solidcon];
        INSERT INTO [dbo].[tbVendaPDV] SELECT * FROM SolidconTemporario.dbo.tbVendaPDV;
        DROP DATABASE [SolidconTemporario];
        """
        try:
            cursor = self.conn.cursor()
            cursor.execute(query_move)
            self.conn.commit()

            cursor.execute("SELECT COUNT(*) FROM [Solidcon].[dbo].[tbVendaPDV]")
            solidcon_record_count = cursor.fetchone()[0]
            self.solidcon_records_label.setText(f"Registros inseridos no banco Solidcon: {solidcon_record_count}")

            self.log.append("Dados movidos para o banco Solidcon e banco temporário apagado.")
        except Exception as e:
            self.log.append(f"Erro ao mover dados para o banco Solidcon: {e}")
            self.conn.rollback()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FlowUtility()
    window.show()
    sys.exit(app.exec_())
