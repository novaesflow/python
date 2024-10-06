import pandas as pd

# Caminho do arquivo CSV
csv_file_path = '/Users/novaes/Downloads/estoque_inicial.csv'

# Ler o arquivo CSV com diferentes parâmetros
try:
    df = pd.read_csv(csv_file_path, delimiter=';', encoding='utf-8')
except Exception as e:
    df = pd.read_csv(csv_file_path, delimiter=';', encoding='latin1')

# Definir a função para gerar as linhas de INSERT SQL
def generate_insert_sql(row):
    # Converter os valores numéricos de string para float, trocando vírgula por ponto
    estoque = float(row['cdestoque'].replace(',', '.'))
    vlitem = float(row['vlitem'].replace(',', '.'))
    
    return f"""
INSERT INTO [solidcon].[dbo].[estoqueinicial]
  (cdpessoafilial, cdpessoaobra, codigo, cdembalagem, cdestoque, vlitem)
VALUES
  ({row['cdpessoafilial']}, {row['cdpessoaobra']}, {row['codigo']}, '{row['cdembalagem']}', {estoque:.2f}, {vlitem:.2f});
"""

# Gerar todas as linhas de INSERT SQL
sql_statements = df.apply(generate_insert_sql, axis=1)

# Caminho para salvar o novo arquivo SQL
sql_file_path = '/Users/novaes/Downloads/estoqueinicial_inserts.sql'

# Salvar as linhas de INSERT SQL em um arquivo .sql
with open(sql_file_path, 'w') as file:
    for statement in sql_statements:
        file.write(statement)

print(f"Arquivo SQL gerado em: {sql_file_path}")
