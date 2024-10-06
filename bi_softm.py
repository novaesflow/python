"""
================================================================================================================================================================
Nome do Script: Dashboard de Atendimentos SoftM
Autor: Tiego Novaes Santana
Empresa: Flow Optimize®
Data de Criação: 06/09/2024
Versão: 2.1
Descrição: Este script cria um dashboard interativo com KPIs, gráficos de evolução de atendimentos,
situação dos atendimentos e desempenho dos colaboradores para o sistema SoftM.
Ele utiliza dados de um banco SQL Server para geração de métricas e gráficos.
Git Repository: https://github.com/novaesflow/python.git
================================================================================================================================================================
"""
import dash
from dash import dcc, html
import dash_bootstrap_components as dbc
from dash.dependencies import Input, Output
import plotly.graph_objs as go
import pandas as pd
import pyodbc
import numpy as np
from flask_caching import Cache

# Conexão com o banco de dados SQL Server (credenciais fornecidas)
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=localhost;'
    r'DATABASE=softm;'
    r'UID=bi;'  # Usuário fornecido
    r'PWD=bi.softm;'  # Senha fornecida
)
conn = pyodbc.connect(conn_str)

# Carregar os dados da tabela Resumo
query_resumo = "SELECT * FROM Resumo"
df_resumo = pd.read_sql(query_resumo, conn)

# Convertendo a coluna Competencia para datetime
df_resumo['Competencia'] = pd.to_datetime(df_resumo['Competencia'], format='%Y-%m')

# Função para gerar os gráficos de evolução de atendimentos (mostra todos os períodos)
def gerar_grafico_evolucao(df):
    evolucao_fig = go.Figure()

    # Adicionar a série completa de evolução dos atendimentos
    evolucao_fig.add_trace(go.Scatter(
        x=df['Competencia'],
        y=df['Atendimentos'],
        mode='lines+markers',
        line=dict(color='blue', width=2),
        marker=dict(size=8)
    ))

    # Melhorias no layout do gráfico
    evolucao_fig.update_layout(
        title='Evolução do Número de Atendimentos',
        xaxis_title='Competência (Mês/Ano)',
        yaxis_title='Número de Atendimentos',
        height=300,
        xaxis=dict(tickformat="%b %Y"),  # Formato Mês e Ano no eixo X
        yaxis=dict(tickformat=".0f",  # Remove decimais do eixo Y
                   rangemode="tozero"),  # Iniciar eixo Y no zero
        margin=dict(l=40, r=40, t=60, b=40),
        template='plotly_white'
    )

    return evolucao_fig

# Função para gerar gráficos de situação dos atendimentos
def gerar_grafico_situacao(df):
    situacao_fig = go.Figure(data=[
        go.Bar(name='Cancelado', x=['Cancelado'], y=[30], marker=dict(color='red')),
        go.Bar(name='Em Aberto', x=['Em Aberto'], y=[70], marker=dict(color='orange')),
        go.Bar(name='Atendido', x=['Atendido'], y=[250], marker=dict(color='green'))
    ])
    situacao_fig.update_layout(
        title='Situação dos Atendimentos',
        barmode='stack',
        height=300,
        template='plotly_white'
    )
    return situacao_fig

# Função para gerar o gráfico TOP 5 Colaboradores por Atendimentos
def gerar_grafico_colaboradores():
    top_colaboradores_fig = go.Figure(data=[
        go.Bar(name='Colaborador 1', y=['Colaborador 1'], x=[120], orientation='h'),
        go.Bar(name='Colaborador 2', y=['Colaborador 2'], x=[150], orientation='h'),
        go.Bar(name='Colaborador 3', y=['Colaborador 3'], x=[100], orientation='h'),
        go.Bar(name='Colaborador 4', y=['Colaborador 4'], x=[200], orientation='h'),
        go.Bar(name='Colaborador 5', y=['Colaborador 5'], x=[250], orientation='h')
    ])
    top_colaboradores_fig.update_layout(
        title='Top 5 Colaboradores por Atendimentos',
        height=400,
        template='plotly_white'
    )
    return top_colaboradores_fig

# Função para desenhar KPIs e setas
def desenhar_kpis(df, competencia_selecionada):
    df_selecionado = df[df['Competencia'] == competencia_selecionada]
    
    if df_selecionado.empty:
        numero_atendimentos = "Sem dados"
        tempo_medio = "Sem dados"
        comparativo_atendimentos = "Sem dados"
        comparativo_tempo = "Sem dados"
    else:
        numero_atendimentos = int(df_selecionado['Atendimentos'].values[0])
        tempo_medio = df_selecionado['Tempo_medio'].values[0]
        
        # Garantir que os comparativos sejam números (convertendo strings para float)
        try:
            comparativo_atendimentos = float(df_selecionado['Comparativo_atendimento'].values[0])
        except ValueError:
            comparativo_atendimentos = 0.0  # Valor padrão se não for possível converter
        
        try:
            comparativo_tempo = float(df_selecionado['Comparativo_tempo'].values[0])
        except ValueError:
            comparativo_tempo = 0.0  # Valor padrão se não for possível converter
    
    # Adicionar as setas de comparativo (Verde para queda, Vermelha para aumento) com FontAwesome
    seta_atendimentos = html.I(className="fas fa-arrow-down", style={'color': 'green'}) if comparativo_atendimentos < 0 else html.I(className="fas fa-arrow-up", style={'color': 'red'})
    seta_tempo = html.I(className="fas fa-arrow-down", style={'color': 'green'}) if comparativo_tempo < 0 else html.I(className="fas fa-arrow-up", style={'color': 'red'})
    
    return [
        dbc.Col(dbc.Card([
            dbc.CardHeader("Número de Atendimentos"),
            dbc.CardBody(html.H1([f'{numero_atendimentos} ', seta_atendimentos]))
        ], color="primary", inverse=True), width=3),
        
        dbc.Col(dbc.Card([
            dbc.CardHeader("Tempo Médio de Atendimento"),
            dbc.CardBody(html.H1([f'{tempo_medio} ', seta_tempo]))
        ], color="secondary", inverse=True), width=3),
        
        dbc.Col(dbc.Card([
            dbc.CardHeader("Avaliação de Atendimento"),
            dbc.CardBody(html.H1("9.5"))  # Indicador fictício
        ], color="success", inverse=True), width=3),
        
        dbc.Col(dbc.Card([
            dbc.CardHeader("Pontualidade"),
            dbc.CardBody(html.H1("95%"))  # Indicador fictício
        ], color="info", inverse=True), width=3)
    ]

# Instância da aplicação Dash com Bootstrap e FontAwesome
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.SUPERHERO, "https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css"])
server = app.server

# Cache de dados para evitar sobrecarregar o banco de dados
cache = Cache(app.server, config={'CACHE_TYPE': 'simple'})

# Layout da aplicação
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col(html.H1("Dashboard de Atendimentos - SoftM", className='text-center text-primary mb-4'), width=12)
    ]),
    
    dbc.Row(dbc.Col(dcc.Dropdown(
        id='competencia-selecionada',
        options=[{'label': str(competencia)[:7], 'value': competencia} for competencia in df_resumo['Competencia'].unique()],
        value=str(df_resumo['Competencia'].max())[:7],  # Valor inicial
        className='mb-4'
    ), width=12)),
    
    dbc.Row(id='kpis', className='mb-4 justify-content-center'),
    
    dbc.Row([
        dbc.Col(dcc.Graph(id='evolucao-grafico'), width=6),
        dbc.Col(dcc.Graph(id='situacao-grafico'), width=6)
    ]),
    
    dbc.Row(dbc.Col(dcc.Graph(id='top-colaboradores-grafico'), width=12, className="mt-4"))
], fluid=True)

# Callback para atualizar os KPIs e gráficos com base na competência selecionada
@app.callback(
    [Output('kpis', 'children'),
     Output('evolucao-grafico', 'figure'),
     Output('situacao-grafico', 'figure'),
     Output('top-colaboradores-grafico', 'figure')],
    [Input('competencia-selecionada', 'value')]
)
def atualizar_dashboard(competencia_selecionada):
    competencia_selecionada = pd.to_datetime(competencia_selecionada)
    
    # Atualizando KPIs e gráficos
    kpis = desenhar_kpis(df_resumo, competencia_selecionada)
    evolucao_grafico = gerar_grafico_evolucao(df_resumo)  # Agora mostra todos os períodos
    situacao_grafico = gerar_grafico_situacao(df_resumo)
    top_colaboradores_grafico = gerar_grafico_colaboradores()
    
    return kpis, evolucao_grafico, situacao_grafico, top_colaboradores_grafico

# Rodar o servidor
if __name__ == "__main__":
    app.run_server(debug=True)