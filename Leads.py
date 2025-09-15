import dash
from dash import dcc, html, dash_table, Input, Output, callback, State, ALL
import json
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import base64
import hashlib

# Configuração do Google Sheets
def get_google_sheets_data():
    """
    Função para conectar ao Google Sheets e buscar dados da aba 'Resumo'
    Usando export CSV sem necessidade de API ou credenciais
    """
    
    # URL para exportar a planilha como CSV (primeira aba por padrão)
    #csv_url = "https://docs.google.com/spreadsheets/d/1f_SGg-gpfMOyo3q46dw5QHCTIr0HKyXa9KGYSLZQMV0/export?format=csv" # planilha teste na conta Alex
    csv_url = "https://docs.google.com/spreadsheets/d/1Xgt603LlZMQcj1AXrIf4mdRAxDcQbWz9t9H556XgbZ4/export?format=csv" # planilha em produção na conta Gisélia
               
    
    # Ler dados diretamente do CSV
    df = pd.read_csv(csv_url)
    
    # Normalizar nomes de colunas: strip e mapear variações comuns
    df.columns = df.columns.str.strip()
    # Mapear nomes diferentes para os nomes usados no código
    col_map = {}
    if 'Consultora ' in df.columns:
        col_map['Consultora '] = 'Consultora'
    if 'Consultora' in df.columns:
        col_map['Consultora'] = 'Consultora'
    if 'Historico' in df.columns:
        col_map['Historico'] = 'Histórico'
    if 'Histórico' in df.columns:
        col_map['Histórico'] = 'Histórico'
    if 'NomeAba' in df.columns:
        col_map['NomeAba'] = 'Nomeaba'
    # Aplicar mapeamento somente para colunas existentes
    existing_map = {k: v for k, v in col_map.items() if k in df.columns}
    if existing_map:
        df = df.rename(columns=existing_map)
    
    # Imprimir colunas disponíveis para debug
    print("Colunas disponíveis na planilha:")
    print(df.columns.tolist())
    print(f"Shape do DataFrame: {df.shape}")
    
    # Converter DataReferencia para datetime, assumindo formato dd/mm/yyyy
    if 'DataReferencia' in df.columns:
        df['DataReferencia'] = pd.to_datetime(df['DataReferencia'], format='%d/%m/%Y', errors='coerce')
    
    return df


def render_data_debug_sample():
    """Renderiza um pequeno painel com colunas detectadas e sample dos primeiros registros."""
    try:
        cols = df.columns.tolist()
        sample_records = df.head(5).to_dict('records') if not df.empty else []
        return html.Div([
            html.H5('Painel de debug dos dados (colunas e sample):', style={'marginTop': '10px'}),
            html.Div([html.Strong('Colunas detectadas:')]),
            html.Pre(str(cols), style={'whiteSpace': 'pre-wrap', 'textAlign': 'left', 'backgroundColor': '#f8f8f8', 'padding': '8px'}),
            html.Div([html.Strong('Sample (até 5 linhas):')]),
            html.Pre(str(sample_records), style={'whiteSpace': 'pre-wrap', 'textAlign': 'left', 'backgroundColor': '#f8f8f8', 'padding': '8px'})
        ], style={'textAlign': 'left', 'maxWidth': '900px', 'margin': '0 auto', 'color': '#333'})
    except Exception as e:
        return html.Div(f'Erro ao renderizar debug: {e}')


def auto_map_columns(df_local: pd.DataFrame):
    """Tenta mapear automaticamente colunas com nomes comuns para os nomes esperados no app.
    Retorna (df_mapped, mapping_applied)
    """
    mapping = {}

    # regras de mapeamento conhecidas (chave = existente, valor = desejado)
    candidates = {
        'Contato via WhatsApp': 'Contato via wp',
        'Contato via Whatsapp': 'Contato via wp',
        'Contato via wp': 'Contato via wp',
        'ContatoWP': 'Contato via wp',
        'Historico': 'Histórico',
        'historico': 'Histórico',
        'Histórico': 'Histórico',
        'Data de Referencia': 'DataReferencia',
        'Data Referencia': 'DataReferencia',
        'DataReferencia': 'DataReferencia',
        'Proposta Enviada': 'Proposta',
        'Proposta': 'Proposta',
        'Qualificado?': 'Qualificado',
        'Qualificado': 'Qualificado',
        'Positivo?': 'Positivo',
        'Positivo': 'Positivo',
        'Consultora ': 'Consultora',
        'consultora': 'Consultora'
    }

    # Checar colunas do df e aplicar mapping se a chave existir
    for col in df_local.columns:
        c = col.strip()
        if c in candidates and candidates[c] != c:
            mapping[col] = candidates[c]
        elif c in candidates and candidates[c] == c:
            # mapeamento identidade — ainda assim padronizar strip
            mapping[col] = candidates[c]

    # Aplicar mapeamento
    if mapping:
        df_local = df_local.rename(columns=mapping)

    return df_local, mapping


def find_col(df_local, candidates):
    """Retorna o nome da primeira coluna presente em df_local que case com qualquer string em candidates.
    candidates pode ser lista de opções (case-insensitive compare trimmed).
    """
    cols = [c.strip().lower() for c in df_local.columns]
    for cand in candidates:
        cand_norm = cand.strip().lower()
        for i, c in enumerate(cols):
            if c == cand_norm:
                return df_local.columns[i]
    # tentativas parciais: contains
    for cand in candidates:
        cand_norm = cand.strip().lower()
        for i, c in enumerate(cols):
            if cand_norm in c or c in cand_norm:
                return df_local.columns[i]
    return None


# run_auto_map callback will be defined after `app` is created to avoid referencing
# `app` before it's defined.

# Carregar dados
df = get_google_sheets_data()

# Inicializar app Dash
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "CSLeads"

# Layout principal
app.layout = html.Div([
    html.H1("CSLeads", style={'textAlign': 'center', 'marginBottom': 30}),
    
    # Botão para recarregar dados manualmente
    html.Div([
    html.Button("Atualizar dados", id='refresh-data', n_clicks=0, style={'marginRight': '10px'}),
    #html.Button("Auto-map", id='auto-map', n_clicks=0, style={'marginRight': '10px'}),
    #html.Span(id='auto-map-result', style={'fontSize': '0.9em', 'color': '#666', 'marginRight': '10px'}),
    html.Span(id='last-refresh', style={'fontSize': '0.9em', 'color': '#666'})
    ], style={'textAlign': 'center', 'marginBottom': '10px'}),

    dcc.Tabs(id="tabs", value='tab-1', children=[
        dcc.Tab(label='Análise Geral', value='tab-1', style={'fontSize': '18px', 'fontWeight': '600'}),
        # Qualidade tab disabled per user request
        # dcc.Tab(label='Qualidade', value='tab-qualidade'),
        # Performance tab disabled per user request
        # dcc.Tab(label='Performance', value='tab-performance'),
        dcc.Tab(label='Contatos em Atraso', value='tab-2'),
    ]),
    
    html.Div(id='tab-content'),

    # Sinal oculto para disparar callbacks quando os dados forem recarregados
    html.Div(id='refresh-signal', style={'display': 'none'})
    ,
    # Store para consultora selecionada (ao clicar no card)
    dcc.Store(id='selected-consultora', data='')
    ,
    # Store para período selecionado na aba Análise Geral (permitir que outras abas leiam)
    dcc.Store(id='period-store', data={
        'start': (df['DataReferencia'].min().strftime('%Y-%m-%d') if ('DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty) else None),
        'end': (df['DataReferencia'].max().strftime('%Y-%m-%d') if ('DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty) else None)
    })
])


@app.callback(
    Output('auto-map-result', 'children'),
    Input('auto-map', 'n_clicks'),
    prevent_initial_call=True
)
def run_auto_map(n_clicks):
    if n_clicks and n_clicks > 0:
        global df
        before_cols = df.columns.tolist()
        df, mapping = auto_map_columns(df)
        after_cols = df.columns.tolist()
        if mapping:
            return f"Auto-map aplicado: {mapping}"
        else:
            return "Auto-map não encontrou colunas a renomear"
    raise dash.exceptions.PreventUpdate

# Layout da Aba 1 - Análise Geral
def create_tab1_layout():
    # usar a última data disponível em df como padrão
    max_date = None
    if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
        max_dt = df['DataReferencia'].dropna().max()
        try:
            max_date = max_dt.date()
        except Exception:
            max_date = pd.to_datetime(max_dt).date()
    else:
        today = date.today()
        max_date = today

    return html.Div([
        html.Br(),
        
        # Período de prospecção: label e datepickers na mesma linha (topo esquerdo)
        html.Div([
            html.Div("Período de prospecção:", style={'fontWeight': 'bold', 'marginRight': '12px'}),
            html.Div([
                html.Div('De', style={'fontSize': '12px', 'marginBottom': '4px'}),
                dcc.DatePickerSingle(
                    id='date-picker-start',
                    date=max_date,
                    display_format='DD/MM/YYYY',
                    persistence=True,
                    persistence_type='session'
                )
            ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'flex-start'}),
            html.Div([
                html.Div('Até', style={'fontSize': '12px', 'marginBottom': '4px'}),
                dcc.DatePickerSingle(
                    id='date-picker-end',
                    date=max_date,
                    display_format='DD/MM/YYYY',
                    persistence=True,
                    persistence_type='session'
                )
            ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'flex-start'})
        ], style={'display': 'flex', 'alignItems': 'center', 'gap': '16px', 'justifyContent': 'flex-start', 'marginBottom': '12px'}),
        
        html.Br(),
        html.Br(),
        
        # Cards de resumo
        html.Div(id='summary-cards'),
        
        html.Br(),
        
        # Título centralizado e gráfico de pizza
        html.H3("Distribuição de Qualificados", style={'textAlign': 'center', 'marginTop': '10px'}),
        html.Div([
            dcc.Graph(id='qualificado-pie-chart')
        ]),
        
        html.Br(),
        
        # Tabela resumo por dia
        html.Div([
            html.H3("Resumo por Dia", style={'textAlign': 'center'}),
            html.Div([
                dash_table.DataTable(
                    id='daily-summary-data-table',
                    data=[],
                    columns=[],
                    sort_action='custom',
                    sort_mode='single',
                    sort_by=[{'column_id': 'DATA', 'direction': 'desc'}],
                    style_cell={'textAlign': 'center'},
                    style_header={'backgroundColor': 'lightblue', 'fontWeight': 'bold'}
                )
            ], id='daily-summary-table')
        ])
    ])

# Layout da Aba 2 - Contatos em Atraso
def create_tab2_layout():
    return html.Div([
        html.Br(),
        
        # Configuração de dias para atraso (centralizado)
        html.Div([
            html.Div([
                html.Label("Dias para considerar contato em atraso:", style={'fontSize': '16px', 'fontWeight': '600', 'marginRight': '8px'}),
                dcc.Input(
                    id='days-overdue-config',
                    type='number',
                    value=7,
                    min=1,
                    persistence=True,
                    persistence_type='session',
                    style={'width': '100px', 'textAlign': 'center'}
                )
            ], style={'display': 'flex', 'alignItems': 'center', 'justifyContent': 'center'})
        ], style={'marginBottom': '20px', 'textAlign': 'center'}),

        # Cards de contatos em atraso por consultora
        html.Div(id='overdue-cards'),
        
        html.Br(),
        
        # Tabela de contatos em atraso
        html.Div([
            html.H3("Contatos em Atraso"),
            html.Div(id='overdue-table')
        ])
    ])


# Layout da Aba Qualidade
def create_tab_qualidade_layout():
    return html.Div([
        html.Br(),
    html.H2('Qualidade', style={'textAlign': 'center'}),
    html.Div(id='period-badge-qualidade', style={'textAlign': 'center', 'marginTop': '6px', 'color': '#555'}),

        # KPI cards de qualidade
        html.Div(id='quality-kpi-cards'),

        html.Br(),

        # Funil de conversão
        html.Div([
            html.H3('Funil de Conversão', style={'textAlign': 'center'}),
            dcc.Graph(id='conversion-funnel')
        ]),

        html.Br(),

        # Motivos de não conversão
        html.Div([
            html.H3('Principais motivos de não conversão', style={'textAlign': 'center'}),
            dcc.Graph(id='reasons-bar')
        ])
    ])


# Layout da Aba Performance
def create_tab_performance_layout():
    return html.Div([
        html.Br(),
    html.H2('Performance por Consultora', style={'textAlign': 'center'}),
    html.Div(id='period-badge-performance', style={'textAlign': 'center', 'marginTop': '6px', 'color': '#555'}),

        # Filtros de período (reusar datepickers existentes via callbacks)
        html.Div(id='performance-kpis'),

        html.Br(),

        # Gráfico comparativo por consultora
        html.Div([
            dcc.Graph(id='consultora-comparative-bar')
        ]),

        html.Br(),

        # Série temporal
        html.Div([
            dcc.Graph(id='time-series-metrics')
        ]),

        html.Br(),

        # Tempo médio de avanço entre etapas
        html.Div([
            html.H3('Tempo médio entre etapas (dias)', style={'textAlign': 'center'}),
            html.Div(id='avg-time-steps')
        ]),

        html.Br(),

        # Ranking de consultoras
        html.Div([
            html.H3('Ranking de Consultoras por Taxa de Conversão', style={'textAlign': 'center'}),
            dcc.Graph(id='ranking-consultoras')
        ])
    ])

# Callback para renderizar conteúdo das abas
@app.callback(Output('tab-content', 'children'),
              Input('tabs', 'value'))
def render_content(active_tab):
    if active_tab == 'tab-1':
        return create_tab1_layout()
    elif active_tab == 'tab-qualidade':
        return create_tab_qualidade_layout()
    elif active_tab == 'tab-performance':
        return create_tab_performance_layout()
    elif active_tab == 'tab-2':
        return create_tab2_layout()


@app.callback(Output('period-badge-qualidade', 'children'), Input('period-store', 'data'))
def update_badge_qualidade(period_store):
    if period_store and period_store.get('start') and period_store.get('end'):
        try:
            s = pd.to_datetime(period_store.get('start')).strftime('%d/%m/%Y')
            e = pd.to_datetime(period_store.get('end')).strftime('%d/%m/%Y')
            return html.Span(f'Período: {s} — {e}', style={'fontWeight': '600'})
        except Exception:
            return ''
    return ''


@app.callback(Output('period-badge-performance', 'children'), Input('period-store', 'data'))
def update_badge_performance(period_store):
    if period_store and period_store.get('start') and period_store.get('end'):
        try:
            s = pd.to_datetime(period_store.get('start')).strftime('%d/%m/%Y')
            e = pd.to_datetime(period_store.get('end')).strftime('%d/%m/%Y')
            return html.Span(f'Período: {s} — {e}', style={'fontWeight': '600'})
        except Exception:
            return ''
    return ''


@app.callback(
    Output('period-store', 'data'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date')],
    prevent_initial_call=False
)
def sync_period_store(start_date, end_date):
    # se date pickers existirem, armazenar; caso contrário manter o existente
    data = {'start': None, 'end': None}
    if start_date:
        try:
            data['start'] = pd.to_datetime(start_date).strftime('%Y-%m-%d')
        except Exception:
            data['start'] = None
    if end_date:
        try:
            data['end'] = pd.to_datetime(end_date).strftime('%Y-%m-%d')
        except Exception:
            data['end'] = None
    return data

# Callback para cards de resumo na aba 1
@app.callback(
    Output('summary-cards', 'children'),
    [Input('date-picker-start', 'date'),
     Input('date-picker-end', 'date'),
     Input('refresh-data', 'n_clicks')]
)
def update_summary_cards(start_date, end_date, n_clicks):
    if not start_date or not end_date:
        return html.Div()

    # Converter inputs de data para datetime
    try:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
    except Exception:
        return html.Div()

    # Filtrar dados por período
    if 'DataReferencia' in df.columns:
        filtered_df = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)]
    else:
        filtered_df = pd.DataFrame()

    # Registros sem consultora definida
    if 'Consultora' in filtered_df.columns:
        sem_consultora_df = filtered_df[filtered_df['Consultora'].isna() | (filtered_df['Consultora'].astype(str).str.strip() == '')]
    else:
        sem_consultora_df = pd.DataFrame()

    total_sem_consultora = len(sem_consultora_df)

    # Agrupar por DataReferencia e contar, ordenando por data crescente
    if not sem_consultora_df.empty and 'DataReferencia' in sem_consultora_df.columns:
        # Agrupa por DataReferencia (datetime), ordena crescente
        sem_consultora_group = sem_consultora_df.groupby(sem_consultora_df['DataReferencia']).size().sort_index()
        tooltip_lines = [f"{dt.strftime('%d/%m/%Y')}: {qtde}" for dt, qtde in sem_consultora_group.items()]
        tooltip_text = '\n'.join(tooltip_lines)
    else:
        tooltip_text = "Nenhum registro sem consultora no período."

    sem_consultora_card = html.Div([
        html.H4(f"{total_sem_consultora}", style={'margin': '0', 'fontSize': '2em', 'color': '#333'}),
        html.P("Sem consultora", style={'margin': '0'})
    ], style={'textAlign': 'center', 'backgroundColor': "#f56e5f", 'padding': '20px',
              'borderRadius': '5px', 'margin': '0', 'width': '180px'}, title=tooltip_text)
    if not start_date or not end_date:
        return html.Div()

    # Converter inputs de data para datetime
    try:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
    except Exception:
        return html.Div()

    # Filtrar dados por período
    if 'DataReferencia' in df.columns:
        filtered_df = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)]
    else:
        filtered_df = pd.DataFrame()

    # Calcular métricas
    total_registros = len(filtered_df)
    positivos = len(filtered_df[filtered_df['Positivo'] == 'Sim']) if 'Positivo' in filtered_df.columns else 0
    negativos = len(filtered_df[filtered_df['Positivo'] == 'Não']) if 'Positivo' in filtered_df.columns else 0

    # Registros por consultora
    if 'Consultora' in filtered_df.columns:
        por_consultora = filtered_df['Consultora'].value_counts()
    else:
        print("Aviso: Coluna 'Consultora' não encontrada. Colunas disponíveis:", filtered_df.columns.tolist())
        por_consultora = pd.Series()  # Serie vazia

    # Paleta pastéis (fallback)
    pastel_palette = ['#FDEBD0', '#E8F8F5', '#F6EBF6', '#FEF9E7', '#E8F6FF', '#FFF0F5', '#FBEFF2', '#EAF8F1']

    # Mapeamento explícito nome -> cor (edite conforme sua preferência)
    consultora_color_map = {
        'Lidiane': '#DCEEFF',   # azul claro
        'Lidiane ': '#DCEEFF',  # possível variação com espaço
        'Jéssica': '#F6E8F6',   # rosa / lilás claro
        'Jessica': '#F6E8F6',   # sem acento
        # adicione outros nomes conhecidos aqui
    }

    # função utilitária para obter cor por nome com fallback determinístico
    def consultora_color(name: str) -> str:
        try:
            if not name:
                return pastel_palette[0]
            key = str(name).strip()
            if key in consultora_color_map:
                return consultora_color_map[key]
            # fallback determinístico baseado em hash
            idx = int(hashlib.md5(key.encode('utf-8')).hexdigest(), 16) % len(pastel_palette)
            return pastel_palette[idx]
        except Exception:
            return pastel_palette[0]

    # construir os cards pais (total, positivos, negativos, whatsapp, proposta)
    total_card = html.Div([
        html.H4(f"{total_registros}", style={'margin': '0', 'fontSize': '2em'}),
        html.P("Total de registros", style={'margin': '0'})
    ], style={'textAlign': 'center', 'backgroundColor': '#f0f0f0', 'padding': '20px',
             'borderRadius': '5px', 'margin': '0', 'width': '180px'})

    pos_parent = html.Div([
        html.H4(f"{positivos}", style={'margin': '0', 'fontSize': '2em', 'color': 'green'}),
        html.P("Positivos", style={'margin': '0'})
    ], style={'textAlign': 'center', 'backgroundColor': '#e8f7ea', 'padding': '20px',
              'borderRadius': '5px', 'margin': '0', 'width': '260px'})

    neg_parent = html.Div([
        html.H4(f"{negativos}", style={'margin': '0', 'fontSize': '2em', 'color': 'red'}),
        html.P("Negativos", style={'margin': '0'})
    ], style={'textAlign': 'center', 'backgroundColor': '#fdecea', 'padding': '20px',
              'borderRadius': '5px', 'margin': '0', 'width': '260px'})

    whatsapp_total = 0
    if 'Contato via wp' in filtered_df.columns:
        whatsapp_total = filtered_df[filtered_df['Contato via wp'].astype(str).str.strip().str.lower() == 'sim'].shape[0]
    whatsapp_parent = html.Div([
        html.H4(f"{whatsapp_total}", style={'margin': '0', 'fontSize': '2em', 'color': '#2b8c6b'}),
        html.P("Contatos WhatsApp", style={'margin': '0'})
    ], style={'textAlign': 'center', 'backgroundColor': '#eefaf7', 'padding': '20px',
              'borderRadius': '5px', 'margin': '0', 'width': '260px'})

    proposta_total = 0
    if 'Proposta' in filtered_df.columns:
        proposta_total = filtered_df[filtered_df['Proposta'].astype(str).str.strip().str.lower() == 'sim'].shape[0]
    proposta_parent = html.Div([
        html.H4(f"{proposta_total}", style={'margin': '0', 'fontSize': '2em', 'color': '#6a4a9f'}),
        html.P("Proposta", style={'margin': '0'})
    ], style={'textAlign': 'center', 'backgroundColor': '#f3eef9', 'padding': '20px',
              'borderRadius': '5px', 'margin': '0', 'width': '260px'})

    # criar linhas de children para cada grupo
    # POSITIVOS children
    pos_children = []
    if 'Positivo' in filtered_df.columns and 'Consultora' in filtered_df.columns:
        pos_by_cons = filtered_df[filtered_df['Positivo'] == 'Sim']['Consultora'].value_counts()
    else:
        pos_by_cons = pd.Series()

    for consultora, cnt in pos_by_cons.items():
        bg = consultora_color(consultora)
        pos_children.append(html.Div([
            html.H4(f"{cnt}", style={'margin': '0', 'fontSize': '1.6em', 'color': 'black'}),
            html.P(f"{consultora}", style={'margin': '0', 'fontSize': '0.9em'})
        ], style={'textAlign': 'center', 'backgroundColor': bg, 'padding': '14px', 'borderRadius': '6px',
                  'margin': '6px', 'width': '160px'}))

    # NEGATIVOS children
    neg_children = []
    if 'Positivo' in filtered_df.columns and 'Consultora' in filtered_df.columns:
        neg_by_cons = filtered_df[filtered_df['Positivo'] == 'Não']['Consultora'].value_counts()
    else:
        neg_by_cons = pd.Series()

    for consultora, cnt in neg_by_cons.items():
        bg = consultora_color(consultora)
        neg_children.append(html.Div([
            html.H4(f"{cnt}", style={'margin': '0', 'fontSize': '1.6em', 'color': 'black'}),
            html.P(f"{consultora}", style={'margin': '0', 'fontSize': '0.9em'})
        ], style={'textAlign': 'center', 'backgroundColor': bg, 'padding': '14px', 'borderRadius': '6px',
                  'margin': '6px', 'width': '160px'}))

    # WHATSAPP children
    whatsapp_children = []
    if 'Contato via wp' in filtered_df.columns and 'Consultora' in filtered_df.columns:
        wp_by_cons = filtered_df[filtered_df['Contato via wp'].astype(str).str.strip().str.lower() == 'sim']['Consultora'].value_counts()
    else:
        wp_by_cons = pd.Series()

    for consultora, cnt in wp_by_cons.items():
        bg = consultora_color(consultora)
        whatsapp_children.append(html.Div([
            html.H4(f"{cnt}", style={'margin': '0', 'fontSize': '1.6em', 'color': 'black'}),
            html.P(f"{consultora}", style={'margin': '0', 'fontSize': '0.9em'})
        ], style={'textAlign': 'center', 'backgroundColor': bg, 'padding': '14px', 'borderRadius': '6px',
                  'margin': '6px', 'width': '160px'}))

    # PROPOSTA children
    proposta_children = []
    if 'Proposta' in filtered_df.columns and 'Consultora' in filtered_df.columns:
        prop_by_cons = filtered_df[filtered_df['Proposta'].astype(str).str.strip().str.lower() == 'sim']['Consultora'].value_counts()
    else:
        prop_by_cons = pd.Series()

    for consultora, cnt in prop_by_cons.items():
        bg = consultora_color(consultora)
        proposta_children.append(html.Div([
            html.H4(f"{cnt}", style={'margin': '0', 'fontSize': '1.6em', 'color': 'black'}),
            html.P(f"{consultora}", style={'margin': '0', 'fontSize': '0.9em'})
        ], style={'textAlign': 'center', 'backgroundColor': bg, 'padding': '14px', 'borderRadius': '6px',
                  'margin': '6px', 'width': '160px'}))

    # TOTAL por consultora (coluna 1) - quantidade total de registros por consultora
    total_children = []
    if 'Consultora' in filtered_df.columns:
        total_by_cons = filtered_df['Consultora'].value_counts()
    else:
        total_by_cons = pd.Series()

    for consultora, cnt in total_by_cons.items():
        bg = consultora_color(consultora)
        total_children.append(html.Div([
            html.H4(f"{cnt}", style={'margin': '0', 'fontSize': '1.6em', 'color': 'black'}),
            html.P(f"{consultora}", style={'margin': '0', 'fontSize': '0.9em'})
        ], style={'textAlign': 'center', 'backgroundColor': bg, 'padding': '14px', 'borderRadius': '6px',
                  'margin': '6px', 'width': '160px'}))

    # montar layout em colunas: cada coluna contém o card pai e, abaixo, os cards das consultoras
    col_sem_consultora = html.Div([
        sem_consultora_card
    ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'width': '180px', 'gap': '14px'})

    col_total = html.Div([
        total_card,
        html.Div(total_children, style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'gap': '8px'})
    ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'width': '180px', 'gap': '14px'})

    col_pos = html.Div([
        pos_parent,
        html.Div(pos_children, style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'gap': '8px'})
    ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'width': '260px', 'gap': '14px'})

    col_neg = html.Div([
        neg_parent,
        html.Div(neg_children, style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'gap': '8px'})
    ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'width': '260px', 'gap': '14px'})

    col_wp = html.Div([
        whatsapp_parent,
        html.Div(whatsapp_children, style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'gap': '8px'})
    ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'width': '260px', 'gap': '14px'})

    col_prop = html.Div([
        proposta_parent,
        html.Div(proposta_children, style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'gap': '8px'})
    ], style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center', 'width': '260px', 'gap': '14px'})

    # linha de colunas
    columns_row = html.Div([col_sem_consultora, col_total, col_pos, col_neg, col_wp, col_prop],
                           style={'display': 'flex', 'justifyContent': 'center', 'gap': '48px', 'width': '100%', 'marginTop': '8px'})

    return columns_row

# Callback para gráfico de pizza
@app.callback(
    Output('qualificado-pie-chart', 'figure'),
    [Input('date-picker-start', 'date'),
     Input('date-picker-end', 'date'),
     Input('refresh-data', 'n_clicks')]
)
def update_pie_chart(start_date, end_date, n_clicks):
    # Se datas não fornecidas, não renderizar
    if not start_date or not end_date:
        return {}

    # Converter inputs de data para datetime
    try:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
    except Exception:
        return {}

    # Filtrar pelo período usando DataReferencia
    if 'DataReferencia' in df.columns:
        filtered_df = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)]
    else:
        # sem coluna DataReferencia, não há como filtrar por período
        return px.pie(values=[1], names=['Dados não disponíveis'], title="Distribuição de Qualificados")

    # Normalizar e mapear para três categorias
    if 'Qualificado' not in filtered_df.columns or filtered_df.empty:
        qualificado_counts = pd.Series([0], index=['Nenhum'])
    else:
        def map_qualificado(v):
            try:
                s = str(v).strip().lower()
            except Exception:
                return 'Não'
            if s == '' or s in ['nan', 'none', 'na']:
                return 'Não'
            if 'sim' in s:
                return 'Sim'
            if 'talvez' in s or 'maybe' in s:
                return 'Talvez'
            if 'não' in s or 'nao' in s or s.startswith('n'):
                return 'Não'
            return 'Não'

        mapped = filtered_df['Qualificado'].apply(map_qualificado)
        qualificado_counts = mapped.value_counts()

    # Garantir ordem consistente: Sim, Talvez, Não
    order = ['Sim', 'Talvez', 'Não']
    labels = [lab for lab in order if lab in qualificado_counts.index]
    values = [int(qualificado_counts.get(lab, 0)) for lab in labels]

    if sum(values) == 0:
        fig = px.pie(values=[1], names=['Dados não disponíveis'])
    else:
        fig = px.pie(values=values, names=labels)

    # Paleta e efeito de destaque — tons suaves profissionais
    color_map = {'Sim': '#4a90e2', 'Talvez': '#f5b041', 'Não': '#e05252'}
    colors = [color_map.get(lab, '#888888') for lab in labels]

    # Mostrar rótulo no formato: "Sim — 12 (40%)" e hover com valor e percentual
    try:
        # texttemplate controla exatamente o texto dentro das fatias
        fig.update_traces(
            texttemplate='%{label} — %{value} (%{percent:.0%})',
            textposition='inside',
            hovertemplate='%{label}: %{value} (%{percent:.0%})<extra></extra>',
            # aplicar pull para destacar as fatias (efeito 'explodido' 3D-like)
            pull=[0.08 if v > 0 else 0 for v in values],
            marker=dict(colors=colors, line=dict(color='white', width=1))
        )
    except Exception:
        # fallback simples
        fig.update_traces(textinfo='percent')

    # removido background SVG para evitar faixa cinza atrás do gráfico

    return fig

# Callback para tabela resumo por dia
@app.callback(
    Output('daily-summary-data-table', 'data'),
    Output('daily-summary-data-table', 'columns'),
    [Input('date-picker-start', 'date'),
     Input('date-picker-end', 'date'),
     Input('refresh-data', 'n_clicks'),
     Input('daily-summary-data-table', 'sort_by')]
)
def update_daily_summary(start_date, end_date, n_clicks, sort_by):
    # Se datas não estiverem definidas, retornar vazio (DataTable já existe no layout)
    if not start_date or not end_date:
        return [], []

    # Converter inputs de data para datetime
    try:
        start = pd.to_datetime(start_date)
        end = pd.to_datetime(end_date)
    except Exception:
        return [], []

    # Filtrar dados por período
    if 'DataReferencia' in df.columns:
        filtered_df = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)]
    else:
        filtered_df = pd.DataFrame()

    # Agrupar por data
    daily_summary = []

    if not filtered_df.empty and 'DataReferencia' in filtered_df.columns:
        unique_dates = filtered_df['DataReferencia'].dt.date.unique()
    else:
        unique_dates = []

    for date_val in unique_dates:
        day_df = filtered_df[filtered_df['DataReferencia'].dt.date == date_val]

        # Calcular métricas do dia
        total_contatos = len(day_df)
        positivos = len(day_df[day_df['Positivo'] == 'Sim']) if 'Positivo' in day_df.columns else 0
        contatos_wp = len(day_df[day_df['Contato via wp'] == 'Sim']) if 'Contato via wp' in day_df.columns else 0
        propostas = len(day_df[day_df['Proposta'] == 'Sim']) if 'Proposta' in day_df.columns else 0

        # Calcular atraso (contatos positivos sem retorno há mais de 7 dias)
        days_diff = (datetime.now().date() - date_val).days
        sem_retorno = len(day_df[day_df['Positivo'] == 'Sim']) if ('Positivo' in day_df.columns and days_diff >= 7) else 0

        daily_summary.append({
            'DATA': date_val.strftime('%d/%m/%Y'),
            'Nº contatos': total_contatos,
            'Positivos': positivos,
            'Contatos': contatos_wp,
            'Proposta': propostas,
            'Sem retorno geral': sem_retorno
        })

    # Criar DataFrame e preparar dados/colunas para o DataTable
    if daily_summary:
        summary_df = pd.DataFrame(daily_summary)
        # coluna auxiliar datetime para ordenação correta
        summary_df['DATA_dt'] = pd.to_datetime(summary_df['DATA'], dayfirst=True, format='%d/%m/%Y', errors='coerce')

        # Ordenação customizada enviada pelo componente (quando o usuário clica no cabeçalho)
        if sort_by and isinstance(sort_by, list) and len(sort_by) > 0:
            sort = sort_by[0]
            col_id = sort.get('column_id')
            direction = sort.get('direction', 'asc') == 'asc'
            if col_id == 'DATA':
                summary_df = summary_df.sort_values('DATA_dt', ascending=direction)
            else:
                # ordenar por coluna visível (se existir)
                if col_id in summary_df.columns:
                    summary_df = summary_df.sort_values(col_id, ascending=direction, kind='mergesort')
        else:
            # comportamento padrão: ordenar DATA decrescente
            summary_df = summary_df.sort_values('DATA_dt', ascending=False)

        # Preparar colunas sem a coluna auxiliar
        out_df = summary_df.drop(columns=['DATA_dt'])
        columns = [{"name": col, "id": col} for col in out_df.columns]
        data = out_df.to_dict('records')
        return data, columns
    else:
        return [], []


# -----------------
# CALLBACKS - QUALIDADE
# -----------------


@app.callback(
    Output('quality-kpi-cards', 'children'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_quality_kpis(start_date, end_date, n_clicks, period_store):
    # calcular KPIs: conversão geral, proposta/positivos, contato_wp/total, taxa atraso
    # Priorizar período vindo do period-store (seleção na aba Análise Geral)
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            start = None
            end = None
    else:
        # Se date pickers não estiverem presentes (ou vazios), inferir período a partir do df
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return html.Div()
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return html.Div()

    # localizar colunas candidatas de forma resiliente
    col_qual = find_col(df, ['Qualificado', 'Qualificado?'])
    col_positivo = find_col(df, ['Positivo', 'Positivo?'])
    col_proposta = find_col(df, ['Proposta', 'Proposta Enviada'])
    col_contato = find_col(df, ['Contato via wp', 'Contato via WhatsApp', 'Contato via Whatsapp', 'ContatoWP'])

    if not any([col_qual, col_positivo, col_proposta, col_contato]):
        return html.Div([
            html.H4('Colunas essenciais ausentes para KPIs de Qualidade'),
            render_data_debug_sample()
        ], style={'color': 'darkred', 'textAlign': 'center'})

    # filtrar pelo período (usar period_store ou DataReferencia se disponível). Se DataReferencia não existir, usar todo o df
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        start = pd.to_datetime(period_store.get('start'))
        end = pd.to_datetime(period_store.get('end'))
        if 'DataReferencia' in df.columns:
            filtered = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
        else:
            filtered = df.copy()
    else:
        if 'DataReferencia' in df.columns:
            filtered = df.copy()
            try:
                start = pd.to_datetime(start_date) if start_date else df['DataReferencia'].min()
                end = pd.to_datetime(end_date) if end_date else df['DataReferencia'].max()
                filtered = filtered[(filtered['DataReferencia'] >= start) & (filtered['DataReferencia'] <= end)].copy()
            except Exception:
                filtered = df.copy()
        else:
            filtered = df.copy()

    total = len(filtered)
    qualificados = len(filtered[filtered[col_qual].astype(str).str.contains('sim', case=False, na=False)]) if col_qual and col_qual in filtered.columns else 0
    positivos = len(filtered[filtered[col_positivo] == 'Sim']) if col_positivo and col_positivo in filtered.columns else 0
    propostas = len(filtered[filtered[col_proposta] == 'Sim']) if col_proposta and col_proposta in filtered.columns else 0
    contatos_wp = len(filtered[filtered[col_contato] == 'Sim']) if col_contato and col_contato in filtered.columns else 0

    # leads atrasados: usar mesma lógica da aba atrasos com 7 dias padrão
    today = datetime.now().date()
    if 'Positivo' in filtered.columns and 'DataReferencia' in filtered.columns:
        atrasados = filtered[filtered['Positivo'] == 'Sim'].copy()
        atrasados['DiasAtraso'] = atrasados['DataReferencia'].apply(lambda x: (today - x.date()).days if pd.notnull(x) else None)
        atrasados = len(atrasados[atrasados['DiasAtraso'] >= 7])
    else:
        atrasados = 0

    # Calcular taxas com proteção contra divisão por zero
    def pct(n, d):
        return f"{(n/d*100):.1f}%" if d and d > 0 else "0.0%"

    kpis = [
        {'label': 'Taxa de conversão geral', 'value': pct(qualificados, total), 'raw': f"{qualificados}/{total}"},
        {'label': 'Taxa de proposta enviada', 'value': pct(propostas, positivos), 'raw': f"{propostas}/{positivos}"},
        {'label': 'Taxa de contato via WhatsApp', 'value': pct(contatos_wp, total), 'raw': f"{contatos_wp}/{total}"},
        {'label': 'Taxa de atraso', 'value': pct(atrasados, total), 'raw': f"{atrasados}/{total}"}
    ]

    cards = []
    for k in kpis:
        cards.append(html.Div([
            html.H4(k['value'], style={'margin': '0', 'fontSize': '1.6em'}),
            html.P(k['label'], style={'margin': '0', 'fontSize': '0.9em'}),
            html.Small(k['raw'], style={'color': '#666'})
        ], style={'textAlign': 'center', 'backgroundColor': '#f7fbff', 'padding': '16px', 'borderRadius': '6px', 'margin': '8px', 'width': '220px', 'display': 'inline-block'}))

    return html.Div(cards, style={'display': 'flex', 'justifyContent': 'center', 'flexWrap': 'wrap'})


@app.callback(
    Output('conversion-funnel', 'figure'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_conversion_funnel(start_date, end_date, n_clicks, period_store):
    # Build funnel: Total Leads → Positivos → Contato via WhatsApp → Proposta → Qualificado
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return {}
    else:
        # Se date pickers não estiverem presentes, inferir do df
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return {}
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return {}

    # localizar colunas
    col_positivo = find_col(df, ['Positivo', 'Positivo?'])
    col_contato = find_col(df, ['Contato via wp', 'Contato via WhatsApp', 'Contato via Whatsapp', 'ContatoWP'])
    col_proposta = find_col(df, ['Proposta', 'Proposta Enviada'])
    col_qual = find_col(df, ['Qualificado', 'Qualificado?'])

    if not any([col_positivo, col_contato, col_proposta, col_qual]):
        fig = go.Figure()
        fig.add_annotation(text='Colunas essenciais do funil não encontradas', xref='paper', yref='paper', showarrow=False)
        fig.update_layout(title='Colunas faltando para funil')
        return fig

    # filtrar pelo período
    if 'DataReferencia' in df.columns and period_store and period_store.get('start') and period_store.get('end'):
        start = pd.to_datetime(period_store.get('start'))
        end = pd.to_datetime(period_store.get('end'))
        d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
    elif 'DataReferencia' in df.columns:
        try:
            sd = pd.to_datetime(start_date) if start_date else df['DataReferencia'].min()
            ed = pd.to_datetime(end_date) if end_date else df['DataReferencia'].max()
            d = df[(df['DataReferencia'] >= sd) & (df['DataReferencia'] <= ed)].copy()
        except Exception:
            d = df.copy()
    else:
        d = df.copy()

    total = len(d)
    positivos = len(d[d[col_positivo] == 'Sim']) if col_positivo and col_positivo in d.columns else 0
    contatos_wp = len(d[d[col_contato] == 'Sim']) if col_contato and col_contato in d.columns else 0
    propostas = len(d[d[col_proposta] == 'Sim']) if col_proposta and col_proposta in d.columns else 0
    qualificados = len(d[d[col_qual].astype(str).str.contains('sim', case=False, na=False)]) if col_qual and col_qual in d.columns else 0

    steps = ['Total Leads', 'Positivos', 'Contato via WhatsApp', 'Proposta', 'Qualificado']
    values = [total, positivos, contatos_wp, propostas, qualificados]

    fig = go.Figure(go.Funnel(
        y=steps,
        x=values,
        textinfo='value+percent previous',
        marker={'color': ['#4a90e2', '#5dade2', '#f5b041', '#f4a261', '#58d68d']}
    ))
    fig.update_layout(margin=dict(l=20, r=20, t=30, b=20))
    return fig


@app.callback(
    Output('reasons-bar', 'figure'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_reasons_bar(start_date, end_date, n_clicks, period_store):
    # Extrair motivos da coluna 'Histórico' e categorizar automaticamente
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return {}
    else:
        # Inferir período a partir do df se necessário
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return {}
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return {}

    if 'DataReferencia' in df.columns:
        d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
    else:
        d = df.copy()

    # verificar coluna essencial 'Histórico'
    # normalizar variantes de Histórico
    hist_col = find_col(d, ['Histórico', 'Historico', 'Historico '])
    if not hist_col:
        fig = px.bar(x=[1], y=['Sem dados'], orientation='h').update_layout(title='Coluna "Histórico" não encontrada')
        fig.update_layout(annotations=[dict(text=f"Colunas: {df.columns.tolist()}", xref='paper', yref='paper', x=0, y=-0.2, showarrow=False)])
        return fig

    raw = d[hist_col].astype(str).fillna('').str.lower()

    raw = d['Histórico'].astype(str).fillna('').str.lower()

    # regras simples de categorização (padrões comuns)
    categories = {
        'não tem sistema': ['não tem sistema', 'nao tem sistema', 'sem sistema'],
        'sem interesse': ['sem interesse', 'não interessado', 'nao interessado', 'sem interesse no momento'],
        'responsável ausente': ['responsável ausente', 'responsavel ausente', 'sem responsável', 'sem responsavel'],
        'preço': ['preço', 'preco', 'caro', 'muito caro']
    }

    def classify(text):
        for cat, kws in categories.items():
            for kw in kws:
                if kw in text:
                    return cat
        # fallback: tentar tokens comuns
        if 'nao' in text or 'não' in text:
            return 'sem interesse'
        if text.strip() == '':
            return 'sem resposta'
        return 'outros'

    cats = raw.apply(classify)
    counts = cats.value_counts().head(10)

    fig = px.bar(x=counts.values, y=counts.index, orientation='h', labels={'x': 'Contagem', 'y': 'Motivo'})
    fig.update_layout(margin=dict(l=80, r=20, t=20, b=20))
    return fig


# -----------------
# CALLBACKS - PERFORMANCE
# -----------------


@app.callback(
    Output('performance-kpis', 'children'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_performance_kpis(start_date, end_date, n_clicks, period_store):
    # KPIs resumidos por consultora (nº leads, % positivos, % propostas, % qualificados, % atrasados)
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return html.Div()
    else:
        # Inferir período a partir do df se os date pickers não estiverem disponíveis
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return html.Div()
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return html.Div()

    # Verificar coluna essencial 'Consultora'
    if 'Consultora' not in df.columns:
        return html.Div([
            html.H4('Coluna "Consultora" ausente'),
            html.P('A visão por consultora exige a coluna "Consultora" na planilha.'),
            render_data_debug_sample()
        ], style={'color': 'darkred', 'textAlign': 'center'})

    if 'DataReferencia' in df.columns:
        d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
    else:
        d = df.copy()

    if 'Consultora' not in d.columns or d.empty:
        return html.Div('Nenhum dado por consultora disponível')

    group = d.groupby('Consultora')
    resumo = []
    today = datetime.now().date()
    for name, g in group:
        total = len(g)
        positivos = len(g[g['Positivo'] == 'Sim']) if 'Positivo' in g.columns else 0
        propostas = len(g[g['Proposta'] == 'Sim']) if 'Proposta' in g.columns else 0
        qualificados = len(g[g['Qualificado'].astype(str).str.contains('sim', case=False, na=False)]) if 'Qualificado' in g.columns else 0
        atrasos = 0
        if 'Positivo' in g.columns and 'DataReferencia' in g.columns:
            atraso_df = g[g['Positivo'] == 'Sim'].copy()
            atraso_df['DiasAtraso'] = atraso_df['DataReferencia'].apply(lambda x: (today - x.date()).days if pd.notnull(x) else None)
            atrasos = len(atraso_df[atraso_df['DiasAtraso'] >= 7])

        resumo.append({
            'Consultora': name,
            'Leads': total,
            '% Positivos': (positivos/total*100) if total else 0,
            '% Propostas': (propostas/total*100) if total else 0,
            '% Qualificados': (qualificados/total*100) if total else 0,
            '% Atrasados': (atrasos/total*100) if total else 0
        })

    resumo_df = pd.DataFrame(resumo).sort_values('Leads', ascending=False)

    # Renderizar uma tabela simples com os KPIs
    return dash_table.DataTable(
        data=resumo_df.to_dict('records'),
        columns=[{'name': c, 'id': c} for c in resumo_df.columns],
        style_table={'overflowX': 'auto'},
        style_cell={'textAlign': 'center'}
    )


@app.callback(
    Output('consultora-comparative-bar', 'figure'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_consultora_comparative(start_date, end_date, n_clicks, period_store):
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return {}
    else:
        # Inferir período se necessário
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return {}
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return {}

    if 'Consultora' not in df.columns:
        fig = go.Figure()
        fig.add_annotation(text='Coluna "Consultora" não encontrada para comparativo', xref='paper', yref='paper', showarrow=False)
        fig.update_layout(title='Coluna "Consultora" não encontrada')
        return fig

    if 'DataReferencia' in df.columns:
        d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
    else:
        d = df.copy()

    if 'Consultora' not in d.columns or d.empty:
        return {}

    group = d.groupby('Consultora')
    data = []
    for name, g in group:
        total = len(g)
        positivos = len(g[g['Positivo'] == 'Sim']) if 'Positivo' in g.columns else 0
        propostas = len(g[g['Proposta'] == 'Sim']) if 'Proposta' in g.columns else 0
        qualificados = len(g[g['Qualificado'].astype(str).str.contains('sim', case=False, na=False)]) if 'Qualificado' in g.columns else 0
        atrasos = 0
        today = datetime.now().date()
        if 'Positivo' in g.columns and 'DataReferencia' in g.columns:
            atraso_df = g[g['Positivo'] == 'Sim'].copy()
            atraso_df['DiasAtraso'] = atraso_df['DataReferencia'].apply(lambda x: (today - x.date()).days if pd.notnull(x) else None)
            atrasos = len(atraso_df[atraso_df['DiasAtraso'] >= 7])

        data.append({'Consultora': name, 'Leads': total, 'PctPositivos': positivos/total*100 if total else 0, 'PctPropostas': propostas/total*100 if total else 0, 'PctQualificados': qualificados/total*100 if total else 0, 'PctAtrasados': atrasos/total*100 if total else 0})

    plot_df = pd.DataFrame(data)
    # criar gráfico agrupado com barras por consultora
    fig = px.bar(plot_df, x='Consultora', y=['Leads', 'PctPositivos', 'PctPropostas', 'PctQualificados', 'PctAtrasados'], barmode='group')
    fig.update_layout(xaxis={'categoryorder':'total descending'}, margin=dict(l=20, r=20, t=30, b=80))
    return fig


@app.callback(
    Output('time-series-metrics', 'figure'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_time_series(start_date, end_date, n_clicks, period_store):
    # evolução diária de leads, propostas, qualificados
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return {}
    else:
        # Inferir período se os date pickers não estiverem presentes
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return {}
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return {}

    if 'DataReferencia' in df.columns:
        d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
    else:
        d = df.copy()

    if d.empty or 'DataReferencia' not in d.columns:
        fig = go.Figure()
        fig.add_annotation(text='Coluna "DataReferencia" ausente ou sem dados para série temporal', xref='paper', yref='paper', showarrow=False)
        fig.update_layout(title='DataReferencia ausente ou sem dados')
        return fig

    d['date'] = d['DataReferencia'].dt.date
    daily = d.groupby('date').agg(
        leads=('Empresa', 'count') if 'Empresa' in d.columns else ('DataReferencia', 'count'),
        propostas=('Proposta', lambda x: (x=='Sim').sum()) if 'Proposta' in d.columns else ('DataReferencia', 'count'),
        qualificados=('Qualificado', lambda x: x.astype(str).str.contains('sim', case=False, na=False).sum()) if 'Qualificado' in d.columns else ('DataReferencia', 'count')
    ).reset_index()

    fig = px.line(daily, x='date', y=['leads', 'propostas', 'qualificados'])
    fig.update_layout(xaxis_title='Data', yaxis_title='Contagem', legend_title='Métrica', margin=dict(l=20, r=20, t=30, b=20))
    return fig


@app.callback(
    Output('avg-time-steps', 'children'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_avg_time_steps(start_date, end_date, n_clicks, period_store):
    # calcular tempo médio entre etapas estimado por datas nos campos disponíveis
    # pressupõe existência de colunas com timestamps ou DataReferencia para fases; se não existir, aproximar com DataReferencia
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return html.Div()
    else:
        # Inferir período a partir do df se necessário
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return html.Div()
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return html.Div()

    # usar DataReferencia como data de entrada e tentar inferir datas de contato/proposta/qualificacao a partir de colunas booleanas com registros posteriores
    if 'DataReferencia' not in df.columns:
        return html.Div('Dados insuficientes para calcular tempos médios')

    d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()

    # Tentativa: se existirem colunas 'DataContato', 'DataProposta', 'DataQualificado' usar elas; caso contrário, não calcular
    date_cols = [c for c in ['DataContato', 'DataProposta', 'DataQualificado'] if c in d.columns]
    if not date_cols:
        # sem datas, retornar aviso explícito com colunas detectadas
        return html.Div([
            html.H4('Colunas de etapa ausentes'),
            html.P('Não foram encontradas colunas de data para etapas (ex.: DataContato, DataProposta, DataQualificado).'),
            render_data_debug_sample()
        ], style={'color': 'darkred', 'textAlign': 'center'})

    # converter colunas de data presentes
    for c in date_cols:
        d[c] = pd.to_datetime(d[c], errors='coerce')

    # calcular diferenças médias entre pares disponíveis
    rows = []
    steps = [('DataReferencia', 'DataContato'), ('DataContato', 'DataProposta'), ('DataProposta', 'DataQualificado')]
    for a, b in steps:
        if a in d.columns and b in d.columns:
            valid = d.dropna(subset=[a, b]).copy()
            if not valid.empty:
                diffs = (valid[b] - valid[a]).dt.days
                mean_days = diffs.mean()
                rows.append({'Etapa': f'{a} → {b}', 'Dias médios': round(mean_days, 1)})

    if not rows:
        return html.Div('Dados insuficientes para cálculo de tempos médios')

    table_df = pd.DataFrame(rows)
    return dash_table.DataTable(data=table_df.to_dict('records'), columns=[{'name': c, 'id': c} for c in table_df.columns], style_cell={'textAlign': 'center'})


@app.callback(
    Output('ranking-consultoras', 'figure'),
    [Input('date-picker-start', 'date'), Input('date-picker-end', 'date'), Input('refresh-data', 'n_clicks'), Input('period-store', 'data')]
)
def update_ranking_consultoras(start_date, end_date, n_clicks, period_store):
    # Priorizar period-store
    if period_store and isinstance(period_store, dict) and period_store.get('start') and period_store.get('end'):
        try:
            start = pd.to_datetime(period_store.get('start'))
            end = pd.to_datetime(period_store.get('end'))
        except Exception:
            return {}
    else:
        # Inferir período se necessário
        if not start_date or not end_date:
            if 'DataReferencia' in df.columns and not df['DataReferencia'].dropna().empty:
                start = df['DataReferencia'].min()
                end = df['DataReferencia'].max()
            else:
                return {}
        else:
            try:
                start = pd.to_datetime(start_date)
                end = pd.to_datetime(end_date)
            except Exception:
                return {}

    if 'Consultora' not in df.columns:
        fig = go.Figure()
        fig.add_annotation(text='Coluna "Consultora" ausente para ranking', xref='paper', yref='paper', showarrow=False)
        fig.update_layout(title='Consultora ausente')
        return fig

    if 'DataReferencia' in df.columns:
        d = df[(df['DataReferencia'] >= start) & (df['DataReferencia'] <= end)].copy()
    else:
        d = df.copy()

    if 'Consultora' not in d.columns:
        return {}

    group = d.groupby('Consultora')
    ranking = []
    for name, g in group:
        total = len(g)
        qualificados = len(g[g['Qualificado'].astype(str).str.contains('sim', case=False, na=False)]) if 'Qualificado' in g.columns else 0
        conv_rate = (qualificados/total*100) if total else 0
        ranking.append({'Consultora': name, 'Taxa Conversão': conv_rate})

    rdf = pd.DataFrame(ranking).sort_values('Taxa Conversão', ascending=False)
    fig = px.bar(rdf, x='Consultora', y='Taxa Conversão')
    fig.update_layout(margin=dict(l=20, r=20, t=20, b=80))
    return fig

# Callback para cards de atraso por consultora
@app.callback(
    Output('overdue-cards', 'children'),
    [Input('days-overdue-config', 'value'), Input('refresh-data', 'n_clicks'), Input('tabs', 'value'), Input('selected-consultora', 'data')]
)
def update_overdue_cards(days_config, n_clicks, active_tab, selected_consultora):
    if not days_config:
        return html.Div()

    # Só calcular quando a aba de atrasos estiver ativa
    if active_tab != 'tab-2':
        return html.Div()

    today = datetime.now().date()

    # Filtrar contatos em atraso
    if 'Positivo' in df.columns and 'DataReferencia' in df.columns:
        overdue_df = df[df['Positivo'] == 'Sim'].copy()
        overdue_df['dias_atraso'] = overdue_df['DataReferencia'].apply(lambda x: (today - x.date()).days if pd.notnull(x) else None)
        overdue_df = overdue_df[overdue_df['dias_atraso'] >= days_config]
    else:
        overdue_df = pd.DataFrame()

    # Contar por consultora
    if 'Consultora' in overdue_df.columns and not overdue_df.empty:
        overdue_by_consultora = overdue_df['Consultora'].value_counts()
    else:
        overdue_by_consultora = pd.Series()

    # Paleta de cores pastéis (tons claros) para aplicar aos cards ciclicamente
    pastel_palette = [
        '#FDEBD0',  # pêssego claro
        '#E8F8F5',  # menta clara
        '#F6EBF6',  # lavanda clara
        '#FEF9E7',  # amarelo suave
        '#E8F6FF',  # azul claro
        '#FFF0F5',  # rosa claro
    ]

    cards = []
    for idx, (consultora, count) in enumerate(overdue_by_consultora.items()):
        card_id = {'type': 'overdue-card', 'index': str(consultora)}
        # cor de fundo baseada no índice da consultora
        bg_color = pastel_palette[idx % len(pastel_palette)]

        # estilo base
        base_style = {'backgroundColor': bg_color, 'padding': '16px', 'borderRadius': '5px',
                      'margin': '10px', 'width': '200px', 'display': 'inline-block', 'border': '1px solid transparent', 'cursor': 'pointer'}

        # se a consultora estiver selecionada, aplicar destaque (override)
        if selected_consultora and str(consultora) == str(selected_consultora):
            highlight_style = base_style.copy()
            highlight_style.update({'backgroundColor': '#e8f4ff', 'border': '2px solid #4a90e2', 'boxShadow': '0 4px 8px rgba(74,144,226,0.15)'})
            card_style = highlight_style
        else:
            card_style = base_style

        cards.append(
            html.Button([
                html.Div([
                    html.H4(f"{count}", style={'margin': '0', 'fontSize': '2em', 'color': 'black'}),
                    html.P(f"{consultora}", style={'margin': '0', 'fontSize': '2em', 'color': 'black'}),
                    #html.P("em atraso", style={'margin': '0', 'fontSize': '0.8em'})
                ], style={'textAlign': 'center'})
            ],
            id=card_id,
            n_clicks=0,
            style=card_style)
        )

    return html.Div(cards, style={'display': 'flex', 'justifyContent': 'center', 'flexWrap': 'wrap', 'gap': '10px'})

# Callback para tabela de contatos em atraso
@app.callback(
    Output('overdue-table', 'children'),
    [Input('days-overdue-config', 'value'), Input('refresh-data', 'n_clicks'), Input('tabs', 'value'), Input('selected-consultora', 'data')]
)
def update_overdue_table(days_config, n_clicks, active_tab, selected_consultora):
    if not days_config:
        return html.Div()

    # Só calcular quando a aba de atrasos estiver ativa
    if active_tab != 'tab-2':
        return html.Div()

    today = datetime.now().date()

    # Filtrar contatos em atraso
    if 'Positivo' in df.columns and 'DataReferencia' in df.columns:
        overdue_df = df[df['Positivo'] == 'Sim'].copy()
        overdue_df['DiasAtraso'] = overdue_df['DataReferencia'].apply(lambda x: (today - x.date()).days if pd.notnull(x) else None)
        overdue_df = overdue_df[overdue_df['DiasAtraso'] >= days_config]
    else:
        overdue_df = pd.DataFrame()

    # Se uma consultora foi selecionada, filtrar somente por ela
    if selected_consultora:
        if 'Consultora' in overdue_df.columns:
            overdue_df = overdue_df[overdue_df['Consultora'].astype(str) == str(selected_consultora)]
        else:
            overdue_df = pd.DataFrame()

    # Selecionar e ordenar colunas
    colunas_disponiveis = ['DiasAtraso', 'Empresa', 'Consultora', 'Histórico']
    colunas_existentes = [col for col in colunas_disponiveis if col in overdue_df.columns]

    if colunas_existentes and not overdue_df.empty:
        table_df = overdue_df[colunas_existentes].copy()
        table_df = table_df.sort_values('DiasAtraso', ascending=False) if 'DiasAtraso' in table_df.columns else table_df
    else:
        table_df = pd.DataFrame({'Aviso': ['Nenhuma coluna disponível']})

    return dash_table.DataTable(
        data=table_df.to_dict('records'),
        columns=[{"name": col, "id": col} for col in table_df.columns],
        style_cell={'textAlign': 'left', 'whiteSpace': 'normal', 'height': 'auto'},
        style_header={'backgroundColor': 'lightcoral', 'fontWeight': 'bold'},
        style_data_conditional=[
            {
                'if': {'filter_query': '{DiasAtraso} >= 15'},
                'backgroundColor': '#ffcccc',
            },
            {
                'if': {'filter_query': '{DiasAtraso} >= 30'},
                'backgroundColor': '#ff9999',
            }
        ]
    )


@app.callback(
    Output('selected-consultora', 'data'),
    [Input({'type': 'overdue-card', 'index': ALL}, 'n_clicks')],
    [State('selected-consultora', 'data')],
    prevent_initial_call=True
)
def select_consultora(n_clicks_list, current_selected):
    # identificar qual botão disparou
    ctx = dash.callback_context
    if not ctx.triggered:
        raise dash.exceptions.PreventUpdate
    triggered = ctx.triggered[0]['prop_id'].split('.')[0]
    try:
        triggered_id = json.loads(triggered)
        name = triggered_id.get('index', '')
    except Exception:
        raise dash.exceptions.PreventUpdate

    # toggle: se clicar na mesma consultora desmarca
    if current_selected == name:
        return ''
    return name

# Callback para atualizar dados manualmente
@app.callback(
    Output('last-refresh', 'children'),
    Input('refresh-data', 'n_clicks'),
    prevent_initial_call=True
)
def refresh_data(n_clicks):
    if n_clicks > 0:
        # Recarregar dados do Google Sheets
        global df
        df = get_google_sheets_data()
        ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        return f"Última atualização: {ts}"
    raise dash.exceptions.PreventUpdate

if __name__ == '__main__':
    # Rodar o servidor em host e porta especificados
    app.run(host='0.0.0.0', port=8555, debug=False, dev_tools_hot_reload=False, use_reloader=False)