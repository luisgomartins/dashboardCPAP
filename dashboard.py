import pandas as pd
import plotly.express as px
from dash import Dash, dash_table, dcc, html, Input, Output
import dash_bootstrap_components as dbc
from datetime import datetime

# Carregar dados
file_path = 'CPAP_actions_open.xlsx'
df_original = pd.read_excel(file_path, sheet_name='datasheet')

# Converter colunas de data
df_original['Date'] = pd.to_datetime(df_original['Date'], errors='coerce')
df_original['Target'] = pd.to_datetime(df_original['Target'], errors='coerce')

# Verificar ações atrasadas
today = pd.to_datetime(datetime.today().date())
df_original['Status'] = df_original['Target'].apply(lambda x: 'Late' if x < today else 'On Time')

# Inicializar o app Dash
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

# Lista de responsáveis e status
responsibles = sorted(df_original['Responsible'].dropna().unique())
status_options = ['On Time', 'Late']

# Layout da dashboard
app.layout = html.Div(
    [
        dbc.Container([
            html.H1("Dashboard CPAP - Actions Open", className='text-center mb-4 mt-4'),

            # Filtros
            dbc.Row([
                dbc.Col([
                    html.Label("Responsible:"),
                    dcc.Dropdown(
                        options=[{'label': r, 'value': r} for r in responsibles],
                        multi=True,
                        id='responsible-filter',
                        placeholder='Select Responsible...'
                    )
                ], md=4),

                dbc.Col([
                    html.Label("Status:"),
                    dcc.Dropdown(
                        options=[{'label': s, 'value': s} for s in status_options],
                        multi=True,
                        id='status-filter',
                        placeholder='Select Status...'
                    )
                ], md=4),
            ]),

            html.Br(),

            # Cards de métricas
            dbc.Row([
                dbc.Col(dbc.Card([
                    dbc.CardBody([
                        html.H5("Total Actions", className="card-title"),
                        html.H2(id="total-actions", className="card-text")
                    ])
                ], color="primary", inverse=True)),

                dbc.Col(dbc.Card([
                    dbc.CardBody([
                        html.H5("On Time", className="card-title"),
                        html.H2(id="on-time-actions", className="card-text")
                    ])
                ], color="success", inverse=True)),

                dbc.Col(dbc.Card([
                    dbc.CardBody([
                        html.H5("Late", className="card-title"),
                        html.H2(id="late-actions", className="card-text")
                    ])
                ], color="danger", inverse=True)),
            ], className='mb-4'),

            # Tabela
            html.H3("Actions Table", className='text-start'),
            dash_table.DataTable(
                id='table',
                style_table={'overflowY': 'auto', 'maxHeight': '250px'},
                style_cell={
                    'textAlign': 'left',
                    'padding': '10px',
                    'whiteSpace': 'normal',
                    'height': 'auto',
                    'fontFamily': 'Arial',
                    'fontSize': '14px'
                },
                style_header={
                    'backgroundColor': '#007bff',
                    'color': 'white',
                    'fontWeight': 'bold',
                    'textAlign': 'center'
                },
                style_data={
                    'border': '1px solid grey'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': 'rgb(248, 248, 248)'
                    },
                    {
                        'if': {'filter_query': '{Status} = "Late"'},
                        'backgroundColor': '#FFBABA',
                        'color': 'black'
                    },
                    {
                        'if': {'state': 'active'},
                        'backgroundColor': '#CCE5FF',
                        'border': '1px solid #004085'
                    }
                ]
            ),

            html.Hr(),

            # Gráficos Totais
            html.H3("Total Actions by Responsible", className='text-start'),
            dbc.Row([
                dbc.Col(dcc.Graph(id='bar-total'), md=6),
                dbc.Col(dcc.Graph(id='pie-total'), md=6)
            ]),

            html.Hr(),

            # Gráficos de Atrasados
            html.H3("Late Actions by Responsible", className='text-start'),
            dbc.Row([
                dbc.Col(dcc.Graph(id='bar-late'), md=6),
                dbc.Col(dcc.Graph(id='pie-late'), md=6)
            ])
        ], fluid=True)
    ],
    style={
        "backgroundColor": "#FFFFFF",
        "minHeight": "100vh",
        "padding": "20px"
    }
)

# Callback para atualizar a dashboard
@app.callback(
    Output('table', 'data'),
    Output('table', 'columns'),
    Output('bar-total', 'figure'),
    Output('pie-total', 'figure'),
    Output('bar-late', 'figure'),
    Output('pie-late', 'figure'),
    Output('total-actions', 'children'),
    Output('on-time-actions', 'children'),
    Output('late-actions', 'children'),
    Input('responsible-filter', 'value'),
    Input('status-filter', 'value')
)
def update_dashboard(responsible_filter, status_filter):
    df = df_original.copy()

    # Aplicar filtros
    if responsible_filter:
        df = df[df['Responsible'].isin(responsible_filter)]
    if status_filter:
        df = df[df['Status'].isin(status_filter)]

    # Contagens
    total = len(df)
    on_time = len(df[df['Status'] == 'On Time'])
    late = len(df[df['Status'] == 'Late'])

    # Gráficos totais
    total_by_responsible = df.groupby('Responsible').size().reset_index(name='Total Actions')
    bar_total = px.bar(total_by_responsible, x='Responsible', y='Total Actions', text='Total Actions', title='Total Actions by Responsible')
    pie_total = px.pie(df, names='Responsible', title='Percentage of Actions by Responsible')

    # Gráficos atrasados
    df_late = df[df['Status'] == 'Late']
    late_by_responsible = df_late.groupby('Responsible').size().reset_index(name='Late Actions')
    bar_late = px.bar(late_by_responsible, x='Responsible', y='Late Actions', text='Late Actions', title='Late Actions by Responsible')
    pie_late = px.pie(df_late, names='Responsible', title='Percentage of Late Actions by Responsible')

    # Tabela
    columns = [{'name': i, 'id': i} for i in df.columns]
    data = df.to_dict('records')

    return data, columns, bar_total, pie_total, bar_late, pie_late, total, on_time, late


# Rodar servidor
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8050)
