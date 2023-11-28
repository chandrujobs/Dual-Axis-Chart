import pandas as pd
import dash
import dash_bootstrap_components as dbc
from dash import dcc, html, Input, Output
import plotly.graph_objects as go

# Load the dataset
data = pd.read_excel("C:/Users/Chandru/OneDrive/Desktop/Python Visuals/Sample - Superstore.xls", sheet_name="Orders")

# Preprocess the data
data['Year'] = data['Order Date'].dt.year
annual_sales = data.groupby('Year')['Sales'].sum()
annual_profit = data.groupby('Year')['Profit'].sum()

# Create a Dash application
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Define the app layout
app.layout = dbc.Container([
    dbc.Row([
        dbc.Col([
            dcc.Dropdown(
                id='year-dropdown',
                options=[{'label': year, 'value': year} for year in data['Year'].unique()],
                value=data['Year'].unique(),
                multi=True
            )
        ], width=12)
    ]),
    dbc.Row([
        dbc.Col([
            dcc.Graph(id='dual-axis-chart')
        ], width=12)
    ])
], fluid=True)

# Define callback to update graph
@app.callback(
    Output('dual-axis-chart', 'figure'),
    [Input('year-dropdown', 'value')]
)
def update_graph(selected_years):
    filtered_sales = annual_sales[annual_sales.index.isin(selected_years)]
    filtered_profit = annual_profit[annual_profit.index.isin(selected_years)]

    # Create figure with secondary y-axis
    fig = go.Figure()

    # Add traces
    fig.add_trace(go.Bar(x=filtered_sales.index, y=filtered_sales, name='Sales'))
    fig.add_trace(go.Scatter(x=filtered_profit.index, y=filtered_profit, name='Profit', yaxis='y2'))

    # Create axis objects
    fig.update_layout(
        title='Sales and Profit Trends Over Time',
        xaxis_title='Year',
        yaxis_title='Sales',
        yaxis2=dict(
            title='Profit',
            overlaying='y',
            side='right'
        )
    )

    return fig

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True, port=8061)
