"""
Phil - Signed DAF Agreements.py

Script to build dashboard listing status of confirmed
"Fund Agreement" documentation attached to CSuite fund records
"""

# Imports
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, dash_table
import re
from datetime import datetime, timedelta
from SQL import queries     # Stores Mitch's SQL Scripts

QUERIES = queries.QUERIES

cursor = queries.cursor
engine = queries.engine



# 1. Load data
df_funds = pd.read_sql_query(QUERIES['open_daf_funds'], engine)
df_files = pd.read_sql_query(QUERIES['fund_files'], engine)

# 2. Update df_funds with counts of files / file types
file_counts = df_files.groupby('funit_id').size().reset_index(name='total_files')

# 3. Count files with "fund agreement" in the name
def has_fund_agreement(name):
    return bool(re.search(r'fund\s+agreement', name, re.IGNORECASE))

agreement_counts = df_files.groupby('funit_id').apply(
    lambda x: x['file_name'].apply(has_fund_agreement).sum()
).reset_index(name='fund_agreement_in_name')

# 4. Count signed fund agreements
signed_agreements = df_files[df_files['File Category'] == 'FUND AGREEMENT - SIGNED'].groupby('funit_id').size().reset_index(name='fund_agreement_signed')

# Merge all the new information with df_funds
df_funds = df_funds.merge(file_counts, on='funit_id', how='left')
df_funds = df_funds.merge(agreement_counts, on='funit_id', how='left')
df_funds = df_funds.merge(signed_agreements, on='funit_id', how='left')

# Fill NaN values with 0 for funds without any files
df_funds[['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']] = df_funds[['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']].fillna(0)

# Convert columns to integer type
df_funds[['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']] = df_funds[['total_files', 'fund_agreement_in_name', 'fund_agreement_signed']].astype(int)

print(df_funds[['funit_id', 'total_files', 'fund_agreement_in_name', 'fund_agreement_signed']])

# Get today's date -2 (CSuite data is on a 2 day lag)
date_minus_two = datetime.today() - timedelta(days=2)

import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output, State, dash_table, callback_context

# Assuming df_funds is already created and contains the necessary columns

# Create the summary DataFrame
summary_df = df_funds.groupby('Steward Name').agg({
    'funit_id': 'count',
    'fund_agreement_in_name': lambda x: (x > 0).sum(),
    'fund_agreement_signed': lambda x: (x > 0).sum()
}).reset_index()

summary_df.columns = ['Steward Name', 'Total Funds', '# Funds with "Fund Agreement" in filename',
                      '# Funds with Fund Agreement signed']

# Create the Dash app
app = Dash(__name__)

# Define the layout
app.layout = html.Div([
    html.H1("Fund Steward Summary"),
    dcc.Graph(id='summary-graph'),
    html.Button('Clear Selection', id='clear-button', style={'display': 'none'}),
    html.Div(id='drill-down-table')
])


# Callback for the main graph and clear button visibility
@app.callback(
    [Output('summary-graph', 'figure'),
     Output('clear-button', 'style')],
    [Input('summary-graph', 'clickData'),
     Input('clear-button', 'n_clicks')],
    [State('summary-graph', 'figure')]
)
def update_graph(clickData, n_clicks, current_figure):
    ctx = callback_context
    if not ctx.triggered:
        button_style = {'display': 'none'}
    else:
        input_id = ctx.triggered[0]['prop_id'].split('.')[0]

        if input_id == 'clear-button':
            clickData = None
            button_style = {'display': 'none'}
        elif clickData:
            button_style = {'display': 'block'}
        else:
            button_style = {'display': 'none'}

    fig = px.bar(summary_df, x='Steward Name',
                 y=['Total Funds', '# Funds with "Fund Agreement" in filename', '# Funds with Fund Agreement signed'],
                 title="Fund Summary by Steward", barmode='group')
    fig.update_layout(clickmode='event+select')

    if clickData:
        selected_steward = clickData['points'][0]['x']
        for trace in fig.data:
            trace.marker.color = ['rgba(0, 116, 217, 0.7)' if x == selected_steward else 'rgba(172, 172, 172, 0.7)' for
                                  x in fig.data[0].x]

    return fig, button_style


# Callback for the drill-down table
@app.callback(
    Output('drill-down-table', 'children'),
    [Input('summary-graph', 'clickData'),
     Input('clear-button', 'n_clicks')]
)
def display_click_data(clickData, n_clicks):
    ctx = callback_context
    if not ctx.triggered:
        return "Click on a bar to see detailed fund information."

    input_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if input_id == 'clear-button' or clickData is None:
        return "Click on a bar to see detailed fund information."

    steward_name = clickData['points'][0]['x']
    steward_funds = df_funds[df_funds['Steward Name'] == steward_name]

    columns = [
        {"name": "Fund Name", "id": "Fund Name"},
        {"name": "Fund ID", "id": "funit_id"},
        {"name": "Total Files", "id": "total_files"},
        {"name": "Fund Agreement in Name", "id": "fund_agreement_in_name"},
        {"name": "Fund Agreement Signed", "id": "fund_agreement_signed"}
    ]

    return [
        html.H3(f"Funds for {steward_name}"),
        dash_table.DataTable(
            id='steward-funds-table',
            columns=columns,
            data=steward_funds[
                ['Fund Name', 'funit_id', 'total_files', 'fund_agreement_in_name', 'fund_agreement_signed']].to_dict(
                'records'),
            page_size=1000,
            style_table={'overflowX': 'auto'},
            style_cell={
                'minWidth': '0px', 'maxWidth': '180px',
                'whiteSpace': 'normal',
                'height': 'auto',
                'textAlign': 'left'
            },
            style_cell_conditional=[
                {'if': {'column_id': 'Fund Name'}, 'width': '30%'},
                {'if': {'column_id': 'funit_id'}, 'width': '15%'},
                {'if': {'column_id': 'total_files'}, 'width': '15%'},
                {'if': {'column_id': 'fund_agreement_in_name'}, 'width': '20%'},
                {'if': {'column_id': 'fund_agreement_signed'}, 'width': '20%'}
            ]
        )
    ]


if __name__ == '__main__':
    app.run_server(debug=True)