"""
Phil - Steward Treemap.py

Generates a treemap of fund balances by steward and fund type.
"""

import psycopg2
import sqlalchemy
import pandas as pd
import plotly.express as px
import datetime
import os

POSTGRES_HOST = '10.20.141.105'
POSTGRES_USER = 'postgres'
CSUITE_POSTGRES_DB = 'csuite'
# POSTGRES_PASSWORD = os.environ['CSUITE_POSTGRES_SQL_SERVER_PWD']
POSTGRES_PASSWORD = os.environ['POSTGRES_DW_PASSWORD']

save_directory = (
    r"C:\Users\Mitchell.Hollberg\Community Foundation for Greater Atlanta\IT - Documents\Public\Reports"
)

today = datetime.date.today()

# Data is 2 days old
data_as_of = today - datetime.timedelta(days=2)
# Format today's date as mm/dd/yyyy
data_as_of = data_as_of.strftime('%m/%d/%Y')

conn = psycopg2.connect(host=POSTGRES_HOST,
                        database=CSUITE_POSTGRES_DB,
                        user=POSTGRES_USER,
                        password=POSTGRES_PASSWORD)

cursor = conn.cursor()

# Create sqlalchemy connection
engine = sqlalchemy.create_engine(f'postgresql://{POSTGRES_USER}'
                                  f':{POSTGRES_PASSWORD}@{POSTGRES_HOST}/'
                                  f'{CSUITE_POSTGRES_DB}')
#
# # Load to dataframe
# df = pd.read_sql_query("SELECT * FROM public.funits limit 100;", engine)
# print(df.head())
#
# # Cleanup
# cursor.close()
# conn.close()
# engine.dispose()

funds_sql = """
select
	funit_id,
	funits.fgroup_id,
	fgroups."name" as fundtype,
	funits."name",

	short_name,
	cast(funits.created_ts as date) as fund_created,
	current_fundbalance,
	funits.steward_employee_id,
	e."name" as steward_name

from public.funits as funits

left join public.fgroups as fgroups
on funits.fgroup_id = fgroups.fgroup_id

left join employees_v e 
on e.employee_id = funits.steward_employee_id 

where close_ts is null
"""

funds_df = pd.read_sql_query(funds_sql, engine)

# Replace null steward names with "Unassigned"
funds_df['steward_name'] = funds_df['steward_name'].fillna('Unassigned')

# Remove funds with zero balance
funds_df = funds_df[funds_df['current_fundbalance'] > 0]

# Divide by 1,000,000 to get millions
funds_df['current_fundbalance'] = funds_df['current_fundbalance'] / 1000000

# Add a "CFGA" column that will serve as a top-level category
funds_df['CFGA'] = 'CFGA'

# Build treemap of fund balances by steward by fund type
fig = px.treemap(funds_df,
                 path=['CFGA', 'steward_name', 'fundtype', 'name'],
                 color='fundtype',
                 values='current_fundbalance',
                 custom_data=['steward_name', 'fundtype', 'name',
                              'current_fundbalance']
                 )

fig.update_layout(
    title={
        'text': 'Fund Balances by Steward and Fund Type (in $ Millions).'
    }
)

fig.add_annotation(
    text=f'Data as of {data_as_of}',
    x=0.5,
    y=-0.1,
    showarrow=False,
    font=dict(size=12)
)

# Update hovertemplate
fig.update_traces(
    hovertemplate=("<b>%{customdata[2]}</b><br>Fund Type: %{customdata[1]}"
                   "<br>Steward: %{customdata[0]}<br>Value: $%{customdata[3]:,.1f}M"
                   "<extra></extra>")
)

# fig.show()

# Save file as *.aspx
fig.write_html(fr'{save_directory}\PortfolioSummary.aspx')


# Build treemap of fund balances by fund type then steward
fig2 = px.treemap(funds_df,
                 path=['CFGA', 'fundtype', 'steward_name', 'name'],
                 color='fundtype',
                 values='current_fundbalance',
                 custom_data=['fundtype', 'steward_name', 'name',
                              'current_fundbalance']
                 )

fig2.update_layout(
    title={
        'text': 'Fund Balances by Fund Type by Steward (in $ Millions).'
    }
)

fig2.add_annotation(
    text=f'Data as of {data_as_of}',
    x=0.5,
    y=-0.1,
    showarrow=False,
    font=dict(size=12)
)

# Update hovertemplate
fig2.update_traces(
    hovertemplate=("<b>%{customdata[2]}</b><br>Fund Type: %{customdata[1]}"
                   "<br>Steward: %{customdata[0]}<br>Value: $%{customdata[3]:,.1f}M"
                   "<extra></extra>")
)

# fig.show()

# Save file as *.aspx
fig2.write_html(fr'{save_directory}\PortfolioSummarybyFundType.aspx')


moo = 'boo'


