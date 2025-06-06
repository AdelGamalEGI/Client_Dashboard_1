
import dash
from dash import html, dcc
import dash_bootstrap_components as dbc
import pandas as pd
import plotly.graph_objects as go
import os
from datetime import datetime

# Load data
file_path = 'Project_Management_Template_Updated.xlsx'
df_workstreams = pd.read_excel(file_path, sheet_name='Workstreams')
df_risks = pd.read_excel(file_path, sheet_name='Risk_Register')
df_issues = pd.read_excel(file_path, sheet_name='Issue_Tracker')
df_resources = pd.read_excel(file_path, sheet_name='Resources')
df_budget = pd.read_excel(file_path, sheet_name='Budget_vs_Actual')

# Preprocessing
df_workstreams['Planned Start Date'] = pd.to_datetime(df_workstreams['Planned Start Date'], errors='coerce')
df_workstreams['Planned End Date'] = pd.to_datetime(df_workstreams['Planned End Date'], errors='coerce')
df_workstreams['Progress %'] = pd.to_numeric(df_workstreams['Progress %'], errors='coerce')
df_resources['Allocated/Used Hours'] = pd.to_numeric(df_resources['Allocated/Used Hours'], errors='coerce')

# Calculate Planned % based on time elapsed
today = pd.Timestamp.today()
start_month = pd.Timestamp(today.replace(day=1))
end_month = start_month + pd.offsets.MonthEnd(1)

duration = (df_workstreams['Planned End Date'] - df_workstreams['Planned Start Date']).dt.days
elapsed = (today - df_workstreams['Planned Start Date']).dt.days
df_workstreams['Planned %'] = ((elapsed / duration) * 100).clip(lower=0, upper=100).fillna(0)

# KPI Summary
tasks_this_month = df_workstreams[(df_workstreams['Planned Start Date'] <= end_month) & (df_workstreams['Planned End Date'] >= start_month)]
num_tasks = tasks_this_month.shape[0]
open_issues = df_issues[df_issues['Status'].str.strip().str.lower() == 'open']
num_open_issues = open_issues.shape[0]
open_risks = df_risks[df_risks['Status'].str.strip().str.lower() == 'open']
num_open_risks = open_risks.shape[0]

if 'High' in open_risks['Risk Score'].values:
    risk_color = 'danger'
elif 'Medium' in open_risks['Risk Score'].values:
    risk_color = 'warning'
else:
    risk_color = 'warning' if num_open_risks > 0 else 'secondary'

# Workstream progress chart
def get_color(delta):
    if delta <= 15:
        return 'green'
    elif delta <= 30:
        return 'orange'
    return 'red'

ws_chart = go.Figure()
for _, row in df_workstreams.groupby('Work-stream').mean(numeric_only=True).reset_index().iterrows():
    delta = abs(row['Planned %'] - row['Progress %'])
    color = get_color(delta)
    ws_chart.add_trace(go.Bar(name='Planned', x=[row['Planned %']], y=[row['Work-stream']], orientation='h', marker_color='blue'))
    ws_chart.add_trace(go.Bar(name='Actual', x=[row['Progress %']], y=[row['Work-stream']], orientation='h', marker_color=color))
ws_chart.update_layout(barmode='overlay', title='Workstream Progress', height=300)

# Task table
task_table = dbc.Table.from_dataframe(tasks_this_month[['Activity Name', 'Progress %']], striped=True, bordered=True, hover=True)

# Determine active team members from Budget_vs_Actual linked with Workstreams
df_merged = df_budget.merge(
    df_workstreams[['Activity Name', 'Planned Start Date', 'Planned End Date']],
    on='Activity Name', how='left'
)

df_merged['Planned Start Date'] = pd.to_datetime(df_merged['Planned Start Date'], errors='coerce')
df_merged['Planned End Date'] = pd.to_datetime(df_merged['Planned End Date'], errors='coerce')

active_rows = df_merged[
    (df_merged['Planned Start Date'] <= end_month) & (df_merged['Planned End Date'] >= start_month)
]

assigned1 = active_rows['Assigned Person 1'].dropna().str.strip().str.lower()
assigned2 = active_rows['Assigned Person 2'].dropna().str.strip().str.lower()
active_names = pd.concat([assigned1, assigned2]).unique()

df_resources['Person Name Lower'] = df_resources['Person Name'].str.strip().str.lower()
active_members = df_resources[df_resources['Person Name Lower'].isin(active_names)]

# Photo map
photo_mapping = {
    "Lavjit Singh": "lavjit.jpg",
    "Adel Gamal": "adel.jpg",
    "Don Sunny": "don.jpg",
    "Ganesh Shinde": "ganesh.jpg",
    "Samuel Ezannaya": "samuel.jpg",
    "Stefan Stroobants": "stefan.jpg",
    "Jaco Roesch": "jaco.jpg",
    "Gustav Brand": "gustav.jpg",
    "Seyed Khali": "seyed.jpg"
}

def member_card(name, role):
    img_file = photo_mapping.get(name)
    if img_file:
        img_tag = html.Img(src=f"/assets/{img_file}", height="45px", style={'borderRadius': '50%'})
    else:
        img_tag = html.Div("👤", style={'fontSize': '2rem'})
    return dbc.Card(
        dbc.Row([
            dbc.Col(img_tag, width='auto'),
            dbc.Col([
                html.Div(html.Strong(name)),
                html.Div(html.Small(role, className='text-muted'))
            ])
        ], align='center'),
        className='mb-2 p-2 shadow-sm'
    )

member_cards = [member_card(row['Person Name'], row['Role']) for _, row in active_members.iterrows()]

# KPI Summary Card
kpi_card = dbc.Card([
    dbc.CardHeader("KPI Summary"),
    dbc.CardBody([
        dbc.Row([
            dbc.Col([
                html.Div([
                    html.H2(f"{num_tasks}", className="text-primary text-center mb-0"),
                    html.P("Tasks This Month", className="text-muted text-center")
                ])
            ]),
            dbc.Col([
                html.Div([
                    html.H2([
                        dbc.Badge(f"{num_open_risks}", color=risk_color, className="px-3 py-2", pill=True),
                    ], className="text-center mb-0"),
                    html.P("Open Risks", className="text-muted text-center")
                ])
            ]),
            dbc.Col([
                html.Div([
                    html.H2(f"{num_open_issues}", className="text-primary text-center mb-0"),
                    html.P("Open Issues", className="text-muted text-center")
                ])
            ]),
        ], justify="center")
    ])
], className="p-3 shadow-sm")

# Layout
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server

app.layout = dbc.Container([
    html.H2('Client Dashboard', className='text-center my-4'),
    dbc.Row([
        dbc.Col([kpi_card], width=6),
        dbc.Col([
            dbc.Card([
                dbc.CardHeader('Workstream Progress'),
                dbc.CardBody([dcc.Graph(figure=ws_chart)])
            ], className="shadow-sm")
        ], width=6)
    ], className='mb-4'),
    dbc.Row([
        dbc.Col([
            dbc.Card([
                dbc.CardHeader('Tasks This Month'),
                dbc.CardBody([task_table])
            ], className="shadow-sm")
        ], width=6),
        dbc.Col([
            dbc.Card([
                dbc.CardHeader('Active Team Members'),
                dbc.CardBody(member_cards)
            ], className="shadow-sm")
        ], width=6)
    ])
], fluid=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8050))
    app.run_server(host='0.0.0.0', port=port)
