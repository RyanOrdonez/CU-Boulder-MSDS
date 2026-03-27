"""
SpaceX Launch Dashboard — Plotly Dash Application

Features:
  1. Dropdown to select a launch site (or all sites)
  2. Pie chart showing success counts for the selected site
  3. Range slider to filter by payload mass
  4. Scatter plot of Payload vs Outcome colored by Booster Version
"""

import pandas as pd
import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import plotly.express as px

# ── Load data ────────────────────────────────────────────────────────────────
spacex_df = pd.read_csv(
    "https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/"
    "IBM-DS0321EN-SkillsNetwork/datasets/spacex_launch_dash.csv"
)

max_payload = spacex_df['Payload Mass (kg)'].max()
min_payload = spacex_df['Payload Mass (kg)'].min()

# ── Build app ────────────────────────────────────────────────────────────────
app = dash.Dash(__name__)

site_options = [{'label': 'All Sites', 'value': 'ALL'}] + [
    {'label': site, 'value': site}
    for site in spacex_df['Launch Site'].unique()
]

app.layout = html.Div([
    html.H1("SpaceX Launch Records Dashboard",
            style={'textAlign': 'center', 'color': '#503D36', 'fontSize': 40}),

    # Dropdown for launch site selection
    dcc.Dropdown(
        id='site-dropdown',
        options=site_options,
        value='ALL',
        placeholder='Select a Launch Site',
        searchable=True,
        style={'width': '80%', 'margin': 'auto', 'padding': '3px',
               'fontSize': '20px', 'textAlignLast': 'center'}
    ),
    html.Br(),

    # Pie chart
    html.Div(dcc.Graph(id='success-pie-chart')),
    html.Br(),

    html.P("Payload range (kg):", style={'fontSize': 18, 'paddingLeft': '10%'}),

    # Payload range slider
    dcc.RangeSlider(
        id='payload-slider',
        min=0,
        max=10000,
        step=1000,
        marks={i: f'{i}' for i in range(0, 10001, 2500)},
        value=[min_payload, max_payload]
    ),
    html.Br(),

    # Scatter chart
    html.Div(dcc.Graph(id='success-payload-scatter-chart')),
])


# ── Callbacks ────────────────────────────────────────────────────────────────
@app.callback(
    Output('success-pie-chart', 'figure'),
    Input('site-dropdown', 'value')
)
def update_pie(entered_site):
    if entered_site == 'ALL':
        fig = px.pie(
            spacex_df,
            values='class',
            names='Launch Site',
            title='Total Successful Launches by Site'
        )
    else:
        filtered = spacex_df[spacex_df['Launch Site'] == entered_site]
        outcome_counts = filtered['class'].value_counts().reset_index()
        outcome_counts.columns = ['class', 'count']
        outcome_counts['label'] = outcome_counts['class'].map({1: 'Success', 0: 'Failure'})
        fig = px.pie(
            outcome_counts,
            values='count',
            names='label',
            title=f'Success vs Failure — {entered_site}'
        )
    return fig


@app.callback(
    Output('success-payload-scatter-chart', 'figure'),
    [Input('site-dropdown', 'value'),
     Input('payload-slider', 'value')]
)
def update_scatter(entered_site, payload_range):
    low, high = payload_range
    mask = (spacex_df['Payload Mass (kg)'] >= low) & (spacex_df['Payload Mass (kg)'] <= high)

    if entered_site == 'ALL':
        filtered = spacex_df[mask]
    else:
        filtered = spacex_df[mask & (spacex_df['Launch Site'] == entered_site)]

    fig = px.scatter(
        filtered,
        x='Payload Mass (kg)',
        y='class',
        color='Booster Version Category',
        title=f'Payload vs Outcome — {"All Sites" if entered_site == "ALL" else entered_site}',
        labels={'class': 'Launch Outcome (1=Success)'}
    )
    return fig


# ── Run ──────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    app.run(debug=True, port=8050)
