import pandas as pd
import numpy as np
from datetime import timedelta
from dash import Dash, html, dcc, Input, Output, dash_table
import dash_bootstrap_components as dbc

# Load the Excel file
file_path = "a.xlsx"
df = pd.read_excel(file_path)
df_attendance = pd.read_excel(file_path, sheet_name='Sheet2')

# Clean and prepare main dashboard data
df['Process Date'] = pd.to_datetime(df['Process Date'], errors='coerce')
df['UPT'] = pd.to_timedelta(df['UPT'].astype(str), errors='coerce')
df['Expected Time'] = pd.to_timedelta(df['Expected Time'].astype(str), errors='coerce')
df['Month'] = df['Process Date'].dt.strftime('%Y-%m')
df['Date'] = df['Process Date'].dt.date

# Expected UPT mapping
expected_upt_mapping = {
    'Confirmation': '0:04:37',
    'Line-Item': '0:01:53',
    'Order Entry': '0:04:10',
    'Prep Report': '01:15:00',
    'Complaint Report': '00:15:00',
    'QC Report': '02:20:00',
    'Dairy Purchase': '00:15:00',
    'Sales @9': '00:12:00',
    'MOV': '00:07:00',
    'Collate Report': '00:10:00',
    'E-FOOD': '00:04:37',
    'UFC': '00:04:37',
    'PW': '00:04:37',
    'Acquire': '00:03:42',
    'G & F': '00:01:00'
    
}
expected_upt_td = {k: pd.to_timedelta(v) for k, v in expected_upt_mapping.items()}

# Attendance sheet processing
df_attendance['Processor'] = df_attendance['Processor'].astype(str).str.strip()
df_attendance = df_attendance[~df_attendance['Processor'].isin(['0', '0.0', '', 'nan'])]


# Sum of working days from attendance count column
if 'count' in df_attendance.columns:
    total_working_days = df_attendance['count'].sum()
    print("Total Working Days from attendance count:", total_working_days)
# Dropdown options
available_months = sorted(df['Month'].dropna().unique())
available_employees = sorted(df['Processor'].dropna().unique()) + ['All']
attendance_months = sorted(df_attendance['Month'].dropna().unique())
attendance_processors = sorted(df_attendance['Processor'].unique())
attendance_processors = ['All'] + attendance_processors

def format_timedelta_as_hours(td):
    total_seconds = int(td.total_seconds())
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours:02}:{minutes:02}:{seconds:02}"

# App initialization
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
app.title = "Merged Dashboard"

# Layout
app.layout = dbc.Tabs([
    dbc.Tab(label="Productivity Dashboard", children=[
        dbc.Container([
            html.H2("Monthly Productivity Dashboard", className="text-center my-4"),
            dbc.Row([
                dbc.Col([
                    html.Label("Select Month"),
                    dcc.Dropdown(
                        options=[{'label': m, 'value': m} for m in available_months],
                        value=available_months[0],
                        id='month-dropdown'
                    )
                ], width=4),
                dbc.Col([
                    html.Label("Select Employee"),
                    dcc.Dropdown(
                        options=[{'label': e, 'value': e} for e in available_employees],
                        value=available_employees[0],
                        id='employee-dropdown'
                    )
                ], width=4),
            ], className="mb-4"),

            dbc.Card([
                dbc.CardBody([
                    html.H4(id='selected-info', className='card-title'),
                    dbc.Row([
                        dbc.Col(html.Div([html.H6("Actual UPT Time"), html.P(id='actual-upt')])),
                        dbc.Col(html.Div([html.H6("Expected UPT Time"), html.P(id='expected-upt')])),
                        dbc.Col(html.Div([html.H6("Difference"), html.P(id='upt-diff')])),
                        dbc.Col(html.Div([html.H6("Working Days"), html.P(id='working-days')])),
                        dbc.Col(html.Div([html.H6("Work % Efficiency"), html.P(id='upt-percent')])),
                    ])
                ])
            ], className="mb-4", style={'backgroundColor': 'white', 'boxShadow': '0 2px 8px rgba(0,0,0,0.1)', 'borderRadius': '8px'}),

            html.H4("Weekly Productivity Details", className="my-3"),
            dash_table.DataTable(
                id='weekly-table',
                columns=[
                    {"name": "Week", "id": "Week"},
                    {"name": "Actual Time", "id": "Actual Time"},
                    {"name": "Expected Time", "id": "Expected Time"},
                    {"name": "Difference", "id": "Difference"},
                    {"name": "Working Days", "id": "Working Days"},
                    {"name": "Efficiency", "id": "Efficiency"}
                ],
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'center'},
                style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'},
                page_size=10
            ),

            html.H4("Monthly Process-wise Productivity Details", className="mt-5"),
            dbc.Row([
                dbc.Col([
                    html.Label("Select Employee"),
                    dcc.Dropdown(
                        options=[{'label': e, 'value': e} for e in available_employees],
                        id='employee-process-dropdown',
                        value=available_employees[0]
                    )
                ], width=4),
            ], className="mb-4"),
            dash_table.DataTable(
                id='monthly-process-table',
                columns=[
                    {"name": "Process", "id": "Process"},
                    {"name": "Total Time", "id": "Total Time"},
                    {"name": "Total Items", "id": "Total Items"},
                    {"name": "Avg Time", "id": "Avg Time"},
                    {"name": "Expected UPT", "id": "Expected UPT"},
                    {"name": "Total Production", "id": "Total Production"},
                ],
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'center'},
                style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'},
                page_size=15
            ),

            html.H4("Daily Process-wise Productivity Details", className="mt-5"),
            dbc.Row([
                dbc.Col([html.Label("Select Date"), dcc.Dropdown(id='date-dropdown')], width=4),
                dbc.Col([html.Label("Select Employee"),
                         dcc.Dropdown(id='emp-date-dropdown', options=[{'label': e, 'value': e} for e in available_employees], value=available_employees[0])], width=4),
            ], className="mb-4"),

            dash_table.DataTable(
                id='daily-table',
                columns=[
                    {"name": "Process", "id": "Process"},
                    {"name": "Total Time", "id": "Total Time"},
                    {"name": "Total Items", "id": "Total Items"},
                    {"name": "Avg Time", "id": "Avg Time"},
                    {"name": "Expected UPT", "id": "Expected UPT"},
                    {"name": "Total Production", "id": "Total Production"},
                ],
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'center'},
                style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'},
                page_size=15
            ),

            html.Div(id='daily-summary', className='mt-4 mb-5')
        ], fluid=True, style={'backgroundColor': '#e6f7ff', 'minHeight': '100vh'})
    ]),
    dbc.Tab(label="Attendance Tracker", children=[
        html.Div([
            html.H2("Attendance Tracker Dashboard", style={'textAlign': 'center'}),
            html.Div([
                html.Label("Select Month:"),
                dcc.Dropdown(id='att-month-dropdown', options=[{'label': m, 'value': m} for m in attendance_months])
            ], style={'width': '45%', 'display': 'inline-block'}),
            html.Div([
                html.Label("Select Processor:"),
                dcc.Dropdown(id='att-processor-dropdown', options=[{'label': p, 'value': p} for p in attendance_processors], value='All')
            ], style={'width': '45%', 'display': 'inline-block', 'marginLeft': '5%'}),
            html.Br(), html.Br(),
            dash_table.DataTable(
                id='att-data-table',
                columns=[{'name': col, 'id': col} for col in df_attendance.columns],
                data=df_attendance.to_dict('records'),
                page_size=20,
                style_table={'overflowX': 'auto'},
                style_cell={'textAlign': 'left', 'padding': '5px'},
                style_header={'backgroundColor': 'lightgrey', 'fontWeight': 'bold'}
            )
        ], style={'padding': '20px'})
    ])
])

# Callbacks
@app.callback(
    Output('selected-info', 'children'),
    Output('actual-upt', 'children'),
    Output('expected-upt', 'children'),
    Output('upt-diff', 'children'),
    Output('working-days', 'children'),
    Output('upt-percent', 'children'),
    Output('weekly-table', 'data'),
    Input('month-dropdown', 'value'),
    Input('employee-dropdown', 'value')
)
def update_monthly_dashboard(month, employee):
    filtered = df[df['Month'] == month]
    if employee != 'All':
        filtered = filtered[filtered['Processor'] == employee]
    filtered = filtered[filtered['Process'] != 'Break']
    if filtered.empty:
        return f"No data for {employee} in {month}", '-', '-', '-', '0', '-', []

    actual_upt = filtered['UPT'].sum()
    expected_upt = filtered['Expected Time'].sum()
    diff = actual_upt - expected_upt
    working_days = filtered['Process Date'].dt.date.nunique()
    percent = (expected_upt.total_seconds() / actual_upt.total_seconds() * 100) if actual_upt.total_seconds() else 0

    filtered['Week'] = filtered['Process Date'].dt.to_period('W').apply(lambda r: r.start_time.strftime('%Y-%m-%d'))
    weekly = filtered.groupby('Week').agg({
        'UPT': 'sum',
        'Expected Time': 'sum',
        'Process Date': lambda x: x.dt.date.nunique()
    }).rename(columns={'Process Date': 'Working Days'}).reset_index()
    weekly['Difference'] = weekly['UPT'] - weekly['Expected Time']
    weekly['Efficiency'] = weekly.apply(lambda r: (r['Expected Time'].total_seconds() / r['UPT'].total_seconds() * 100)
                                        if r['UPT'].total_seconds() else 0, axis=1)

    weekly_data = [dict(
        Week=row['Week'],
        **{k: (format_timedelta_as_hours(row[k]) if isinstance(row[k], timedelta)
              else f"{row[k]:.0f}%") for k in ['UPT', 'Expected Time', 'Difference', 'Working Days', 'Efficiency']}
    ) for _, row in weekly.iterrows()]

    for d in weekly_data:
        d["Actual Time"] = d.pop("UPT")

    return (
        f"Summary for {employee} - {month}",
        format_timedelta_as_hours(actual_upt),
        format_timedelta_as_hours(expected_upt),
        format_timedelta_as_hours(diff) if diff.total_seconds() >= 0 else f"-{format_timedelta_as_hours(abs(diff))}",
        str(working_days),
        f"{percent:.0f}%",
        weekly_data
    )

@app.callback(
    Output('monthly-process-table', 'data'),
    Input('month-dropdown', 'value'),
    Input('employee-process-dropdown', 'value')
)
def update_monthly_process_table(month, employee):
    filtered = df[(df['Month'] == month)]
    if employee != 'All':
        filtered = filtered[filtered['Processor'] == employee]

    grouped = filtered.groupby('Process').agg({'UPT': 'sum', 'Process': 'count'}).rename(columns={'Process': 'Total Items'}).reset_index()
    table_data = []
    for _, row in grouped.iterrows():
        process = row['Process']
        total_items = row['Total Items']
        total_upt = row['UPT']
        avg_time = total_upt / total_items if total_items else timedelta()
        expected = expected_upt_td.get(process, None)
        production = expected * total_items if expected else total_upt
        table_data.append({
            'Process': process,
            'Total Time': format_timedelta_as_hours(total_upt),
            'Total Items': total_items,
            'Avg Time': format_timedelta_as_hours(avg_time),
            'Expected UPT': format_timedelta_as_hours(expected) if expected else '-',
            'Total Production': format_timedelta_as_hours(production)
        })
    return table_data

@app.callback(
    Output('date-dropdown', 'options'),
    Output('date-dropdown', 'value'),
    Input('month-dropdown', 'value')
)
def update_date_options(selected_month):
    filtered_dates = df[df['Month'] == selected_month]['Date'].dropna().unique()
    sorted_dates = sorted(filtered_dates)
    options = [{'label': str(d), 'value': str(d)} for d in sorted_dates]
    default_value = str(sorted_dates[0]) if sorted_dates else None
    return options, default_value

@app.callback(
    Output('daily-table', 'data'),
    Output('daily-summary', 'children'),
    Input('date-dropdown', 'value'),
    Input('emp-date-dropdown', 'value')
)
def update_daily_dashboard(date, employee):
    date = pd.to_datetime(date).date()
    filtered = df[df['Date'] == date]
    if employee != 'All':
        filtered = filtered[filtered['Processor'] == employee]

    grouped = filtered.groupby('Process').agg({'UPT': 'sum', 'Process': 'count'}).rename(columns={'Process': 'Total Items'}).reset_index()
    table_data, total_time, total_production = [], timedelta(), timedelta()
    for _, row in grouped.iterrows():
        process = row['Process']
        total_items = row['Total Items']
        total_upt = row['UPT']
        avg_time = total_upt / total_items if total_items else timedelta()
        expected = expected_upt_td.get(process, None)
        production = expected * total_items if expected else total_upt
        total_time += total_upt
        total_production += production
        table_data.append({
            
            'Process': process,
            'Total Time': format_timedelta_as_hours(total_upt),
            'Total Items': total_items,
            'Avg Time': format_timedelta_as_hours(avg_time),
            'Expected UPT': format_timedelta_as_hours(expected) if expected else '-',
            'Total Production': format_timedelta_as_hours(production)
        })
    eff = (total_production.total_seconds() / total_time.total_seconds() * 100) if total_time.total_seconds() else 0
    summary = html.Div([
        html.Span(f"Total Time: {format_timedelta_as_hours(total_time)}"),
        html.Span(f" | Total Production: {format_timedelta_as_hours(total_production)}", style={'marginLeft': '15px'}),
        html.Span(f" | Efficiency: {eff:.0f}%", style={'marginLeft': '15px', 'color': '#007BFF'})
    ], style={'fontSize': '18px', 'fontWeight': 'bold'})
    return table_data, summary

@app.callback(
    Output('att-data-table', 'data'),
    Input('att-month-dropdown', 'value'),
    Input('att-processor-dropdown', 'value')
)
def update_attendance_table(month, processor):
    filtered_df = df_attendance.copy()
    if month:
        filtered_df = filtered_df[filtered_df['Month'] == month]
    if processor and processor != 'All':
        filtered_df = filtered_df[filtered_df['Processor'] == processor]
    return filtered_df.to_dict('records')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8050)
