import pandas as pd
from pandas.api.types import CategoricalDtype
import numpy as np
# import xlsxwriter

import dash
from dash import dcc
from dash import html
from dash.dependencies import Input, Output

import dash_bootstrap_components as dbc

import plotly.graph_objs as go

# ===================================
# Key parameters
release_date = 'July 2023'
# release_date = 'February 2024'
# ===================================
def sort_status(df):
    """
    convert column 'Status' to categorical
    https://pandas.pydata.org/pandas-docs/stable/user_guide/categorical.html
    """
    # for gas & oil
    status_order = [
        'operating',
        'mothballed',
        'announced',
        'pre-construction',
        'construction',
        'shelved',
    ]
    
    df['Status'] = df['Status'].astype(
        CategoricalDtype(status_order, ordered=False)
    )    
    df = df.sort_values(by=['Country', 'Status', 'Year'])
    
    return df

# ===================================
layout_chosen = '2 columns'  # options: '1 column', '2 columns'
# filepath = 'https://github.com/GlobalEnergyMonitor/GOGPT-dashboard/blob/main/data/Global%20Oil%20and%20Gas%20Plant%20Tracker%20(GOGPT)%20compiled%202023-08-18%20-%20processed%20for%20Dash%202023-09-18_1621.xlsx?raw=true'
filepath = 'https://github.com/GlobalEnergyMonitor/GOGPT-dashboard/blob/main/data/Global%20Oil%20and%20Gas%20Plant%20Tracker%20(GOGPT)%20compiled%202023-08-18%20-%20processed%20for%20Dash%202023-10-17_1906.xlsx?raw=true'
# filepath = 'https://github.com/GlobalEnergyMonitor/GOGPT-dashboard/blob/main/data/Global%20Oil%20and%20Gas%20Plant%20Tracker%20(GOGPT)%20compiled%20own-par%202024-02-16%20-%20processed%20for%20Dash%202024-02-16_1845.xlsx?raw=true'

# # ===================================
dash_data_xl = pd.ExcelFile(filepath, engine='openpyxl')

gogpt_map = pd.read_excel(dash_data_xl, sheet_name='map')
gogpt_status = pd.read_excel(dash_data_xl, sheet_name='status')
gogpt_age = pd.read_excel(dash_data_xl, sheet_name='age')
gogpt_add = pd.read_excel(dash_data_xl, sheet_name='additions')

# clean status df

gogpt_status = sort_status(gogpt_status)

# create list of countries to choose from (GEM country names)
# data in gogpt_status is most complete; 
# for example, gogpt_status includes Albania, which only has cancelled units

gogpt_country_list = gogpt_status['Country'].sort_values().unique().tolist()
if gogpt_country_list[0] != 'all':
    if 'all' in gogpt_country_list:
        gogpt_country_list.remove('all')
    else:
        pass

    gogpt_country_list_for_dropdown = ['all'] + gogpt_country_list

# ===================================
# ### Create country dropdown menu

# create list of dicts needed for dropdown menu

dropdown_options_list_of_dicts = []  # initialize
for country in gogpt_country_list_for_dropdown:
    dropdown_options_list_of_dicts += [{'label': country, 'value': country}] 

# create dropdown menu
country_dropdown = dcc.Dropdown(
    id='country_dropdown',
    options=dropdown_options_list_of_dicts,
    value='all', # default starting value
    placeholder='Select a country' # only shows up if user clears entry
)

# ===================================
# ===================================
# ## Create graphs

# ===================================
# ### Choropleth map
# from https://plotly.com/python/choropleth-maps/

def create_chart_choro(gogpt_map, sel_country):    
    # Get the maximum value to cap displayed values
    min_val = int(gogpt_map['capacity log10 + 1'].min())
    max_val = int(gogpt_map['capacity log10 + 1'].max())

    # Prepare the range of the colorbar
    values = [i for i in range(min_val, max_val+2)]
    ticks = [10**i for i in values]

    # set resolution
    if sel_country == 'all':
        sel_resolution = 110
        # this drives update using fitbounds
        gogpt_map_sel = gogpt_map
        
    else:        
        # for showing individual countries, set higher resolution (smaller scale features) 
        sel_resolution = 50
        # this drives update using fitbounds
        gogpt_map_sel = gogpt_map[gogpt_map['Country'] == sel_country]
        
    # create map
    fig_map = go.Figure(
        data=go.Choropleth(
            # aspects that don't change with sel_country
            colorscale='Viridis',
            # colorbar_title = "Capacity (MW)", # don't use colorbar_title if specifying colorbar with dict, as below
            colorbar={
                'title': 'Capacity (MW)',
                'tickvals': values,
                'ticktext': ticks,
            },
            locationmode='ISO-3', # set of locations match entries in 'locations'
            zauto=False,
            zmin=min_val,
            zmax=max_val,
            
            # aspects that do change with sel_country:
            locations=gogpt_map_sel['iso_alpha'],
            z=gogpt_map_sel['capacity log10 + 1'], # data to be color-coded
            # use separate column for hover text with original capacity values
            # (data for choropleth is log scale)
            hovertemplate=gogpt_map_sel['hover_text']
    ))

    # assign title and arrange
    fig_map.update_layout(
        title_text='<b>Operating Gas & Oil Power Capacity by Country</b>',
        # use margin to get title placement correct
        margin={'r': 100, 't': 100, 'l': 60, 'b': 150},
        title_x=0.5,
        dragmode=False,
        geo=dict(
                showframe=False,
                showcoastlines=False,
                projection_type='equirectangular',
                resolution=sel_resolution,
                # fitbounds="locations", # TO DO: check if this isn't needed, because of step below
                visible=True,
            ),
    )
    fig_map.add_annotation(dict(font=dict(color='dark blue',size=12),
                        x=.05,
                        y=-0.6,
                        # height=4,
                        align='left',
                        # width=800,
                        showarrow=False,
                        text=
                            "Total operating gas and oil power capacity by country as " + 
                            "<br>of latest data release (see release date below). Hover over " +
                            "<br>each country on the map to see total operating capacity. " +
                            "<br>Filter to select a specific country. " + 
                            "<br>Read our <a href='https://globalenergymonitor.org/projects/global-oil-gas-plant-tracker/methodology/'>Methodology page</a> for more information on our " +
                            "<br>data inclusion criteria and methodology.",
                        textangle=0,
                        xanchor='left',
                        xref= 'paper',
                        yref= 'paper',
                        # borderwidth= 800
                        )
    )

    # based on: https://plotly.com/python/choropleth-maps/
    # referred by: https://plotly.com/python/map-configuration/#automatic-zooming-or-bounds-fitting
    # need visible=True to show all country outlines; explained in: https://plotly.com/python/map-configuration/
    fig_map.update_geos(fitbounds="locations", visible=True)
                            
    return fig_map


def update_chart_choro(fig_map, sel_country):
    sel_locations = gogpt_map[gogpt_map['Country']==sel_country]['iso_alpha']
    fig_map.update_layout(locations=sel_locations)
    fig_map.update_geos(fitbounds="locations", visible=True)
    return fig_map


# initialize with global view
fig_map = create_chart_choro(
    gogpt_map=gogpt_map,
    sel_country='all'
)

# ===================================
# ### Total Capacity by Plant Status
# (stacked bar, vertical)

# format of color dictionary follows GreenInfo map code
gogpt_map_colors = {
  'announced': {'id': 0, 'text': 'Announced', 'color': '#1a0e5c'},
  'pre-construction': {'id': 1, 'text': 'pre-construction', 'color': '#00ba89'},
  'construction': {'id': 2, 'text': 'Construction', 'color': '#cb2359'},
  'shelved': {'id': 3, 'text': 'Shelved', 'color': '#f8693e'},
  'operating': {'id': 4, 'text': 'Operating', 'color': '#ffb421'},
  'mothballed': {'id': 5, 'text': 'Mothballed', 'color': '#ebff45'},
}


def create_chart_by_status(gogpt_status, sel_country):
    fig_status = go.Figure() # initialize
    # select data for sel_country
    df = gogpt_status[gogpt_status['Country']==sel_country]
    statuses = df['Status'].unique().tolist()
    
    for status in statuses:
        df_status = df[df['Status']==status]
        color_status = gogpt_map_colors[status]['color']

        fig_status.add_trace(go.Bar(
            x=df_status['Year'], 
            y=df_status['Capacity (MW)'], 
            name=status, 
            marker_color=color_status,
            hovertemplate=status + ': %{y:,.0f} MW<extra></extra>' # Capacity
        ))

    fig_status.update_layout(
        barmode='stack',
        title='<b>Gas & Oil Power Capacity by Status</b>',
        title_x=0.5,
        margin={'r': 10, 't': 100, 'l': 100, 'b': 150},
        yaxis=dict(
            title='Megawatts (MW)',

        ),

        xaxis = dict(
        tickmode = 'array',
        tickvals = [2021.5, 2022, 2022.5, 2023, 2023.5], #2024
        ticktext = ['H1 2021', 'H2 2021', 'H1 2022*\u200b', 'H2 2022', 'H1 2023'], #'H2 2023'
        tickangle=0
        ),

        legend=dict(
            # orientation='h',
            yanchor='top',
            # y=-0.1,
            y=1,
            xanchor='left',
            x=1,
            traceorder='normal',
        ),
    )

    fig_status.add_annotation(dict(font=dict(color='dark blue',size=10),
                            x=.02,
                            y=-0.23,
                            align='left',
                            showarrow=False,
                            text=
                                '*\u200bIn H1 2022 status category "proposed" was expanded '+
                                '<br>into "announced" and "pre-construction"',
                            textangle=0,
                            xanchor='left',
                            xref= 'paper',
                            yref= 'paper',
                            )
    )
    fig_status.add_annotation(dict(font=dict(color='dark blue',size=12),
                        x=.02,
                        y=-0.7,
                        align='left',
                        showarrow=False,
                        text=
                            'Oil and gas power capacity by status over bi-annual intervals. ' +
                            '<br>Hover over each graph section to see total capacity for the specific ' +
                            '<br>status category for that time period. ' +
                            '<br>Read our <a href="https://globalenergymonitor.org/projects/global-oil-gas-plant-tracker/methodology/">Methodology page</a> for definitions of our status categories.',
                        textangle=0,
                        xanchor='left',
                        xref='paper',
                        yref='paper')
    )
    return fig_status


fig_status = create_chart_by_status(
    gogpt_status=gogpt_status, 
    sel_country='all')


# ===================================
# ### Operating Coal Power by Plant Age and Type 
# By "type" it means technology
# 'Gas Turbine', 'Steam Turbine', 'Combined Cycle', 'Internal Combustion Combined Cycle', 'Integrated Solar Combined Cycle', 'Allam-Fetvedt Cycle', 'Internal Combustion'
# from Data Color Picker (learnui)
age_tech_pallette = {
    'Gas Turbine': '#1a0e5c', 
    'Steam Turbine': '#007c91', 
    'Combined Cycle': '#00ba89', 
    'Internal Combustion Combined Cycle': '#cb2359', 
    'Integrated Solar Combined Cycle': '#f8693e', 
    'Allam-Fetvedt Cycle': '#ffb421',
    'Internal Combustion': '#ebff45',
    'Unknown': '#808080' # grey
}
# not used: '#444e86', # medium blue; instead put in grey

def create_chart_age_type(gogpt_age, sel_country):
    fig_age = go.Figure() # initialize

    technologies_in_order = [
    'Gas Turbine',
    'Steam Turbine',
    'Combined Cycle',
    'Internal Combustion Combined Cycle',
    'Integrated Solar Combined Cycle',
    'Allam-Fetvedt Cycle',
    'Internal Combustion',
    'Unknown'
    ]
    
    gogpt_age_sel_country = gogpt_age[gogpt_age['Country'] == sel_country].drop('Country', axis=1)
    gogpt_age_sel_country = gogpt_age_sel_country.set_index('Decade')
    decades = ['0-9 years', '10-19 years', '20-29 years', '30-39 years', '40-49 years', '50+ years']
    for decade in decades:
        if decade not in gogpt_age_sel_country.index:
            print(f'decade not in {sel_country}: {decade}') # for debuggins
            new_row_df = pd.DataFrame(
                data=[[0]*len(technologies_in_order)], 
                columns=technologies_in_order, 
                index=[decade]
                )
            # old way was causing errors so switched to updated concat syntax 
            # gogpt_age_sel_country = gogpt_age_sel_country.concat(new_row_df) https://github.com/keras-team/keras-io/issues/1518#issuecomment-1741335324
            gogpt_age_sel_country = pd.concat([gogpt_age_sel_country, new_row_df], ignore_index=True) # axis=1

    gogpt_age_sel_country = gogpt_age_sel_country.sort_index()

    gogpt_age_sel_country.columns.tolist()
    for technology in technologies_in_order:
        fig_age.add_trace(go.Bar(
            name=technology,
            x=gogpt_age_sel_country[technology], 
            y=gogpt_age_sel_country.index, 
            orientation='h',
            marker_color=age_tech_pallette[technology],
            hovertemplate=technology + ': %{x:,.0f} MW<extra></extra>',
        ))

    fig_age.update_layout(
        barmode='stack',
        title='<b>Operating Gas & Oil Power Capacity by Age and Type</b>',
        title_x=0.5,
        margin={'r': 100, 't': 100, 'l': 60, 'b': 150},
        xaxis=dict(
            title='Megawatts (MW)',
        ),
        legend=dict(
            orientation='h',
            yanchor='top',
            y=1,
            xanchor='left',
            x=1,
            traceorder='normal',
        ),

    )

    fig_age.add_annotation(dict(font=dict(color='dark blue',size=12),
                        x=.02,
                        y=-0.6,
                        showarrow=False,
                        text=
                        'Operating oil and gas power capacity by age and technology type.' +
                        '<br>Hover over each graph section to see total capacity' +
                        '<br>for the specific technology type for that age range.',                        textangle=0,
                        xanchor='left',
                        align='left',
                        xref='paper',
                        yref='paper')
    )

    # reverse axis to put youngest at the top
    fig_age['layout']['yaxis']['autorange'] = "reversed"
    
    return fig_age

fig_age = create_chart_age_type(
    gogpt_age=gogpt_age, 
    sel_country='all')

# ===================================
# ### Coal Power Additions and Retirements
# * Has bars and line; see https://plotly.com/python/graphing-multiple-chart-types/

def create_chart_additions(gogpt_add, sel_country):
    fig_add = go.Figure() # initialize figure

    df = gogpt_add[gogpt_add['Country']==sel_country]
    
    df = df.rename(columns={
        'Added (MW)': 'Added'
    })
    
    for status in ['Added']:
        df_add = df[['Year', status]].set_index('Year')

        color = '#680266'
        fig_add.add_trace(go.Bar(
            x=df_add.index, 
            y=df_add[status], # values are capacities (MW)
            name=status, 
            marker_color=color,
            hovertemplate=status + ': %{y:,.0f} MW<extra></extra>',
        ))


    # update overall layout
    fig_add.update_layout(
        barmode='stack', 
        title='<b>Gas & Oil Power Capacity Added</b>',
        title_x=0.5,
        # showlegend=True,
        margin={'r': 10, 't': 100, 'l': 100, 'b': 150},
        yaxis=dict(
            title='Megawatts (MW)',
        ),
        xaxis = dict(
            title='Plant Start Year',
            tickmode = 'array',
            tickvals = [2002, 2005, 2008, 2011, 2014, 2017, 2020, 2023],
            ticktext = ['2002', '2005', '2008', '2011', '2014', '2017', '2020', '2023']
        ),
        # legend=dict(
        #     orientation='h',
        #     yanchor='top',
        #     y=-0.2,
        #     xanchor='left',
        #     x=.4,
        #     traceorder='normal',
        # ),
    )
    fig_add.add_annotation(dict(font=dict(color='dark blue',size=10),
                         x=.02,
                         y=-0.2,
                         showarrow=False,
                         text='*\u200bH1 2023 data',
                         textangle=0,
                         xanchor='left',
                         xref='paper',
                         yref='paper')
    )
    fig_add.add_annotation(dict(font=dict(color='dark blue',size=12),
                        x=.02,
                        y=-0.7,
                        align='left',
                        showarrow=False,
                        text=
                            'Annual oil and gas power capacity additions ' +
                            '<br>by start year. Hover over each graph section to see ' +
                            '<br>total oil and gas power capacity addition for the specific ' +
                            '<br>start year. Note that this graph subtracts any capacity for ' +
                            '<br>plants that have retired since 2000.',
                        textangle=0,
                        xanchor='left',
                        xref='paper',
                        yref='paper')
    )

    return fig_add

# initialize chart with global data
fig_add = create_chart_additions(
    gogpt_add=gogpt_add,
    sel_country='all'
)

# ===================================
# Create app & server

app = dash.Dash(
    __name__, 
    external_stylesheets=[dbc.themes.BOOTSTRAP] 
    )

# title based on: https://community.plotly.com/t/how-do-you-set-page-title/40115
app.title = "Gas & Oil Power dashboard"
server = app.server

# ===================================
# Create graphs of charts

dropdown_title = html.H5(children='Select a country:', style={'marginLeft': 10, 'marginRight': 10, 'marginBottom': 10, 'marginTop': 25}), 
download_text = html.H6(children='Download figure data:')
download_button = html.Button("Download Excel file", id="btn_xlsx"),

choro_graph = dcc.Graph(
    id='chart_choro', 
    figure=fig_map, 
    config={'displayModeBar': False},
    style={'marginLeft': 'auto', 'marginRight': 'auto', 'marginBottom': 'auto', 'marginTop': 'auto'},
    )

status_graph = dcc.Graph(
    id='chart_status', 
    figure=fig_status,
    config={'displayModeBar': False},
    style={'marginLeft': 'auto', 'marginRight': 'auto', 'marginBottom': 'auto', 'marginTop': 'auto'},

    )

age_graph = dcc.Graph(
    id='chart_age', 
    figure=fig_age,
    config={'displayModeBar': False},
        style={'marginLeft': 'auto', 'marginRight': 'auto', 'marginBottom': 'auto', 'marginTop': 'auto'},

    )

add_graph = dcc.Graph(
    id='chart_add', 
    figure=fig_add,
    config={'displayModeBar': False},
        style={'marginLeft': 'auto', 'marginRight': 'auto', 'marginBottom': 'auto', 'marginTop': 'auto'},

    )

# ===================================
# Define layout

if layout_chosen == '1 column':
    # 1-column version
    app.layout = dbc.Container(fluid=True, children=[
        dbc.Row([dbc.Col(country_dropdown)], align='center'),
        dbc.Row([dbc.Col(choro_graph)], align='center'),
        dbc.Row([dbc.Col(status_graph)], align='center'),
        dbc.Row([dbc.Col(age_graph)], align='center'),
        dbc.Row([dbc.Col(add_graph)], align='center'),
    ],
    )
elif layout_chosen == '2 columns':
    # 2-column version
    # download based on: https://dash.plotly.com/dash-core-components/download
    app.layout = dbc.Container(fluid=True, children=[
        dbc.Row([
            dbc.Col([
                dbc.Row(dropdown_title),
                dbc.Row(country_dropdown),
            ], md=4),
            dbc.Col([], xl=5) # spacer
            # # section for download button:
            # dbc.Col([
            #     dbc.Row(download_text),
            #     dbc.Row(download_button),
            #     html.Div(id='dynamic-dropdown-container', children=[]), # EXPERIMENTAL
            #     dcc.Download(id="download-dataframe-xlsx"),
            # ], md = 2),
        ]),
        dbc.Row([
            dbc.Col(choro_graph, xl=5, align="start"), 
            dbc.Col(status_graph, xl=5, align="start"),
        ]),
        dbc.Row([
            dbc.Col(age_graph, xl=5, align="start"),
            dbc.Col(add_graph, xl=5, align="start"),
        ]),
        dbc.Row([
            dbc.Col([
                html.H6(f'Data from Global Oil and Gas Plant Tracker, {release_date} release'),
            ], align="evenly"),
        ]),
    ],
    )
#dash.Dash(external_stylesheets=
@app.callback(
    Output('chart_choro', 'figure'),
    Output('chart_status', 'figure'),
    Output('chart_age', 'figure'),
    Output('chart_add', 'figure'),
    Input('country_dropdown', 'value'), 
)
def update_figure(sel_country):
    fig_map = create_chart_choro(
        gogpt_map=gogpt_map, 
        sel_country=sel_country
    )
    fig_status = create_chart_by_status(
        gogpt_status=gogpt_status, 
        sel_country=sel_country
    )
    fig_age = create_chart_age_type(
        gogpt_age=gogpt_age,
        sel_country=sel_country
    )
    fig_add = create_chart_additions(
        gogpt_add=gogpt_add, 
        sel_country=sel_country
    )
    fig_map.update_layout(transition_duration=500)
    fig_status.update_layout(transition_duration=500)
    fig_age.update_layout(transition_duration=500)
    fig_add.update_layout(transition_duration=500)

    return fig_map, fig_status, fig_age, fig_add

# # Section for download file
# @app.callback(
#     Output("download-dataframe-xlsx", "data"), # for download button
#     Output('dynamic-dropdown-container', 'children'), # EXPERIMENTAL
#     # Output('country_dropdown', 'value'),
#     Input("btn_xlsx", "n_clicks"), # for download button
#     prevent_initial_call=True,
# )
# def func(n_clicks, children):
#     """For download button"""
#     # hardcode value for testing
#     sel_country = 'Bangladesh'

#     status_sel = gogpt_status.copy()[gogpt_status['Country']==sel_country]
#     status_sel['Status'] = status_sel['Status'].astype(str)
#     status_sel = status_sel.set_index(['Year', 'Status'])[['Capacity (MW)']]
#     status_sel = status_sel.unstack(-1)

#     # drop level for columns to remove 'Capacity (MW)'
#     status_sel = status_sel.droplevel(0, axis=1)
#     status_sel = status_sel.reset_index()
#     status_sel = status_sel.rename(columns={'Year': 'year'})

#     # TO DO: write header row that contains Country (and more)
#     table_title = f"{sel_country} - Coal Power Capacity (MW) by Status"

#     # Create a Pandas Excel writer using XlsxWriter as the engine.
#     writer = pd.ExcelWriter('Global Coal Plant Tracker dashboard export.xlsx', engine='xlsxwriter')

#     # Write each dataframe to a different worksheet.
#     status_sheet_name = 'Capacity by Status'
#     status_sel.to_excel(
#         writer, 
#         sheet_name=status_sheet_name, 
#         startrow=1, 
#         index=False
#     )
#     sheet1 = writer.sheets[status_sheet_name]
#     sheet1.write('A1', table_title)
#     # df2.to_excel(writer, sheet_name='Sheet2')
#     # df3.to_excel(writer, sheet_name='Sheet3')

#     # Close the Pandas Excel writer and output the Excel file.
#     writer.save()

#     return dcc.send_file('Global Coal Plant Tracker dashboard export.xlsx')

#     # return dcc.send_data_frame(
#     #     status_sel.to_excel, 
#     #     "gogpt_dash_export.xlsx", 
#     #     sheet_name="Capacity by Status", 
#     #     index=False,
#     # )

if __name__ == '__main__':
    app.run_server()
