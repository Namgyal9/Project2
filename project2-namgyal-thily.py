#!/usr/bin/env python
# coding: utf-8

# # Libraries maybe needed to downloaded. you can do !pip install Library name

# In[1]:


from dash import Dash, dcc, html, Input, Output, State, dash_table
import dash_daq as daq
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import plotly.colors as colors


# ## code for joining and reading the excel sheet

# In[2]:


df = pd.read_excel("All Four Factors.xlsx")
df2 = pd.read_excel("Opportunity Analysis Data.xlsx")
df3 = pd.read_excel("Strength AnalysisData.xlsx")
df4 = pd.read_excel("Threat Analysis Data.xlsx")
df5 = pd.read_excel("Weakness Analysis Data.xlsx")

merged_df = pd.concat([df, df2,df3,df4,df5],ignore_index = True)
merged_df.fillna(0, inplace = True)
merged_df.drop_duplicates(inplace = True)

categories = ["Opportunity","Strength","Threat","Weakness"]
direction = 3
Attributes = ["EST. VALUE IN CURRENCY", "MIN PROB ADJUSTED VALUE", "AVERAGE PROB ADJUSTED VALUE",
           "MAX PROB ADJUSTED VALUE", "REALISTIC PROB ADJUSTED VALUE",
           "3 POINT BASED PROB ADJUSTED VALUE", "PERT BASED PROB ADJUSTED VALUE"]
print(categories)
filtered_df = merged_df[merged_df["CATEGORY"].isin(categories)]
df_sorted = filtered_df.sort_values(by="CATEGORY").groupby("CATEGORY").apply(lambda x: x.reset_index(drop=True))
print(len(filtered_df))
df_sorted


# ## example table of data that is being plotted in visual combined graph

# In[3]:


df_sorted2 = filtered_df.groupby("FACTOR TYPE").sum().reset_index()
subset_df = df_sorted2.iloc[:2]
sum_values = subset_df.sum()
diff_row = pd.DataFrame(data=sum_values).T
diff_row.index = [2]
df_sorted2 = pd.concat([df_sorted2, diff_row])
df_sorted2.loc[df_sorted2["FACTOR TYPE"] == "NEGATIVEPOSITIVE", "FACTOR TYPE"] = "diff"
df_sorted2


# ## example table of data that is being plotted for S W O T

# In[4]:


# Example
category = "Strength"
category_df = df_sorted.loc[category]
category_df


# ## Main code that generates the graphs and sets unique variety of colors
# ### we are using the colors from the pyplot library and using the plotly graph objects to make stacked and grouped bar charts.
# 

# In[5]:


x_values = Attributes

def fig_create(graph_name, mode, status):  # mode decides which type of graph , status is so we can reuse the code.
    if status == 0:
        category_df = df_sorted.loc[graph_name]
        grouped_values = "PARAM NAME"
    elif status == 1:
        category_df = df_sorted2.iloc[:graph_name]
        grouped_values = "FACTOR TYPE"
        graph_name = "VISUALS COMBINED"

    traces = []
    color_scale = colors.qualitative.D3 
    color_index = 0  
    for param_name in category_df[grouped_values].unique():
        y_values = category_df[category_df[grouped_values] == param_name][Attributes].values.tolist()[0]
        color = color_scale[color_index % len(color_scale)]  
        trace = go.Bar(x=x_values, y=y_values, name=param_name, marker=dict(color=color))
        traces.append(trace)
        color_index += 1  

    fig = go.Figure(data=traces)
    fig.update_xaxes(tickangle=-45, tickfont=dict(family="Arial", size=10))
    fig.update_layout(title=f'<b>{graph_name}</b>')
    fig.update_layout(barmode=mode)

    return fig


# # Rendering everything here and connecting through callbacks, used only toggle however. Also the look and design of the dash board is here. 

# In[ ]:


app = Dash(__name__)

app.layout = html.Div([
    html.H1("SWOT Analysis", style={"text-align": "center"}),
    html.Br(),
    daq.ToggleSwitch(id='my-toggle-switch',label='Grouped | Stacked',labelPosition='bottom',style={"text-align": "center"}),
    html.Br(),
    html.Div([
    html.Div(children=[
        html.Div(children=[
            html.Label('POSITIVE', style={"font-size": "24px", "background": "green", "color": "white", "fontWeight": "bold", "textAlign": "center"}),
            dcc.Graph(id="graph1", figure=fig_create(categories[1],"stack",0)),
            html.Br(),
            dcc.Graph(id="graph2", figure=fig_create(categories[0],"stack",0))
        ], style={'padding': 10, 'flex': 1, 'textAlign': 'center'}),
        
        html.Div(children=[
            html.Label('NEGATIVE', style={"font-size": "24px", "background": "red", "color": "white", "fontWeight": "bold", "textAlign": "center"}),
            dcc.Graph(id="graph3", figure=fig_create(categories[3],"stack",0)),
            html.Br(),
            dcc.Graph(id="graph4", figure=fig_create(categories[2],"stack",0))
        ], style={'padding': 10, 'flex': 1, 'textAlign': 'center'})
    ], style={'display': 'flex', 'flex-direction': 'row', 'justify-content': 'center'}),
    
    html.Div(children=[
        html.Label('Cumulative Graphs', style={"font-size": "24px", "background": "blue", "color": "white", "fontWeight": "bold", "textAlign": "center"}),
        html.Br(),
        dcc.Graph(id="graph5", figure=fig_create(categories[0:2],"group",0)),
        dcc.Graph(id = "graph6", figure = fig_create(categories[2:4],"group",0)),
        dcc.Graph(id = "graph7",figure = fig_create(direction,"group",1))
    ], style={'padding': 10, 'textAlign': 'center', 'width': '100%'})
])])


@app.callback(
    Output("graph1", "figure"),
    Output("graph2", "figure"),
    Output("graph3", "figure"),
    Output("graph4", "figure"),
    Input("my-toggle-switch", "value")
)
def update_graphs(value):
    if value:
        mode = "stack"
    else:
        mode = "group"
    
    figure1 = fig_create(categories[1], mode,0)
    figure2 = fig_create(categories[0], mode,0)
    figure3 = fig_create(categories[3], mode,0)
    figure4 = fig_create(categories[2], mode,0)
        
    return figure1, figure2, figure3, figure4






if __name__ == '__main__':
    app.run_server()

