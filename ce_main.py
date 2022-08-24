# -*- coding: utf-8 -*-
"""
Created on Mon Jun  6 13:00:04 2022

@author: tanch
"""

'''
i. Import libraries
'''
import pandas as pd
import pymrio
import pycirk
import matplotlib.pyplot as plt
import time
import seaborn as sns
# import altair as alt
from datetime import datetime
import numpy as np

# from ce_pymrio import make_io
from ce_pycirk import (create_analyse, run_pycirk, update_analyse, update_headers, 
category_bar, scenario_impacts_bar, impact_scenarios_bar, stripplot, update_dom_extr)

#%%
'''
1. Pycirk scenario setup
'''

### Initialisation ###

## Setting input data for Pycirk modelling 
cat_consump_list = ['global warming GWP100',
                    'Total Energy Use',
                    'Domestic Extraction',
                    # 'Land use',
                    'Water Withdrawal Blue - Total',
                    'Value Added',
                    'Employment']

## Plastic scenario directory ##
folder_plast = r"E:\Pycirk Output\tests\plastics"
## Construction scenario directory ##
folder_constn = r"E:\Pycirk Output\tests\construction"
folder_dict = {'plastics' : folder_plast, 
                'construction': folder_constn
                }

co_intervention = ["MP",
                   "PR",
                   "Modularity",
                   "Circular SC"]
pl_intervention = ['Prevention',
                   'RP',
                   'BQ&ER',
                   'Co-op VC']
levels = ["low","med","high"]

pl_cols = [f"{i}: {l}" for i in pl_intervention for l in levels]
co_cols = [f"{i}: {l}" for i in co_intervention for l in levels]

# column headers for result dataframes
header_dict = {'plastics': pl_cols,
               'construction': co_cols
               }

folder_data_main = r".\ref"
analyse = pd.read_excel(folder_data_main+'\scenarios_template.xlsx',"analyse",header=[3],nrows=(0))

#%% Modelling iterations
#initialise dictionary of results

region_consumption = "All"
region_production = "NL"
category_production = "All"

results_dict = {}   

for k,v in folder_dict.items():
    results_dict[k]={}
    analysis_df = pd.DataFrame()
    for category in cat_consump_list:
        category_consumption = category
        analysis_df = pd.concat([analysis_df,create_analyse(category_consumption,region_consumption,category_production,region_production,[3])])
        
    update_analyse(v+'\scenarios.xlsx', analysis_df)
    results = run_pycirk(v,save = False)
    results_dict[k] = results
    
# results_update = {}  
for k,v in results_dict.items():
    # updated_results = update_headers(v,header_dict[k])
    updated_results = update_dom_extr(v)
    results_dict[k] = updated_results
#     results_update[k] = updated_results


#%%
# results_dict = {'plastics': df_plastic,
#                 'construction': df_construction}
for key, df in results_dict.items():
    # stripplot(df,percent = False,)
    # stripplot(df,percent = True,)
    
    category_bar(df, percent = False)
    category_bar(df, percent = True)
    
    impact_scenarios_bar(df)
    
    scenario_impacts_bar(df, percent = False)
    scenario_impacts_bar(df, percent = True)

    
df_plastic = results_dict['plastics'].copy()
df_construction = results_dict['construction'].copy()
df_construction.drop(columns=["baseline"], inplace=True)

df_concat = pd.concat([df_plastic,df_construction], axis= 1)
scenario_impacts_bar(df_concat,percent=True)
