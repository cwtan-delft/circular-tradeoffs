# -*- coding: utf-8 -*-
"""
Created on Mon Jun  6 13:00:04 2022

@author: tanch
"""

'''
i. Set priority chains
'''
# create dictionary for priority chain modelling data
priority_chain_dict = {'plastics' : {},
                'construction': {}
                }

## Plastic scenario directory ##
priority_chain_dict['plastics']['folder'] = r".\tests\plastics"
## Construction scenario directory ##
priority_chain_dict['construction']['folder'] = r".\tests\construction"

folder_data_ref = r".\ref"

#%%
'''
ii. Import libraries
'''
import pandas as pd
import pymrio
import pycirk
import matplotlib.pyplot as plt
import time
import seaborn as sns
from datetime import datetime
import numpy as np

# from ce_pymrio import make_io
from ce_pycirk import (create_analyse, run_pycirk, update_analyse, 
                       update_headers, category_bar, scenario_impacts_bar, 
                       impact_scenarios_bar, stripplot, update_dom_extr,
                       save_results_to_excel)

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
priority_chain_dict['plastics']['column'] = pl_cols
priority_chain_dict['construction']['column'] = co_cols


analyse = pd.read_excel(folder_data_ref+'\\scenarios_template.xlsx',"analyse",
                        header=[3],nrows=(0))

#%% 
'''
Modelling iterations
'''

#initialise dictionary of results
region_consumption = "All"
region_production = "NL"
category_production = "All"


#iterate through set of scenarios in specified folders in folder_dict
#generate the scenarios file and run pycirk
for k,v in priority_chain_dict.items():
    analysis_df = pd.DataFrame()
    for category in cat_consump_list:
        category_consumption = category
        analysis_df = pd.concat(
            [analysis_df,create_analyse(category_consumption,region_consumption,
                                        category_production,region_production,[3])
             ])
    folder = v['folder']
    update_analyse(folder+'\\scenarios.xlsx', analysis_df)
    results = run_pycirk(folder,save = False)
    updated_results = update_headers(results,v['column'])
    updated_results = update_dom_extr(updated_results)
    priority_chain_dict[k]['results'] = updated_results
    
# to output sceanrio EXIOBASE tables to excel
# save_results_to_excel(results_dict)


#%%

for k, v in priority_chain_dict.items():
    df = v['results']
    # stripplot(df,percent = False,)
    # stripplot(df,percent = True,)
    
    savestring = v['folder'] + '\\plots\\plot'

    
    category_bar(df, percent = False,savestring = savestring)
    category_bar(df, percent = True,savestring = savestring)
    
    impact_scenarios_bar(df)
    
    scenario_impacts_bar(df, percent = False,savestring = savestring)
    scenario_impacts_bar(df, percent = True,savestring = savestring)

    
df_plastic = priority_chain_dict['plastics']['results'].copy()
df_construction = priority_chain_dict['construction']['results'].copy()
df_construction.drop(columns=["baseline"], inplace=True)

df_concat = pd.concat([df_plastic,df_construction], axis= 1)
scenario_impacts_bar(df_concat,percent=True,savestring = savestring)
