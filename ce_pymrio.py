# -*- coding: utf-8 -*-
"""
@author: tanch
"""

import pymrio
import pandas as pd
import time

#%%
'''
Pymrio processing functions
'''


exio3 = pymrio.parse_exiobase3(path= r"C:\Users\tanch\OneDrive - Universiteit Leiden\2022 Thesis\Data for Pycirk\IOT_2016_pxp.zip")

Z_MI = exio3.Z.index
Y_MI = exio3.Y.index

def get_ext_index(ext_df):
    stressor = ext_df.index.get_level_values('name')
    stressor.rename("stressor", inplace=True)
    unit = pd.DataFrame(data = ext_df.index.get_level_values('unit'), index=stressor,columns=['unit'])
    return stressor, unit

def make_io(scenario_munch,mrio_munch):
    io = pymrio.IOSystem()
    start = time.time()
    print("creating Z")
    io.Z = pd.DataFrame(scenario_munch.Z,index = exio3.Z.index,columns=exio3.Z.columns, )
    print("creating Y")
    io.Y = pd.DataFrame(scenario_munch.Y,index = exio3.Y.index,columns=exio3.Y.columns, )
    
    ext = {
        'W' : 'primary inputs',
        'E' : 'emission extensions',
        'M' : 'material extensions',
        'R' : 'resource extensions',   
        }
    
    ext_dict = {}
    for key in ext.keys():
        print(f"creating {key} ext")
        stressor, unit = get_ext_index(mrio_munch[key])
        df_F = pd.DataFrame(scenario_munch[key], index=(stressor), columns=(io.Z.columns))
        if key != "W":
            df_FY = pd.DataFrame(scenario_munch[key+'Y'], index=(stressor), columns=(io.Y.columns))
        elif key == 'W':
            df_FY = None
        ext_dict[key] = pymrio.Extension(name=ext[key], F=df_F, F_Y=df_FY, unit=unit)
        
    
    io.E = ext_dict['E']
    io.M = ext_dict['M']
    io.R = ext_dict['R']
    io.W = ext_dict['W']
    
    end = time.time()
    print(f'done in {end-start}s!')
    print(str(io))
    
    return io