# -*- coding: utf-8 -*-
"""
Created on Tue Jul 13 17:08:26 2021

@author: schongf
"""
import pandas as pd
import numpy as np

#Input your RunROMEO excel file address.
excel = r'C:\Users\schongf\Desktop\Python trial\RunRomeo_v1.7 - SARJAR_3R2A v6.xlsm'
Input1 = pd.read_excel (excel, sheet_name='XREF_Master')

#Input a map from ROMEO to PIMS tags. See excel for format. 
Input2 = pd.read_excel ('PIMS ROM map.xlsx')

#Dataframe manipulation for Extraction and data cleanup
df = Input1[['Alias', 'VariableName','UOM','DataSource_Tag', 'DataSource_Tag_Description']]\
    .rename(columns={'DataSource_Tag':'Plant','VariableName':'ROMEO', 'DataSource_Tag_Description':'Description'})
df2 = df[df['Alias'].str.contains('Select|\*|SG|MMB_IN|MMB_OUT')==False]
df3 = df2.merge(Input2,how='left', on='ROMEO')
df4 = df3[["Alias", "Plant", "ROMEO", "PIMS", "UOM", "Description"]]
df4['UOM'] = df4['UOM'].fillna('pending')
df4['Plant'] = df4['Plant'].fillna(df4['Description'])
df4['Plant'] = df4['Plant'].fillna('Calculated')
df4 = df4.fillna(0)
df5 = df4[["Alias", "Plant", "ROMEO", "PIMS", "UOM"]]

#Array used for iteration and output generation is nf1
nf1 = df5.to_numpy()

#Create empty lists
Tag = []
Alias = []
UOM = []
Platform = []
Platform_Sel = ["Plant Data", "ROMeo Reference Model OMV", "PIMS OMV"]
Dimension = []

#Create Dictionary to extract unit types. See csv for format. You may add new units.
Dict_Archive = pd.read_csv('Unit_Dict.csv')
Dim_Dict = Dict_Archive.set_index('Unit')['Type'].to_dict()

#Looping through the array nf1 to create the outputs
for row in nf1:
    b = row.tolist()
    Dimension.append(Dim_Dict[b[4]])
    for x in range(1,4):
        if b[x] != 0:
            Alias.append(b[0])
            Tag.append(b[x])
            Platform.append(Platform_Sel[x-1])
            UOM.append(b[4])

#Name your output files. You will have two output files for pg1 and pg2 of MV SQL DB Xref Bulk Uploader
dest_text = 'output'

#Creating pandas output to export as csv
####For Variable Platform page####
pd_output = pd.DataFrame(list(zip(Alias,Platform,Tag,UOM)),columns= ["Alias","Platform","Tag","UOM"])
Destination_address2 = dest_text+'pg2'+'.csv'
pd_output.to_csv(Destination_address2)
    
####For Variables page#####
df5 = df4[["Alias","UOM"]]
df5.insert(2, "Dimension", Dimension, True)

Destination_address1 = dest_text+'pg1'+'.csv'
df5.to_csv(Destination_address1)
print('csv files produced')
#Done#