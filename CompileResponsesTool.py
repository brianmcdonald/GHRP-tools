from os import listdir
from openpyxl import load_workbook
from os.path import isfile, join
import os
import glob
import pandas as pd
import xlrd

# file locations, replace as required and ensure the template file of headers is present 
responsefolder = "/mnt/c/Users/bmcdonald/International Organization for Migration - IOM/PRD - Documents/03 Crisis Response/15_2020_COVID-19/01 Appeals/Global HRP/Monitoring Framework/Responses from Missions/"
homefolder = "/home/brian/Code/GHRP"

writer = pd.ExcelWriter('temp2-round4.xlsx', engine = 'openpyxl')
files = os.listdir(responsefolder)
os.chdir(homefolder)
# template = pd.read_excel("/mnt/c/Users/bmcdonald/OneDrive - International Organization for Migration - IOM/Code/GHRP/template.xlsx", 'Sheet1',header=0)
template = pd.read_excel("/home/brian/Code/GHRP/template.xlsx")
master = template
os.chdir(responsefolder)
for f in files:
    try:
        data = pd.read_excel(f, 'Round 4', index_col=None, header=0)
        country = f.split('- ' )[1]
        data["Country"] = country.split('.xl')[0]
        master = master.append(data[0:7])
    except:
        print(f, 'error')    
master = master.iloc[:, 0:18]

os.chdir(homefolder)
master.to_excel(writer, sheet_name='Round 4', index=False)
writer.save()
writer.close()
print("Compilation of responses complete")