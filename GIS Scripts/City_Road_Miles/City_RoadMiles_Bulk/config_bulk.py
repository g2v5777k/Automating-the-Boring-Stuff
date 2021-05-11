import pandas as pd
import arcpy
import requests
import os
import getpass
import zipfile
from bs4 import BeautifulSoup

# Dirs
root = os.path.dirname(os.path.abspath(__file__))
inputs = os.path.join(root, 'input')
outpath = os.path.join(root, 'output')
temp = os.path.join(root, 'temp')
roadsDir = os.path.join(root, 'roads')
scratch = os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb')

# Create DFs of input cities and the fips codes
citiesList = os.path.join(inputs, 'cities_bulk.xlsx')
readList = pd.read_excel(citiesList, engine='openpyxl')
cityDF = pd.DataFrame(readList)
fipsList = os.path.join(inputs, 'us-state-ansi-fips.xlsx')
readFips = pd.read_excel(fipsList, converters={'fip': lambda x:str(x)}, engine='openpyxl') # keeps leading zeros
fipsDF = pd.DataFrame(readFips)

cityDF.State = cityDF.State.str.strip()
fipsDF.state_abbr = fipsDF.state_abbr.str.strip()

# create a merged file to get fips, create dict after
statesDict = {}
merged = pd.merge(cityDF, fipsDF, how='left', left_on=['State'], right_on=['state_abbr'])

if not os.listdir(outpath):
    for state, fip in zip(merged['State'].drop_duplicates(), merged['fip'].drop_duplicates()):
        statesDict[state] = fip
        path = os.path.join(str(outpath), str(state))
        os.mkdir(path)
else:
    print('Output folder is not empty, please remove all folders in output before running...')

# List of road features to remove
exlcudedRoads = [
        'S1100',
        'S1200',
        'S1630',
        'S1500',
        'S1710',
        'S1720',
        'S1730',
        'S1740',
        'S1750',
        'S1780',
        'S1820',
        'S1830'
    ]