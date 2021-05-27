import arcpy
# import os
# import getpass
# from logging_decorator import *
from scratchPrep import *
from BoundaryRoadCalcs import *

# Dirs
root = os.path.dirname(os.path.abspath(__file__))
inputs = os.path.join(root, 'input')
outpath = os.path.join(root, 'output')
temp = os.path.join(root, 'temp')
scratch = os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb')

for file in os.listdir(inputs):
    if 'Roads.shp' in file:
        roadShp = file
    elif 'pophu.shp' in file:
        popBlks = file
    elif 'Polygon.shp' in file:
        poly = file
roads = {
    '_path': os.path.join(inputs, roadShp),
    '_scratch': os.path.join(scratch, roadShp.split('.')[0])
}
blocks = {
    '_path': os.path.join(inputs, popBlks),
    '_scratch': os.path.join(scratch, popBlks.split('.')[0])
}
poly = {
    '_path': os.path.join(inputs, poly),
    '_scratch': os.path.join(scratch, poly.split('.')[0])
}

