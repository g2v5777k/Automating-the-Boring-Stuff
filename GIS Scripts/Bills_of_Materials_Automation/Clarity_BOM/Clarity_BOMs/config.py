import os
from os import path
import arcpy
import getpass
from openpyxl import load_workbook
from pathlib import Path
import PySimpleGUI as sg
from site import addsitedir
from sys import executable


root = Path.cwd()
input_dir = root / 'input'
# output = str(root / 'output')
cur_dir = os.path.join(os.path.dirname('__file__'))

# inputs
# input_gdb = r'C:\Users\ColtonLuttrell\Desktop\Desktop_Files\Python\Market_Specific_Scripts\Clarity_BOMs\input\Clarity_20201117.gdb'

# List of clarity features
clarity_features = ['Anchors', 'FiberCable', 'conduit', 'OLT_Boundaries', 'Structures', 'SpliceClosure', 'Slackloops', 'Strand']

# Workbook vars
cwd = os.getcwd()
template = os.path.join(cwd + '\\template', 'Clarity_BOM_Template.xlsx')
wb = load_workbook(filename=template)
ws = wb['Sheet1']

# sg initialization
interpreter = executable
path_to_interpreter = path.dirname(interpreter)
sitepkg = path_to_interpreter + "\\site-packages"
addsitedir(sitepkg)

gui = [[sg.Text('Input GDB', size=(20, 1)), sg.InputText(size=(70, 1)), sg.FolderBrowse()],
       [sg.Text('BOM Output Location:', size=(20, 1)), sg.InputText(size=(70, 1)), sg.FolderBrowse()],
       [sg.Text('OLT(s) Comma Separated:', size=(20, 1)), sg.InputText(size=(70, 1))],
       [sg.Submit(), sg.Exit()]]

window = sg.Window('Clarity BOM (Simplified)').Layout(gui)

# check if scratch.gdb exists on users machine, if not make one.
scratchExists = os.path.exists(os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb'))
if scratchExists == False:
    arcpy.CreateFileGDB_management(os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS'), 'scratch.gdb')
else:
    pass
scratch = os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb')


