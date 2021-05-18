import arcpy
import os
from driveV3_dl import authenticate
from datetime import datetime
import pandas as pd
import shutil
# dirs
root = os.path.dirname(os.path.abspath(__file__))
inputPath = os.path.join(root, 'inputGDB')
date = datetime.now()
print(f'Run Date: {date.strftime("%m/%d/%Y")}')
print('------------------')
# start, end = str(date.strftime('%Y-%m-%d')) + ' 00:00:00', str(date.strftime('%Y-%m-%d')) + ' 23:59:00'
start, end  = '2021-05-12 00:00:00', '2021-05-12 23:59:00'
output = os.path.join(root, 'output')
genName = 'RDOF_FieldTracker'
gdb = os.path.join(inputPath, genName + '.gdb')
sde_con = r'C:\GISMO\AEG\RDOF\Collector\RDOFCollector\RDOFCollector.sde'
# fcs to use
fcsDict = {
'AddressVerification': ['AddressVerification_Today','AddressVerification_NameDissolve', 'AddressVerification_CBGs', 'AddressVerification_CBGs_Counts', 'AddressVerification_Lifetime'],
'ProblemPoint': ['ProblemPoint_Today','ProblemPoint_NameDissolve', 'ProblemPoint_CBGs', 'ProblemPoint_CBGs_Counts', 'ProblemPoint_Lifetime'],
'RoadPoints': ['RoadPoints_Today','RoadPoints_NameDissolve', 'RoadPoints_CBGs', 'RoadPoints_CBGs_Counts', 'RoadPoints_Lifetime'],
'UtilityPoint': ['UtilityPoint_Today','UtilityPoint_NameDissolve', 'UtilityPoint_CBGs', 'UtilityPoint_CBGs_Counts', 'UtilityPoint_Lifetime']
}
countDict = {}
userList = []
userListTot =[]
# Initialize xl
writer = pd.ExcelWriter(os.path.join(output, 'Fielders_Point_Counts_{}.xlsx').format('05122021')) #format(date.strftime('%m%d%Y')))


