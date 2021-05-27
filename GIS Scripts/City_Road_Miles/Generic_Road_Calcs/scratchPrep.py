import arcpy
import os
import getpass
# from config_generic import *

'''
Attempt to create a generic clear gdb, and transfer to scratch.gdb script. Relies on a config file where paths to 
shapefiles are defined in a dictionary with '_path' being the key.
'''

class scratchPrep():
    def __init__(self, scratch=os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb')):
        '''
        :param scratch: path to scratch.gdb
        '''
        self.scratch = scratch
    def clear_gdb(self):
        '''
        Clear contents of scratch.gdb.
        :return: None.
        '''
        print('Clearing Workspace\n')
        arcpy.env.workspace = self.scratch
        fc_list = arcpy.ListFeatureClasses()
        fc_tables = arcpy.ListTables()
        try:
            for fc in fc_list:
                arcpy.Delete_management(self.scratch + '\\' + fc)
            for table in fc_tables:
                arcpy.Delete_management(self.scratch + '\\' + table)
        except TypeError:
            pass
    def transfer(self, dictList = []):
        '''
        Transfer shapefiles in our directories to scratch.gdb for ease of use.
        :return: None.
        '''
        print('Transferring input files...\n')
        shps = []
        for d in dictList:
            for shp, path in d.items():
                if '_path' in shp:
                    shps.append(path)
                else:
                    pass
        for shp in shps:
            arcpy.FeatureClassToGeodatabase_conversion(shp, self.scratch)
        print('\n')


