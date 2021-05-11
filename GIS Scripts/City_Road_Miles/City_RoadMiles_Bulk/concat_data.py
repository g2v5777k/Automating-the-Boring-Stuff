import os
import pandas as pd
import arcpy
import getpass
import time
import shutil

'''
Script - concat_data
Created by - Colton Luttrell
Date Created - 4/2/2021

Requirements:
    1) Calculate_RoadMiles_Bulk.py, and data within output folder in root.
    
I/O:
    Input: data within root/output
    Output: concats .xlsxs and gdbs within state dirs

Version History:
    v1 - 4/2/2021 - CL
    

'''
root = os.path.dirname(os.path.abspath(__file__))
outpath = os.path.join(root, 'output')
temp = os.path.join(root, 'temp')
scratch = os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb')
def clear_gdb():
    '''
    Clear contents of scratch.gdb.
    :return:
    None.
    '''
    print('Clearing Workspace\n')
    fc_list = arcpy.ListFeatureClasses()
    fc_tables = arcpy.ListTables()
    try:
        for fc in fc_list:
           arcpy.Delete_management(scratch + '\\' + fc)
        for table in fc_tables:
            arcpy.Delete_management(scratch + '\\' + table)
    except TypeError:
        pass

def concat_xlsx():
    '''
    Navigates through each state dir in ouput, and concats it.
    :return:
    xlsx with all data in output
    '''

    print('Concatenating excels...')
    totCity = []
    totCBs = []
    totDFCity = []
    totDFCBs =[]

    for folder in os.listdir(outpath):
        if '.xlsx' not in folder:
            stateDir = os.path.join(outpath, folder)
            os.chdir(stateDir)
            cityPath = os.path.join(stateDir, 'Cities_Polygons_RoadMiles.xlsx')
            cbsPath = os.path.join(stateDir, 'CBs_in_Cities_RoadMiles.xlsx')
            totCity.append(cityPath)
            totCBs.append(cbsPath)
    for xl in totCity:
        readCity = pd.read_excel(xl)
        dfCity = pd.DataFrame(readCity)
        totDFCity.append(dfCity)
    for xl in totCBs:
        readCBs = pd.read_excel(xl)
        dfCBs = pd.DataFrame(readCBs)
        totDFCBs.append(dfCBs)

    concat_dir = os.path.join(outpath, 'All_Muni_Deliverable_Data')
    os.mkdir(concat_dir)
    os.chdir(concat_dir)
    totalCity = pd.concat(totDFCity)
    # clean up sheets
    del totalCity['OBJECTID_12']
    del totalCity['NAMELSAD']
    del totalCity['LSAD']
    del totalCity['CLASSFP']
    del totalCity['PCICBSA']
    del totalCity['PCINECTA']
    del totalCity['MTFCC']
    del totalCity['FUNCSTAT']
    del totalCity['INTPTLAT']
    del totalCity['INTPTLON']
    del totalCity['stname']
    del totalCity['fip']
    del totalCity['state_abbr']
    del totalCity['OBJECTID_1']
    del totalCity['Shape_Length']
    del totalCity['Shape_Area']
    totalCity.rename(columns={'LENGTH_GEO': 'RoadMiles'})
    totalCity.to_excel('All_Muni_Cities_Miles.xlsx', index=False)

    totalCBs = pd.concat(totDFCBs)
    del totalCBs['OBJECTID_12']
    del totalCBs['TARGET_FID']
    del totalCBs['stname']
    del totalCBs['fip']
    del totalCBs['state_abbr']
    del totalCBs['Shape_Area_1']
    del totalCBs['Shape_Length_12']
    del totalCBs['Shape_Length']
    del totalCBs['Shape_Area']
    totalCBs.rename(columns={'LENGTH_GEO': 'RoadMiles'})
    totalCBs.to_excel('All_Muni_CBs_Miles.xlsx', index=False)

def merge_gdbs():
    '''
    Merges all gdbs in output folders, code is not optimal. Took a lot of back and forth to get desired results. Will clean up in v2.
    :return:
    One gdb with fcs of all states cities of interest
    '''

    print('Merging GDBs....')
    # lists of fcs in the gdbs
    # concat_dir = os.path.join(outpath, 'All_Muni_Deliverable_Data')
    # os.mkdir(concat_dir)
    # os.chdir(concat_dir)
    roadsGDB = []
    cbsGDB =[]
    citiesGDB = []
    finalFCs = ['All_Muni_Roads.shp', 'All_Muni_CBs.shp', 'All_Muni_City_Polys.shp']
    n = 0
    # walk directories to get fcs in invdividual state folder's gdbs. Will find a better dir walking solution in v2.
    for folder in os.listdir(outpath):
        n += 1
        if 'All_Muni_Deliverable_Data' not in folder:
            if '.gdb' not in folder:
                stateDir = os.path.join(outpath, folder)
                os.chdir(stateDir)
                gdbPath = os.path.join(stateDir, 'city_results.gdb')
                arcpy.env.workspace = gdbPath
                fcs = arcpy.ListFeatureClasses()
                for fc in fcs:
                    arcpy.Rename_management(fc, fc + '_' + str(n))

                fcs2 = arcpy.ListFeatureClasses()
                # convert to temp shps with new fc names
                for fc2 in fcs2:
                    arcpy.FeatureClassToGeodatabase_conversion(fc2, temp)

    # append shps in temp dir to lists to convert to new output gdb
    for shp in os.listdir(temp):
        if '.shp' in shp:
            if '.shp.xml' not in shp:
                if 'Polygons' in shp:
                    citiesGDB.append(shp)
                if 'CBs' in shp:
                    cbsGDB.append(shp)
                if 'Roads_City_Polys' in shp:
                    roadsGDB.append(shp)
    arcpy.CreateFileGDB_management(os.path.join(outpath, 'All_Muni_Deliverable_Data'), 'All_Munis_GDB_{0}.gdb'.format(str(time.strftime("%m%d%Y"))))
    totGDB = os.path.join(outpath, 'All_Muni_Deliverable_Data\All_Munis_GDB_{}.gdb'.format(str(time.strftime("%m%d%Y"))))
    arcpy.env.workspace = temp
    arcpy.Merge_management(roadsGDB, os.path.join(temp, 'All_Muni_Roads'))
    arcpy.Merge_management(cbsGDB, os.path.join(temp, 'All_Muni_CBs'))
    arcpy.Merge_management(citiesGDB, os.path.join(temp, 'All_Muni_City_Polys'))

    # Finally transfer to new gdb, delete temp, and delete output dirs
    for finalFC in finalFCs:
        finalpath = os.path.join(temp, finalFC)
        arcpy.FeatureClassToGeodatabase_conversion(finalpath, totGDB)
    for fc in os.listdir(temp):
        os.remove(os.path.join(temp, fc))

    # for folder in os.listdir(outpath):
    #     if 'Data' in folder:
    #         pass
    #     else:
    #         os.chdir(root)
    #         shutil.rmtree(os.path.join(outpath, folder))



def main():
    clear_gdb()
    concat_xlsx()
    merge_gdbs()


if __name__ == '__main__':
    main()
