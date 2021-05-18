from config import *
'''
Script - RDOF Field Tracking
Created by - Colton Luttrell
Date Created - 4/28/2021

Requirements:
    1) RDOFCollector.sde connection file, from 42 server.
    
I/O:
    Input: RDOF Collector .sde
    Output: .xlsx with number of field points created per day/lifetime per fielder username, total HLD fiber miles, and points within a CBG/total points.
    
Version History:
    v1 - 4/28/2021
    v2 - Added cbgPoints() to calculate number of points within CBGs, and total CBG points. Added calculation of lifetime num of points per fielder.
'''
def downloadSDE():
    '''
    Creates a local .gdb of collector db to work from.
    :return: None
    '''
    print('Transferring sde data to local GDB...')
    arcpy.env.workspace = gdb
    arcpy.CreateFileGDB_management(inputPath, genName)
    for path, datasets, fcs in arcpy.da.Walk(os.path.join(sde_con)):
        for data, fc in zip(datasets, fcs):
            try:
                arcpy.Copy_management(os.path.join(sde_con, data), os.path.join(gdb, data), 'FeatureDataset')
            except Exception as e:
                exe_type, exe_obj, exe_tb = sys.exc_info()
                print(exe_type, exe_obj, exe_tb)
def getFielderPoints():
    '''
    Gather point fcs and list out count by created_user, export to csv.
    :return: None
    '''
    try:
        print('Gathering fielder points...')
        arcpy.env.workspace = gdb
        exp = "created_date > timestamp '{}' And created_date < timestamp '{}'".format(start, end)
        for fc, val in fcsDict.items():
            arcpy.FeatureClassToFeatureClass_conversion(os.path.join(gdb, fc), gdb, val[0], exp)
            arcpy.Dissolve_management(val[0],  val[1], ['created_user'], [['created_user', 'COUNT']])
            arcpy.Dissolve_management(fc, val[4], ['created_user'], [['created_user', 'COUNT']])
            cur = arcpy.da.SearchCursor(val[1], ['created_user', 'COUNT_created_user'])
            curTot = arcpy.da.SearchCursor(val[4], ['created_user', 'COUNT_created_user'])
            for row in cur:
                userList.append([row[0], fc, row[1]])
            for rowTot in curTot:
                userListTot.append([rowTot[0], fc, rowTot[1]])
        del cur, row, rowTot, curTot
        df = pd.DataFrame(userList, columns=['User', 'Point_FC', 'Count_Created_Features'])
        dfTot = pd.DataFrame(userListTot, columns=['User', 'Point_FC', 'Lifetime_Points_Created'])
        piv = pd.pivot_table(df, index=['User', 'Point_FC'])
        pivTot = pd.pivot_table(dfTot, index=['User', 'Point_FC'])
        concat = pd.merge(piv, pivTot, how='left', left_on=['User', 'Point_FC'], right_on=['User', 'Point_FC'])
        concat.to_excel(writer, sheet_name='FielderPoints')
        writer.save()
    except Exception as e:
        print('An error has been encountered when extracting fielder points, please check that there is data for todays run.')
        print(f'Error: {e}')
def getCBGPoints():
    '''
    Gathers point counts within CBGs.
    :return: None
    '''
    print('Gathering points within CBGs...')
    arcpy.env.workspace = gdb
    for fc, val in fcsDict.items():
        arcpy.SpatialJoin_analysis(fc, 'Won_CBGs', val[2], 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'WITHIN')
        countDict[fc] = [str(arcpy.GetCount_management(val[2]))]
        countDict[fc].append(str(arcpy.GetCount_management(fc)))
    df = pd.DataFrame.from_dict(countDict, orient='index', columns=['Points_in_CBGs', 'Total_Points'])
    df.to_excel(writer, sheet_name='CBGs_Points')
    writer.save()
def getHLDMiles():
    '''
    Calculates fiber HLD Miles, exports to csv.
    :return: None
    '''
    print('Calculating total HLD miles...')
    arcpy.env.workspace = gdb
    fiber = os.path.join(gdb, 'HLD_Fiber')
    arcpy.Dissolve_management(fiber, fiber + '_dissolved')
    arcpy.AddGeometryAttributes_management(fiber + '_dissolved', 'LENGTH_GEODESIC', 'MILES_US')
    cur = arcpy.da.SearchCursor(fiber + '_dissolved', ['LENGTH_GEO'])
    for row in cur:
        miles = [row[0]]
    del row, cur
    milesDF = pd.DataFrame(miles, columns=['Total_HLD_Fiber_Miles'])
    milesDF.to_excel(writer, sheet_name='HLD_Miles', index=False)
    writer.save()
    writer.close()
def clearDirs():
    '''
    Simply delete files in inputGDB, and output folder for next days run.
    :return: None
    '''
    try:
        gdbPath = os.path.join(inputPath, 'RDOF_FieldTracker.gdb')
        arcpy.Delete_management(gdbPath)
    except PermissionError:
        pass
    for xl in os.listdir(output):
        path = os.path.join(output, xl)
        os.remove(path)
def main():
    downloadSDE()
    getFielderPoints()
    getCBGPoints()
    getHLDMiles()
    for file in os.listdir(r'C:\Scripts\RDOF_FieldTracking\output'):
        filepath = os.path.join(output, file)
        authenticate(filepath, r'1jnEZPl0KJA0siezz-056SwK0Ltx9_Lq3')
    clearDirs()
if __name__ == '__main__':
    main()
