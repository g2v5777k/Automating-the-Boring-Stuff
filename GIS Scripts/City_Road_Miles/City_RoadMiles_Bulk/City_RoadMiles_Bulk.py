from config_bulk import *
from concat_data import *

'''
Script - City Road Miles Bulk
Created by - Colton Luttrell
Date Started - 4/1/2021


Requirements:
    1) xlsx of cities to get miles/2010 hhp counts for. Must have a city and state column. Follow format of template provided in script input directory.

I/O:
    Input: xlsx w/ city names and states in separate columns outlined in template.
    Output: .xlsx for CBs and city boundary road miles/hhps in state specific directory

Notes:
    1) Takes in a city/state csv of interest cities. States parsed will create a new csv by state that goes into temp folder
    2) Will break up into different directories in output folder per state. For example if you had 4 cities in TX, and 3 in MO you'd have two directories (TX, MO) where those cities data will live.
    3) Returns data into these directories.
    4) Once all data has been calculated and stored in their respective directories, concat_data.py runs to merge all gdbs and xlsxs into single files.
       Removes all state specific directories to save user space.

    Script broken up into two sections:
    1) First parses the csv to get states where cities reside. Download all necessary census files by state to temp dir.
    2) Run calculate miles spatial analysis.
    - Loop through (2) until all states have been covered.


Version History:
    v1 - 4/1/2021 CL
    v2 - 4/15/2021 -> Added a CleanRoads() function that pares down the roads used to caluculate mileages, and integrates parallels.
'''

############################################# Start Data Prep Functions #############################################

def make_csv(fips):
    '''
    Create parsed csv for cities in specific state
    :param fips: FIPs code for state
    :return:
    reduced csv for cities in specific state
    '''

    stateSpecificDF = merged[merged['fip'] == fips]
    stateSpecificDF.to_csv(temp + '\\' + 'cities.csv', index=False)

def download_merge_roads(fips):
    '''
    Downloads all road files for a state, merges them, move final shp to temp
    :return:
    Merged road file to temp dir
    '''
    allRoads = []
    print('Downloading road shapefiles....')
    # get all roads for state from census site using bs4 module
    os.chdir(roadsDir)
    roadSnip = 'tl_2020_' + fips
    url = 'https://www2.census.gov/geo/tiger/TIGER2020/ROADS/'
    r = requests.get(url)
    soup = BeautifulSoup(r.text, 'html.parser')
    for link in soup.find_all('a'):
        if roadSnip in link.get_text():
            name = link.get('href')
            fileurl = url + name
            newR = requests.get(fileurl)
            open(name, 'wb').write(newR.content)

    # unzip all the files
    print('Merging road files, this may take a while...')
    for shp in os.listdir(roadsDir):
        with zipfile.ZipFile(shp, 'r') as zip_ref:
            zip_ref.extractall(roadsDir)
    for shp in os.listdir(roadsDir):
        if shp.endswith('.shp'):
            allRoads.append(shp)

    arcpy.env.workspace = temp
    arcpy.Merge_management(allRoads, 'All_' + fips + '_Roads.shp')

    # delete files from roads dir once everything has been downloaded and merged
    for shp in os.listdir(roadsDir):
        os.remove(shp)

def download_census_blks_places(fips):
    '''
    Function uses requests module to download .zip from entered url and filename to output folder
    :param fips: FIPs code of state
    :return:
    files to temp dir
    '''

    dlDict = {'tabblock2010_' + fips + '_pophu.zip':'https://www2.census.gov/geo/tiger/TIGER2010BLKPOPHU/tabblock2010_'+ fips + '_pophu.zip',
              'tl_2020_' + fips + '_place.zip':'https://www2.census.gov/geo/tiger/TIGER2020/PLACE/tl_2020_'+ fips + '_place.zip'}

    print('Downloading CBs and places shapefiles to temp folder...')
    for filename, url in dlDict.items():
        r = requests.get(url, allow_redirects=True)
        os.chdir(os.path.join(temp))
        open(filename, 'wb').write(r.content)
        # unzip to temp dir after downloaded
        with zipfile.ZipFile(filename, 'r') as zip_ref:
            zip_ref.extractall(temp)
    # Delete XML files from temp
    for file in os.listdir(temp):
        if file.endswith('.xml'):
            os.remove(os.path.join(temp,file))
    print('Files downloaded...\n')

############################################# Start Spatial Analysis Functions #############################################
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

def transfer():
    '''
    Transfer shapefiles in our directories to scratch.gdb for ease of use.
    :return:
    None.
    '''
    print('Transferring input files...\n')
    shps = [roads['_path'], blocks['_path'], cities['_path']]
    for shp in shps:
        arcpy.FeatureClassToGeodatabase_conversion(shp, scratch)
    arcpy.TableToTable_conversion(interestCities['_path'], scratch, 'Interest_Cities')
    print('\n')
def get_cities():
    '''
    Join csv file with cities of interest to places.shp and create new fc of just interest cities. (Probably don't need a full function here, but keeping separate for the time being)
    Also, spatial join to get census blocks within the city boundaries.
    :return:
    None.
    '''
    print('Getting city polygons and CBs in city polygons...')
    joined_city = arcpy.AddJoin_management(cities['_scratch'].strip('.shp'), 'NAME', 'Interest_Cities', 'city', 'KEEP_COMMON')
    arcpy.FeatureClassToFeatureClass_conversion(joined_city, scratch, 'Cities_Polys')
    arcpy.SpatialJoin_analysis(blocks['_scratch'].strip('.shp'), 'Cities_Polys', 'CBs_in_Cities', 'JOIN_ONE_TO_ONE','KEEP_COMMON', '', 'HAVE_THEIR_CENTER_IN')

def CleanRoads():
    # First remove unecessary roads, then integrate S1100 and S1200 roads
    print('Removing unwanted road/path features and integrating parallel roads...')
    S1100 = arcpy.FeatureClassToFeatureClass_conversion(roads['_scratch'].split('.shp')[0], scratch, 'S1100', "MTFCC = 'S1100'")
    S1100Int = arcpy.Integrate_management(S1100, '150 FEET')
    S1200 = arcpy.FeatureClassToFeatureClass_conversion(roads['_scratch'].split('.shp')[0], scratch, 'S1200', "MTFCC = 'S1200'")
    S1200Int = arcpy.Integrate_management(S1200, '80 FEET')

    delCursor = arcpy.da.UpdateCursor(roads['_scratch'].split('.shp')[0], ['MTFCC'])
    for row in delCursor:
        for road in exlcudedRoads:
            if row[0] == road:
                delCursor.deleteRow()
    del row, delCursor
    arcpy.Merge_management([S1100Int, S1200Int, roads['_scratch'].split('.shp')[0]], 'Roads_Merged')



def RoadMiles(fips):
    '''
    Intersect roads fc to all cities polygons, get total road miles in a city boundary, and then road miles per cb inside city.
    :return:
    none
    '''

    print('Getting road miles...')
    # roads to city polygons
    arcpy.Intersect_analysis(['Roads_Merged', 'Cities_Polys'], 'Roads_City_Polys_Intersect','','','LINE')
    arcpy.Dissolve_management('Roads_City_Polys_Intersect', 'Roads_in_City_Polys', ['NAME'])
    arcpy.AddGeometryAttributes_management('Roads_in_City_Polys', 'LENGTH_GEODESIC', 'MILES_US')

    # roads to CBs in city polygons
    arcpy.Intersect_analysis(['Roads_Merged', 'CBs_in_Cities'], 'Roads_CBs_Intersect', '', '','LINE')
    arcpy.Dissolve_management('Roads_CBs_Intersect', 'Roads_CBs_Diss', ['BLOCKID10'])
    arcpy.AddGeometryAttributes_management('Roads_CBs_Diss', 'LENGTH_GEODESIC', 'MILES_US')

    # join the road miles fcs back to the Cities polygons, and the CBs polygons, make new fcs for deliverables
    joined_city_roads = arcpy.AddJoin_management('Cities_Polys','NAME','Roads_in_City_Polys', 'NAME', 'KEEP_ALL')
    joined_cb_roads = arcpy.AddJoin_management('CBs_in_Cities','BLOCKID10','Roads_CBs_Diss', 'BLOCKID10', 'KEEP_ALL')
    arcpy.FeatureClassToFeatureClass_conversion(joined_city_roads, scratch, 'Cities_Polygons_RoadMiles')
    arcpy.FeatureClassToFeatureClass_conversion(joined_cb_roads, scratch, 'CBs_in_Cities_RoadMiles')
    # removing unessecary fields
    arcpy.DeleteField_management('Cities_Polygons_RoadMiles', ['city', 'OBJECTID','Shape_Length_1', 'NAME_1'])
    arcpy.DeleteField_management('CBs_in_Cities_RoadMiles', ['Join_Count','OBJECTID', 'Shape_Length_1', 'Shape_Length_2', 'Shape_Length_12'
                                                             'Shape_Area_1', 'OBJECTID_1', 'BLOCKID10_1','STATEFP', 'PLACEFP', 'PLACENS',
                                                             'GEOID', 'NAME', 'NAMELSAD', 'LSAD', 'CLASSFP', 'PCICBSA', 'PCINECTA', 'MTFCC',
                                                             'FUNCSTAT','ALAND', 'AWATER', 'INTPTLAT', 'INTPTLON'])

    # Transfer to deliverable gdb, go to state dir specifically
    for state, fip in statesDict.items():
        if fip == fips:
            currentDir = os.path.join(outpath, state)
    deliverableFCs = ['Cities_Polygons_RoadMiles', 'CBs_in_Cities_RoadMiles', 'Roads_City_Polys_Intersect']
    newGDB = arcpy.CreateFileGDB_management(currentDir, 'city_results.gdb')
    arcpy.FeatureClassToGeodatabase_conversion(deliverableFCs, newGDB)

def create_excel(fips):
    '''
    Creates xlsx sheet from fcs
    :return:
    xlsx to state dir
    '''
    print('Creating Excels...')
    for state, fip in statesDict.items():
        if fip == fips:
            currentDir = os.path.join(outpath, state)
    deliverableFCs = ['Cities_Polygons_RoadMiles', 'CBs_in_Cities_RoadMiles']
    for fc in deliverableFCs:
        arcpy.TableToExcel_conversion(fc, currentDir + '\\'+ fc + '.xlsx')

############################################# Main Functions #############################################
def data_prep(fips):
    '''
    Collection of data prep scripts
    :param fips: state fips
    :return:
    all necessary data to temp folder for calculate road miles functions
    '''
    make_csv(fips)
    download_merge_roads(fips)
    download_census_blks_places(fips)

def calculate_road_miles(fips):
    '''
    Collection of scripts that will perform the spatial analysis portion of script.
    :return:
    data to output
    '''
    # set pathing, globals here as  opposed to defining in config since they'd be defined before files are made.
    global roads, blocks, cities, interestCities
    for file in os.listdir(temp):
        if 'Roads.shp' in file:
            roadShp = file
        elif 'pophu.shp' in file:
            popBlks = file
        elif 'place.shp' in file:
            places = file
        elif '.csv' in file:
            cityList = file
    roads = {
        '_path': os.path.join(temp, roadShp),
        '_scratch': os.path.join(scratch, roadShp)
    }
    blocks = {
        '_path': os.path.join(temp, popBlks),
        '_scratch': os.path.join(scratch, popBlks)
    }
    cities = {
        '_path': os.path.join(temp, places),
        '_scratch': os.path.join(scratch, places)

    }
    interestCities = {
        '_path': os.path.join(temp, cityList),
        '_scratch': os.path.join(scratch, cityList)
    }
    clear_gdb()
    transfer()
    get_cities()
    CleanRoads()
    RoadMiles(fips)
    create_excel(fips)

    # once this function has complete, delete out temp dir for next run
    for result in os.listdir(temp):
            os.remove(os.path.join(temp, result))

def data_clean():
    '''
    Imports from concat_data.py to clean up directories, and produce on dir with all results.
    :return:
    One gdb, and two xlsx files of all states results
    '''
    clear_gdb()
    concat_xlsx()
    merge_gdbs()

def main():
    for state, fips in statesDict.items():
        print('-----------------------------------------')
        print(state, type(state))
        print('Starting ' + state + ' cities...')
        data_prep(fips)
        arcpy.env.workspace = scratch
        calculate_road_miles(fips)
        print(state + ' finished\n')
        print('-----------------------------------------')
    data_clean()
    print('Done! Please check output folder.')
if __name__ == '__main__':
    main()
