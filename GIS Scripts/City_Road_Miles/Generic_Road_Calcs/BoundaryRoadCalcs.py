import arcpy
import os
import getpass
'''
Script - CityRoadCalcs
Created by - Colton Luttrell
Date Started - 5/20/2021

Description:
    Generalizes road calculations for boundary data pull with a class. Relies on pathing to feature classes within a scratch.gdb on the user's local machine. 

Version History:
    v1 - 5/20/2021 --> CL
'''
class roadCalcs():
    def __init__(self, roadsPath, cbsPath, polyPath, outpath, scratch=os.path.join(r'C:\Users', getpass.getuser(), 'Documents', 'ArcGIS','scratch.gdb')):
        '''
        :param roadsPath: path to all roads in scratch.gdb
        :param cbsPath: path to all roads census block in scratch.gdb
        :param polyPath: path to boundary in scratch.gdb
        :param scratch: path to scratch.gdb
        '''
        self.roadsPath = roadsPath
        self.cbsPath = cbsPath
        self.polyPath = polyPath
        self.scratch = scratch
        self.outpath = outpath
    def CleanRoads(self):
        '''
        Removes road features that are not specifically road types from census data (sidewalks, private roads, etc.)
        After removal of features, integrates highway features to remove any 'parallel' roads.
        :return: None.
        '''
        # First remove unecessary roads, then integrate S1100 and S1200 roads
        print('Removing unwanted road/path features and integrating parallel roads...')
        arcpy.env.workspace = self.scratch

        S1100 = arcpy.FeatureClassToFeatureClass_conversion(self.roadsPath, self.scratch, 'S1100',"MTFCC = 'S1100'")
        S1100Int = arcpy.Integrate_management(S1100, '150 FEET')
        S1200 = arcpy.FeatureClassToFeatureClass_conversion(self.roadsPath, self.scratch, 'S1200',"MTFCC = 'S1200'")
        S1200Int = arcpy.Integrate_management(S1200, '80 FEET')
        delCursor = arcpy.da.UpdateCursor(self.roadsPath, ['MTFCC'])
        exlcudedRoads = ['S1100', 'S1200', 'S1630', 'S1500', 'S1710', 'S1720', 'S1730', 'S1740', 'S1750', 'S1780', 'S1820', 'S1830']
        for row in delCursor:
            for road in exlcudedRoads:
                if row[0] == road:
                    delCursor.deleteRow()
        del row, delCursor
        # Get S1400 roads merged
        regRoadsMerged = arcpy.Dissolve_management(self.roadsPath, 'S1400_Merged', ['MTFCC', 'fip'])
        mergedRoads = arcpy.MergeDividedRoads_cartography('S1400_Merged', 'fip', '80 feet', 'regRoadsMerged')

        arcpy.Merge_management([S1100Int, S1200Int, mergedRoads], 'Roads_Merged')
    def RoadMiles(self):
        '''
            Intersect roads fc to all cities polygons, get total road miles in a city boundary, and then road miles per cb inside city.
            :return:none
        '''
        arcpy.env.workspace = self.scratch
        print('Getting road miles...')
        # roads to city polygons
        arcpy.Intersect_analysis(['Roads_Merged', self.polyPath], 'Roads_Polys_Intersect', '', '', 'LINE')
        arcpy.Dissolve_management('Roads_Polys_Intersect', 'Roads_in_Polys', ['NAME'])
        arcpy.AddGeometryAttributes_management('Roads_in_Polys', 'LENGTH_GEODESIC', 'MILES_US')

        # roads to CBs in city polygons
        arcpy.SpatialJoin_analysis(self.cbsPath, self.polyPath, 'CBs_in_poly', 'JOIN_ONE_TO_ONE',
                                   'KEEP_COMMON', '', 'HAVE_THEIR_CENTER_IN')
        arcpy.Intersect_analysis(['Roads_Merged', 'CBs_in_poly'], 'Roads_CBs_Intersect', '', '', 'LINE')
        arcpy.Dissolve_management('Roads_CBs_Intersect', 'Roads_CBs_Diss', ['BLOCKID10'])
        arcpy.AddGeometryAttributes_management('Roads_CBs_Diss', 'LENGTH_GEODESIC', 'MILES_US')

        # join the road miles fcs back to the Cities polygons, and the CBs polygons, make new fcs for deliverables
        joined_city_roads = arcpy.AddJoin_management(self.polyPath, 'NAME', 'Roads_in_Polys', 'NAME', 'KEEP_ALL')
        joined_cb_roads = arcpy.AddJoin_management('CBs_in_poly', 'BLOCKID10', 'Roads_CBs_Diss', 'BLOCKID10',
                                                   'KEEP_ALL')
        arcpy.FeatureClassToFeatureClass_conversion(joined_city_roads, self.scratch, 'Polygon_RoadMiles')
        arcpy.FeatureClassToFeatureClass_conversion(joined_cb_roads, self.scratch, 'CBs_in_Poly_RoadMiles')
        # removing unessecary fields
        arcpy.DeleteField_management('Polygon_RoadMiles', ['city', 'OBJECTID', 'Shape_Length_1', 'NAME_1'])
        arcpy.DeleteField_management('CBs_in_Poly_RoadMiles',
                                     ['Join_Count', 'OBJECTID', 'Shape_Length_1', 'Shape_Length_2', 'Shape_Length_12', 'Shape_Area_1',
                                      'BLOCKID10_1', 'STATEFP', 'PLACEFP', 'PLACENS',
                                      'GEOID', 'NAME', 'NAMELSAD', 'LSAD', 'CLASSFP', 'PCICBSA', 'PCINECTA', 'MTFCC',
                                      'FUNCSTAT', 'ALAND', 'AWATER', 'INTPTLAT', 'INTPTLON'])
        # Transfer to deliverable gdb, go to state dir specifically
        deliverableFCs = ['Polygon_RoadMiles', 'CBs_in_Poly_RoadMiles', 'Roads_Polys_Intersect']
        newGDB = arcpy.CreateFileGDB_management(self.outpath, 'PolygonRoadMiles.gdb')
        arcpy.FeatureClassToGeodatabase_conversion(deliverableFCs, newGDB)
    def create_excel(self):
        '''
        Creates xlsx sheet from fcs
        :return: xlsx to state dir
        '''
        print('Creating Excels...')
        deliverableFCs = ['Polygon_RoadMiles', 'CBs_in_Poly_RoadMiles']
        for fc in deliverableFCs:
            arcpy.TableToExcel_conversion(fc, self.outpath + '\\' + fc + '.xlsx')
