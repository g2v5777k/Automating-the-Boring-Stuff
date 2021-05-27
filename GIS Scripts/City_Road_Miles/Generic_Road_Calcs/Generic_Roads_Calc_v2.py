from config_generic import *

'''
Script - Generic_Roads_Calc
Created by - Colton Luttrell
Date Started - 4/20/2021

I/O:
    Input: Takes in a polygon you want road miles for, a state all roads file, and a state's census blocks.
    Output: .gdb containing the polygon attributed with road miles, CBs with road miles / CB within the polygon. Two .xlsx files with this respective data.

Version History:
    v1 - 4/20/2021 --> CL
    v2 - 5/20/2021 - Rewrote all functions here to be classes in separate python files at an attempt for 
    resuability --> CL
'''
def main():
    print('Starting Calculations...')
    print('------------------------')
    prep = scratchPrep()
    prep.clear_gdb()
    prep.transfer(dictList=[roads, blocks, poly])
    roadClass = roadCalcs(roads['_scratch'], blocks['_scratch'], poly['_scratch'], outpath=outpath)
    roadClass.CleanRoads()
    roadClass.RoadMiles()
    roadClass.create_excel()
    print('------------------------')
    print('Finished!')
if __name__ == '__main__':
    main()