from config import *
'''
Script - RDOF BOM
Created by - Colton Luttrell
Date Started - 7/29/2021
Requirements:
    Current RDOF_Design.gdb
    client_secrets.json file
    credentials.json file
    token.pickle file
I/O:
    Input: Takes in an RDOF OLT/LCP name.
    Output: .xlsx of filled out BOM template (in input dir)
Version History:

Notes:
'''
############# Data Download and Transfer to root/scratch.gdb ################
def downloadGDB():
    '''
    Downloads daily RDOF gdb from Google Drive location. The file token.pickle stores the user's access and
    refresh tokens, and is completed automatically when the auth flow completes the first time (if not already present).
    :return: None
    '''
    # cred auth check
    creds = None
    if os.path.exists('token.pickle'):
        with open(os.path.join(root,'token.pickle'), 'rb') as token:
            creds = pickle.load(token, encoding='latin1')
    # If there are no (valid) credentials available, will redirect you to page where manual auth is required.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(os.path.join(root,'credentials.json'), SCOPES)
            creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
        with open(os.path.join(root,'token.pickle'), 'wb') as token:
            pickle.dump(creds, token)

    # build service instance, look for today's RDOF Design gdb file id.
    service = build('drive', 'v3', credentials=creds)
    pageToken = None
    # loop through drive to find today's gdb
    while True:
        response = service.files().list(driveId='0ALWJmYKL39C_Uk9PVA',
                                        includeItemsFromAllDrives=True,
                                        corpora='drive',
                                        fields='files(id, name)',
                                        supportsAllDrives=True).execute()
        for file in response.get('files', []):
            if gdb in file.get('name'):
                fileId = file.get('id')
                print("Today's RDOF GDB found...")
        pageToken = response.get('nextPageToken', None)
        if pageToken is None:
            break
        return fileId

    # once file is found, download to /input
    request = service.files().get_media(fileId=fileId)
    fh = io.FileIO(os.path.join(inputs, gdb),'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print('Downloading progress -> %d%% ' % int(status.progress() * 100))
    print('Download Complete \n')
def unzip():
    '''
    Unzip daily RDOF gdb in /input
    :return: None
    '''
    with zipfile.ZipFile(os.path.join(inputs, gdb), 'r') as zipref:
        zipref.extractall(inputs)
def clear_gdb():
    '''
    Clear contents of scratch.gdb.
    :return: None
    '''
    arcpy.env.workspace = scratch
    print('Clearing Workspace...')
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
    Transfer layers needed from RDOF_Design.gdb to scratch.gdb. If span layer is present in RDOF_Design, remove it so updated span can be added.
    :return: None
    '''
    print('Transferring files from gdb to scratch...\n')
    arcpy.env.workspace = datasetPath
    for fc in necessaryFCs:
        arcpy.FeatureClassToGeodatabase_conversion(join(datasetPath, fc), scratch)
    print('\n')
############# Start BOM calcs and exports #############
def getLCPFeatures(lcp, lcpNameFixed):
    '''
    Queries out the LCP of interest, and all other total features needed for the BOM. These layers will be used for the rest of the functions in the script.
    :param lcp: Raw lcp name used for query expressions.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Collecting LCP features...')
    global totalConduit, totalFiber, lcpSplices, lcpEquipment, slackLoops, lcpStructures, lcpRisers, lcpCabs
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True
    # Query out LCP boundary
    arcpy.FeatureClassToFeatureClass_conversion('Proposed_OLT_LCP_Boundaries', scratch, f'{lcpNameFixed}_Boundary', f"cab_id = '{lcp}'")

    # Served Addresses
    arcpy.SpatialJoin_analysis('ServedAddress', f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_Adds', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'WITHIN')

    # Drop Fiber
    arcpy.SpatialJoin_analysis('DropFiber', f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_Drops', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'HAVE_THEIR_CENTER_IN')

    # Conduit
    plannedConduit = arcpy.SelectLayerByAttribute_management('Conduit', 'NEW_SELECTION', "inventory_status_code = 'P' OR inventory_status_code = 'Planned'")
    conduit = arcpy.Intersect_analysis([plannedConduit, f'{lcpNameFixed}_Boundary'], f'{lcpNameFixed}_Total_Conduit', '', '', 'LINE')
    totalConduit = arcpy.AddGeometryAttributes_management(conduit, 'LENGTH_GEODESIC', 'FEET_US')

    # Fiber
    plannedFiber = arcpy.SelectLayerByAttribute_management('FiberCable', 'NEW_SELECTION', "inventory_status_code = 'P' OR inventory_status_code = 'Planned'")
    fiber = arcpy.Intersect_analysis([plannedFiber, f'{lcpNameFixed}_Boundary'], f'{lcpNameFixed}_Total_Fiber', '', '', 'LINE')
    totalFiber = arcpy.AddGeometryAttributes_management(fiber, 'LENGTH_GEODESIC', 'FEET_US')

    # Splices
    plannedSplices = arcpy.SelectLayerByAttribute_management('SpliceClosure', 'NEW_SELECTION', "inventory_status_code = 'P' OR inventory_status_code = 'Planned'")
    lcpSplices = arcpy.SpatialJoin_analysis(plannedSplices, f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_SpliceClosures', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON','', 'WITHIN')

    # Structures
    plannedFiberEquip = arcpy.SelectLayerByAttribute_management('FiberEquipment', 'NEW_SELECTION', "inventory_status_code = 'P'")
    lcpEquipment = arcpy.SpatialJoin_analysis(plannedFiberEquip, f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_FiberEquipment', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON','', 'WITHIN')
    plannedStructure = arcpy.SelectLayerByAttribute_management('Structure', 'NEW_SELECTION', "inventorystatuscode = 'P'")
    lcpStructures = arcpy.SpatialJoin_analysis(plannedStructure, f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_Structure', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'WITHIN')

    # Slackloops
    plannedSlackLoops = arcpy.SelectLayerByAttribute_management('SlackLoop', 'NEW_SELECTION', "inventory_status_code = 'P' OR inventory_status_code = 'Planned'")
    slackLoops = arcpy.SpatialJoin_analysis(plannedSlackLoops, f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_SlackLoops', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'HAVE_THEIR_CENTER_IN')

    # Risers
    plannedRisers = arcpy.SelectLayerByAttribute_management('Riser', 'NEW_SELECTION', "inventory_status_code = 'P' OR inventory_status_code = 'Planned'")
    lcpRisers = arcpy.SpatialJoin_analysis(plannedRisers, f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_Risers', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'WITHIN')

    # Cabinets
    lcpCabs = arcpy.SpatialJoin_analysis('Proposed_Cabinets', f'{lcpNameFixed}_Boundary', f'{lcpNameFixed}_Total_ProposedCabs', 'JOIN_ONE_TO_ONE', 'KEEP_COMMON', '','WITHIN')
def addresses(lcpNameFixed):
    '''
    Extract 'Served Addresses' within an LCP's CBG, not in CBG or overbuild poly, and inside an overbuild poly.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Fetching addresses...')
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True

    # get total 'designed to' served addresses in a LCP boundary
    boundaryAdds = arcpy.Intersect_analysis([f'{lcpNameFixed}_Total_Adds', 'DropFiber'], f'{lcpNameFixed}_DesignedToAdds','','','POINT')

    # get cbg served address points within lcp boundary (for cell D4)
    arcpy.SpatialJoin_analysis(f'{lcpNameFixed}_Total_Adds', 'RDOF_CBG', f'{lcpNameFixed}_CBG_Adds', 'JOIN_ONE_TO_ONE','KEEP_COMMON', '', 'WITHIN')
    cbgAdds = arcpy.Intersect_analysis([f'{lcpNameFixed}_CBG_Adds', 'DropFiber'],f'{lcpNameFixed}_DesignedToAdds_inCBGs', '', '', 'POINT')
    cbgAddsCount = arcpy.GetCount_management(cbgAdds)
    ws['D4'] = int(cbgAddsCount[0])

    # get LCP served addresses that are within an overbuild polygon (for cell D6)
    arcpy.SpatialJoin_analysis(f'{lcpNameFixed}_Total_Adds', 'OVERBUILD_POLY', f'{lcpNameFixed}_Overbuild_Adds','JOIN_ONE_TO_ONE', 'KEEP_COMMON', '', 'WITHIN')
    overbuildAdds = arcpy.Intersect_analysis([f'{lcpNameFixed}_Overbuild_Adds', 'DropFiber'], f'{lcpNameFixed}_OverbuildDesignedToAdds','','','POINT')
    overbuildAddsCount = arcpy.GetCount_management(overbuildAdds)
    ws['D6'] = int(overbuildAddsCount[0])

    # get served addresses that are not in CBGs and not in an overbuild polygon (for cell D5)
    arcpy.Erase_analysis(boundaryAdds, overbuildAdds, f'{lcpNameFixed}_ErasedOverbuild')
    erasedTot = arcpy.Erase_analysis(f'{lcpNameFixed}_ErasedOverbuild', cbgAdds, f'{lcpNameFixed}_TotalErased_D5')
    erasedTotCount = arcpy.GetCount_management(erasedTot)
    ws['D5'] = int(erasedTotCount[0])

    # Get total adds passed (designed to), and olt commissioning calculation
    totalAddsPassed = int(cbgAddsCount[0]) + int(overbuildAddsCount[0]) + int(erasedTotCount[0])
    ws['D7'] = totalAddsPassed
    # OLT Commissioning (for cell E153)
    ws['E153'] = math.ceil(totalAddsPassed / 496)
def dropFiber(lcpNameFixed):
    '''
    Calculates # of long drops (>600') and average drop length within an LCP for cells D25 and D28 respectively.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Fetching drop fiber...')
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True

    # Get # of drops over 600' in length (for cell D25)
    totalDrops = arcpy.AddGeometryAttributes_management(f'{lcpNameFixed}_Total_Drops', 'LENGTH_GEODESIC', 'FEET_US')
    dropsOver600 = arcpy.SelectLayerByAttribute_management(totalDrops, 'NEW_SELECTION', "LENGTH_GEO > 600")
    ws['D25'] = int(arcpy.GetCount_management(dropsOver600)[0])

    # Get average drop length (for cell D28)
    cur = arcpy.da.SearchCursor(totalDrops, ['LENGTH_GEO'])
    allDropLengths = [row[0] for row in cur]
    avgDropLength = math.ceil(int(sum(allDropLengths) / len(allDropLengths)))
    ws['D28'] = avgDropLength
def conduit(lcpNameFixed):
    '''
    Calculations for various stats for conduit. Directional Bore, Plow, UG special crossings, adder in same trench, 1.25" & 2" lengths.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Fetching conduit...')
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True

    # Get all conduit/ug fiber within LCP boundary of interest
    totalUGFiber = arcpy.SelectLayerByAttribute_management(totalFiber, 'NEW_SELECTION', "placementtype = 'UG'")

    # Get conduit that is only for drops (for cell E65)
    onlyDrops = arcpy.SelectLayerByAttribute_management(totalConduit, 'NEW_SELECTION', "dropsonly = 'Y'")
    ws['E65'] = sumField(onlyDrops, 'LENGTH_GEO')

    # Missile Bore	(for cell E66). Pending attribution in database.

    # Directional bore, up to 2" conduit. Meaning total footage of conduit (for cell E67).
    ws['E67'] = sumField(totalConduit, 'LENGTH_GEO')

    # Plow - direct bury armored cable or up to 2" conduit = total UG fiber not in conduit (for cell E68)
    initialPlow = arcpy.Erase_analysis(totalUGFiber, totalConduit, f'{lcpNameFixed}_Plow')
    ws['E68'] = sumField(initialPlow, 'LENGTH_GEO')

    # Get all 1.25" conduit x 1.07 (for cell E118)
    oneInch = arcpy.SelectLayerByAttribute_management(totalConduit, 'NEW_SELECTION', "duct_diameter = 1")
    ws['E118'] = math.ceil(sumField(oneInch, 'LENGTH_GEO') * 1.07)

    # Get all 2" conduit x 1.07 (for cell E119)
    twoInch = arcpy.SelectLayerByAttribute_management(totalConduit, 'NEW_SELECTION', "duct_diameter = 2")
    ws['E119'] = math.ceil(sumField(twoInch, 'LENGTH_GEO') * 1.07)

    # Get conduit adder in same trench , i.e. parallel conduit (for cell E72). Pending attribution in database.

def spliceClosures(lcpNameFixed):
    '''
    Various calculations for splice closures.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Fetching splices...')
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True

    # Get all splice closures and fiber equipment within LCP boundary of interest (for cell E90)
    ws['E90'] = int(arcpy.GetCount_management(lcpSplices)[0])

    # Single Fusion Fiber Splicing - Total cable size for all RE, MCA, and NAPMCA type splices + # of HHP's from all NAP splices. (for cell E91)
    noNaps = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "spliceenclosuretype <> 'NAP'")
    noNapsSum = sumField(noNaps, 'cable_size')
    onlyNaps = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "spliceenclosuretype = 'NAP'")
    NapsSum = sumField(onlyNaps, 'hhp_count')
    ws['E91'] = noNapsSum + NapsSum

    # Channell F1 Intercept Enclosure (Green Hornet G6, (4) 24ct splice trays) = total G6 sized splice enclosure that's type RE or MCA or NAPMCA (for cell E124)
    g6Splices = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "splicesize='G6' AND (spliceenclosuretype = 'NAPMCA' OR spliceenclosuretype = 'RE' OR spliceenclosuretype = 'MCA')")
    ws['E124'] = int(arcpy.GetCount_management(g6Splices)[0])

    # Channell Primary Splitter Enclosure (Green Hornet G5, (3) 24ct splice trays, (1) 1x32 splitter bare leads) (for cell E125)
    g5Splices = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "splicesize = 'G5N'")
    spliceEquipment = arcpy.SelectLayerByAttribute_management(lcpEquipment, 'NEW_SELECTION', "equipment_type = 32") # gets only 1x32 splitters
    intersectedEquipment = arcpy.Intersect_analysis([spliceEquipment, g5Splices], f'{lcpNameFixed}_1x32_G5N_Splitters','','','POINT')
    ws['E125'] = int(arcpy.GetCount_management(intersectedEquipment)[0])

    # Channell Reel End Enclosure (Green Hornet G5, (4) 24ct splice trays) - total G5 sized splice enclosures with type RE, MCA, or NAP MCA (for cell E126)
    g5SplicesPared = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "splicesize = 'G5N' AND (spliceenclosuretype = 'NAPMCA' OR spliceenclosuretype = 'RE' OR spliceenclosuretype = 'MCA')")
    ws['E126'] = int(arcpy.GetCount_management(g5SplicesPared)[0])

    # Channell Drop Terminal Enclosure (Green Hornet G5 (1) 24ct splice tray) - total G5 enclosures with type NAP (for cell E127)
    g5SplicesNAP= arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "splicesize = 'G5N' AND spliceenclosuretype = 'NAP'")
    ws['E127'] = int(arcpy.GetCount_management(g5SplicesNAP)[0])

    # Aerial Hanging Bracket for G5/G6 Enclosure - total splice enclosures that have environment aerial (for cell E128)
    aerSplices = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "placementtype = 'AER'")
    ws['E128'] = int(arcpy.GetCount_management(aerSplices)[0])

    # Pole Mount Bracket for FOSC 450 - total qty of FOSC450 enclosures with enviroment aerial (for cell E131)
    ws['E131'] = 0 # don't believe this is attributed correctly yet
def structures():
    '''
    Various calculations for structures.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Fetching structures...')
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True

    # Get all risers within LCP boundary of interest (for cell E53)
    ws['E53'] = int(arcpy.GetCount_management(lcpRisers)[0])

    # Install Cable Marker = to qty of all vaults + all splice closures marked as RE, MCA, and MCANAP that are underground environment (for cell E80)
    allVaults = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structure_size = '2'")
    allUGSplices = arcpy.SelectLayerByAttribute_management(lcpSplices, 'NEW_SELECTION', "placementtype = 'UG' AND (spliceenclosuretype = 'NAPMCA' OR spliceenclosuretype = 'RE' OR spliceenclosuretype = 'MCA')")
    ws['E80'] = math.ceil(int(arcpy.GetCount_management(allVaults)[0]) + int(arcpy.GetCount_management(allUGSplices)[0]))

    # Install Communications Hut = total # of huts proposed in this cabinet area (for cell E88). Pending attribution in database.

    # Install Active 24RU/17RU or GPON Cabinet - total # of cabinets proposed in this area (for cell E89)
    ws['E89'] = int(arcpy.GetCount_management(lcpCabs)[0])

    # Channell Pedestal 12x12x25" - Total of Medium peds (for cell E110)
    mediumPeds = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'MP'")
    ws['E110'] = int(arcpy.GetCount_management(mediumPeds)[0])

    # Channell Pedestal 12x12x34" - Total of Large Peds (for cell E111)
    largePeds = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'LP'")
    ws['E111'] = int(arcpy.GetCount_management(largePeds)[0])

    # Channell Pedestal 14x20x40" (for cell E112). No direction yet on this line item.

    # Get total count of small vaults in the LCP boundary of interest (for cell E113)
    smallVaults = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'SV'") # No features currently, and is not a domain value yet. Assume SV will denote small vaults.
    ws['E113'] = int(arcpy.GetCount_management(smallVaults)[0])

    # Get total count of medium vaults in the LCP boundary of interest (for cell E114)
    mediumVaults = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'MV'")
    ws['E114'] = int(arcpy.GetCount_management(mediumVaults)[0])

    # Vault 30x48x36 - 1 per 17RU cabinet (for cell E115)
    cab17RU = arcpy.SelectLayerByAttribute_management(lcpCabs, 'NEW_SELECTION', "comments = '17 RU'") # have a feeling this attribute will move to a new field in the future.
    ws['E115'] = int(arcpy.GetCount_management(cab17RU)[0])

    # Vault 36x60x36 - 1 per 24RU cabinet (for cell E116)
    cab24RU = arcpy.SelectLayerByAttribute_management(lcpCabs, 'NEW_SELECTION',"comments = '24 RU'")
    ws['E116'] = int(arcpy.GetCount_management(cab24RU)[0])

    # Flower Pot pedestrian (10x10) - total of Drop Vault or Flower Pot (for cell E117)
    ws['E117'] = 0 # attribution does not seem to be available for this yet.

    # Vertiv OLT Cabinet with Battery and Generator Port (Medium Cabinet) = total qty of 17RU sized cabinets (for cell E137)
    ws['E137'] = int(arcpy.GetCount_management(cab17RU)[0]) # keeping this the same value as E114 for now, until further direction.

    # American Production (Large Cabinet) = Total qty of 24RU sized cabinets (for cell E138)
    ws['E138'] = int(arcpy.GetCount_management(cab24RU)[0]) # keeping this the same value as E115 for now, until further direction.

    # 1x16 Optical Splitter with SC/APC pigtails attached, only at passive cabinet = These are only for passive cabinets. We haven't nailed down quite how these will look so let's leave these at 0 (for cell E142)
    ws['E142'] = 0

    # 1x32 Optical Splitter with SC/APC pigtails attached, only at passive cabinet = These are only for passive cabinets. We haven't nailed down quite how these will look so let's leave these at 0 (for cell E143)
    ws['E143'] = 0
def fiber(lcpNameFixed):
    '''
    Various calculations for fiber features within LCP boundary of interest.
    :param lcpNameFixed: LCP name that removes any illegal chars (- particularly) done in main()
    :return: None
    '''
    print('Fetching fiber...')
    global fiber12, fiber24, fiber48, fiber96, fiber144, fiber288, vaults, mediumPeds, largePeds, mediumVaults, largeVaults
    arcpy.env.workspace = scratch
    arcpy.env.overwriteOutput = True

    # Get all structure features needed to fiber calcs (Storage loops, M/L pedestals, overhead guys, etc.). All fiber for LCP was gathered on line 189 in the conduit function.
    mediumPeds = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'MP'")
    largePeds = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'LP'")
    mediumVaults = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'MV'")
    largeVaults = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'LV'")
    vaults = arcpy.SelectLayerByAttribute_management(lcpStructures, 'NEW_SELECTION', "structuretype = 'LV' OR structuretype = 'MV'")

    # Aerial placement of strand (new strand + overhead guy) = total aerial fiber footage * 1.05 + qty of down guys + total overhead guy length (for cell E42).
    aerFiber = arcpy.SelectLayerByAttribute_management(totalFiber, 'NEW_SELECTION', "placementtype = 'AE'")
    # will add the other summations here when downguys and overhead guys are added to the db.
    ws['E42'] = math.ceil(float(sumField(aerFiber, 'LENGTH_GEO') *1.05))

    # Lashing Fiber - total all aerial fiber footages * 1.05 + total aerial storage loops length (for cell E44)
    aerLoops = arcpy.SelectLayerByAttribute_management(slackLoops, 'NEW_SELECTION', "placement = 'AE'")
    ws['E44'] = math.ceil((sumField(aerFiber, 'LENGTH_GEO') * 1.05) + sumField(aerLoops, 'loop_length'))


    # Cable sizes (24, 48, 96, 144, 288, etc) (for cells E102 - E107)
    lcpFiberSizes = arcpy.Dissolve_management(totalFiber, f'{lcpNameFixed}_FiberSizes', 'fibercount', [['LENGTH_GEO', 'SUM']])
    fiber12 = trySelect(lcpFiberSizes, 'fibercount', 12)
    fiber24 = trySelect(lcpFiberSizes, 'fibercount', 24)
    fiber48 = trySelect(lcpFiberSizes, 'fibercount', 48)
    fiber96 = trySelect(lcpFiberSizes, 'fibercount', 96)
    fiber144 = trySelect(lcpFiberSizes, 'fibercount', 144)
    fiber288 = trySelect(lcpFiberSizes, 'fibercount', 288)
    ws['E102'] = fiberCalcs('288')
    ws['E103'] = fiberCalcs('144')
    ws['E104'] = fiberCalcs('96')
    ws['E105'] = fiberCalcs('48')
    ws['E106'] = fiberCalcs('24')
    ws['E107'] = fiberCalcs('12')

    # Ground Rods & Clamps = Total qty of unique cable names (not segments, names) x 2 (for cell 132)
    cur = arcpy.da.SearchCursor(totalFiber, ['cable_name'])
    uniqueNames = []
    cableNames = [row[0] for row in cur]
    [uniqueNames.append(name) for name in cableNames if name not in uniqueNames] # puts only unique cable names in uniqueNames list
    ws['E132'] = int(len(uniqueNames) * 2)
############# Supplemental Helper Functions #############

def sumField(fc, field):
    '''
    Sums fields in a given feature class (i.e. LENGTH_GEO, Storage loops length, etc.)
    :param fc: feature class to sum
    :return: Summed field, or 0 if none.
    '''
    try:
        cur = arcpy.da.SearchCursor(fc, [f'{field}'])
        value = 0
        for row in cur:
            if row[0] == None:
                pass
            else:
                value += row[0]
        return math.ceil(int(value))
    except Exception as error:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f'Exception has occured in function for feature class {fc}: {error} on line {exc_tb.tb_lineno}, {traceback.print_tb(exc_tb)}')
        return 0
def trySelect(fc, field, value):
    '''
    Uses the select by function in arcpy inside a try/except, circumvents error and returns 0 if a field is not present.
    :param fc: feature class
    :param field: field to query
    :param value: value to query field
    :return: Either feature class length or 0 if no field available
    '''
    try:
        select = arcpy.SelectLayerByAttribute_management(fc, 'NEW_SELECTION', f"{field} = '{value}'")
        cur = arcpy.da.SearchCursor(select, ['SUM_LENGTH_GEO'])
        fcLengthList = [row[0] for row in cur]
        fcLength = math.ceil(int(fcLengthList[0]))
        return fcLength
    except Exception as error:
        print(f'--{value} count fiber is not present in LCP--')
        # exc_type, exc_obj, exc_tb = sys.exc_info()
        # print(f'Exception has occured in function: {error} on line {exc_tb.tb_lineno}, {traceback.print_tb(exc_tb)}')
        return 0
def fiberCalcs(fibersize):
    '''
    Supplemental script to calculate lengths for E102-E107 to avoid any division by zero errors. Dislike this way to calulate these lengths, hard to create iteration here given uniqueness of calulations.
    :param fibersize: fiber size feature class
    :param structure: structure feature class
    :return: footage of fiber if non-zero, zero otherwise.
    '''

    if '288' in fibersize:
        try:
            result = math.ceil(fiber288 * 1.07 + (((fiber288 / (fiber144 + fiber288)) * float(arcpy.GetCount_management(vaults)[0])) * 100))
            return result
        except ZeroDivisionError:
            return 0
    if '144' in fibersize:
        try:
            result = math.ceil(fiber144 * 1.07 + (((fiber144 / (fiber144 + fiber288)) * float(arcpy.GetCount_management(vaults)[0])) * 100))
            return result
        except ZeroDivisionError:
            return 0
    if '96' in fibersize:
        try:
            result = math.ceil(fiber96 * 1.07 + (float(arcpy.GetCount_management(largePeds)[0]) * 50))
            return result
        except ZeroDivisionError:
            return 0
    if '48' in fibersize:
        try:
            result = math.ceil(fiber48 * 1.07 + (((fiber48 / (fiber48 + fiber24) + fiber12) * float(arcpy.GetCount_management(mediumPeds)[0])) * 50))
            return result
        except ZeroDivisionError:
            return 0
    if '24' in fibersize:
        try:
            result = math.ceil(fiber24 * 1.07 + (((fiber24 / (fiber48 + fiber24) + fiber12) * float(arcpy.GetCount_management(mediumPeds)[0])) * 50))
            return result
        except ZeroDivisionError:
            return 0
    if '12' in fibersize:
        try:
            result = math.ceil(fiber12 * 1.07 + (((fiber12 / (fiber48 + fiber24) + fiber12) * float(arcpy.GetCount_management(mediumPeds)[0])) * 50))
            return result
        except ZeroDivisionError:
            return 0
############# Main #############
def main(): # lcp will need to be var in main to pass through script.
    start = time.time()
    # prep functions
    # downloadGDB()
    # unzip()
    # clear_gdb()
    # transfer()
    # Create BOMs
    lcps = ['ESC-C02', 'ESC-C04']
    for lcp in lcps:
        print(f'Starting {lcp} BOM Creation')
        print('--------------')
        lcpNameFixed = lcp.replace('-', '_')
        getLCPFeatures(lcp, lcpNameFixed)
        addresses(lcpNameFixed)
        dropFiber(lcpNameFixed)
        conduit(lcpNameFixed)
        spliceClosures(lcpNameFixed)
        structures()
        fiber(lcpNameFixed)
        # Save outfile, exports to outpath location
        outfile = f"{outpath}\\{lcp}_BOM_{date.strftime('%Y%m%d')}.xlsx"
        wb.save(outfile)
        print('\n')
    end = time.time()
    print(f'Total Script Duration (minutes) = {(end - start) / 60}')
if __name__ == '__main__':
    main()
