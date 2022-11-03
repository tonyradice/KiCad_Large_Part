"""
@Title: Large Part KiCad Part Generator Utility File
@Author: Tony Radice
@Date: 10/3/22

"""


def ParseCommandLine(args):
    # Arguments passed back to main routine:
    # [0] File Name WITH the PDef_ prefix (to be removed for output file)
    # Flag: Diagnostic Mode True or False
    #   May in future replace with an integer for level ...
    # Flag: Replace values in source sheet (Not implemented)

    Arguments = ['', False, False]
    i = len(args)

    while i >> 0:
        i -= 1
        cstring = args[i]
        # print('Arg-I-' + str(i) + ' is ' + cstring)

        if cstring.find("/") == -1:               # Not found
            Arguments[0] = cstring

        elif cstring.find("/D") != -1:            # Diagnostic Flag
            Arguments[1] = True

        elif cstring.find("/R") != -1:            # Rewrite Flag (future)
            Arguments[2] = True

    # print('Args-I- ', Arguments)
    return Arguments


def GetTopArgs(CWS, row):
    # Get the entire row of properties from the Table
    endcol = 11         # Currently through column "I"
    Prop = []
    RProps = {0: "Reference", 1: "Value", 2: "Footprint", 3: "Datasheet",
              4: "ki_locked", 5: "ki_description"}

    for i in range(1, endcol):
        value = CWS.cell(row=row, column=i).value
        Prop.append(value)

    # Set defaults
    if Prop[2] in range(0, 6):
        Prop[0] = RProps[Prop[2]]       # Substitute restricted words

    if Prop[3] is None:                 # Default X Value
        Prop[3] = -1.27

    if Prop[4] is None:                 # Default Y Value
        Prop[4] = 0

    if Prop[5] is None:                 # Default Rotation
        Prop[5] = 0

    if Prop[6] is None:                 # Default Text Size
        Prop[6] = 1.27

    if Prop[7] is None:                 # Default Number Size
        Prop[7] = 1.27

    return Prop


def WriteTopBlock(CWS, OFile):
    # Second Version of Write Top Block -
    # CWS: Must refer to Top Block (Non-Graphic) of the part worksheet
    # OFile: Output File Name

    OFile.write('(kicad_symbol_lib (version 20211014)')  # First Line
    OFile.write('(generator Large_Part_Generator)\n')
    row = 11
    col = 2
    just = {"L": 'left', "R": 'right', "T": 'top', "B": 'bottom'}

    SymbolName = CWS.cell(row=row, column=col).value
    OFile.write(' (symbol "')
    OFile.write(SymbolName + '" ')          # Library ID
    OFile.write('(in_bom yes) (on_board yes)\n')

    row = 10            # Determine if the Property is populated
    col = 1
    while CWS.cell(row=row, column=col).value is not None:
        Props = GetTopArgs(CWS, row)

        OFile.write('  (property "')
        OFile.write(Props[0])
        OFile.write('" "')
        OFile.write(str(Props[1]))  #
        OFile.write('" (id ')
        OFile.write(str(Props[2]))  # id Value
        OFile.write(') (at ')
        OFile.write(str(Props[3]) + ' ')
        OFile.write(str(Props[4]) + ' ')
        OFile.write(str(Props[5]) + ')\n')
        OFile.write('   (effects (font (size ')
        OFile.write(str(Props[6]) + ' ')
        OFile.write(str(Props[7]) + '))')

        if Props[8] is not None:
            OFile.write('(justify ')
            OFile.write(just[Props[8]])
            OFile.write(')')

        if Props[9] == 'hide':
            OFile.write('hide)')
        else:
            OFile.write(')')
        OFile.write('\n   )\n')
        row += 1
#        print('WTB-I-Property Line Returned: ')
#        print(Props)
#        print('\n')
    return


def GetPartName(TWS):
    # Simply return the contents of the Part Name cell
    PartName = TWS.cell(row=11, column=2).value
    return PartName


def SymbolClose(OFile):
    OFile.write('   )\n')
    return


def TempClose(OFile):
    OFile.write('  )\n )\n)\n')
    return


def MaxHeight(CWS):
    # What is max height of left or right columns in the Current Worksheet
    # Look at the "Type" Column and determine the largest of Left or Right
    # Returns an integer of number of pins on the side
    PTC = [3, 17]  # Columns describing Pin Types
    MaxHght = 0

    for CIndex in PTC:
        RIndex = 11  # Magic Number?
        MH = 0
        while CWS.cell(row=RIndex, column=CIndex).value is not None:
            RIndex += 1
            MH += 1
        if MH > MaxHght:
            MaxHght = MH

    return MaxHght


def GetBorders(CWS, MaxH):
    # Get the four points to determine the borders of the rectangle for the part
    # Note: Bottom Border has 2 added to it - 1 for top border offset from x axis
    #  and 1 to extend part BELOW bottom-most pin

    UnitScale = 1.27  # Converts to 0.05 inch
    LYBord = str(CWS.cell(row=2, column=2).value * -2 * UnitScale)
    RYBord = str(CWS.cell(row=2, column=2).value * 2 * UnitScale)
    TXBord = str(-2 * UnitScale)
    BXBord = str((MaxH + 2) * -2 * UnitScale)  # See Note
    return [LYBord, RYBord, TXBord, BXBord]


def CreateSymbolBodyHeader(PartName, CWS, MaxH, OFile):
    # Partname is for the Symbol field - Note must have additional append
    # CWS is Current Input Body Worksheet
    # MaxH is the Maximum Body Height determined by number of pins
    # OFile is Output File
    #
    # Assume Body will start on
    # Convert rectangle Left and Right, Top And Bottom Borders
    [LYBord, RYBord, TXBord, BXBord] = GetBorders(CWS, MaxH)

    # Set Default Values
    Stroke = str(CWS.cell(row=3, column=2).value)
    if CWS.cell(row=3, column=2).value is None:
        Stroke = '0'
    SType = CWS.cell(row=4, column=2).value
    if CWS.cell(row=4, column=2).value is None:
        SType = 'default'
    FType = str(CWS.cell(row=5, column=2).value)
    if CWS.cell(row=5, column=2).value is None:
        FType = 'none'

    # Identify Body Style == 1
    SName = PartName + '_1'

#   ToDo: May want a routine that can vary Symbol Header...
    OFile.write('  (symbol "' + SName + '"\n')
    OFile.write('    (rectangle (start ' + LYBord + " " + TXBord + ")")
    OFile.write('(end ' + RYBord + " " + BXBord + ')\n')
    OFile.write('      (stroke (width ' + Stroke + ')')
    OFile.write('(type ' + SType + ') (color 0 0 0 0))\n')
    OFile.write('      (fill (type ' + FType + '))\n')
    OFile.write('    )\n')
    return


def GetPinData(B, N, CWS):
    # CWS is Current Input Body Worksheet
    # Border - L, R, T or B - Determines Column and Offset
    # Number - Offset from Row 10 (Row data starts at row 11)
    Offset = {"L": 1, "R": 15, "T": 29, "B": 43}
    PD = ['', '', '', '', 0, 0, 0, 0, 0, 0, 0, 0]
    OFN = N + 10  # Row Data Offset
    so = Offset[B]

#   Wrote these out to identify what value IS
    PD[0] = CWS.cell(row=OFN, column=so).value              # Pin Name
    PD[1] = CWS.cell(row=OFN, column=(so + 1)).value        # Pin Number
    PD[2] = CWS.cell(row=OFN, column=(so + 2)).value        # Pin Type *
    PD[3] = CWS.cell(row=OFN, column=(so + 4)).value        # Pin Shape *
    PD[4] = CWS.cell(row=OFN, column=(so + 5)).value        # Location X
    PD[5] = CWS.cell(row=OFN, column=(so + 6)).value        # Location Y
    PD[6] = CWS.cell(row=OFN, column=(so + 7)).value        # Rotation
    PD[7] = CWS.cell(row=OFN, column=(so + 8)).value        # Length *
    PD[8] = CWS.cell(row=OFN, column=(so + 9)).value        # Name Size 1 *
    PD[9] = CWS.cell(row=OFN, column=(so + 10)).value       # Name Size 2 *
    PD[10] = CWS.cell(row=OFN, column=(so + 11)).value      # Nmbr Size 1 *
    PD[11] = CWS.cell(row=OFN, column=(so + 12)).value      # Nmbr Size 2 *
    #    print('PinData-D-: ',PD)

    # * Do I want to do mappings here?  May be easier...

    PTMap = {'PI': 'power_in', 'PO': 'power_out', 'I': 'input', 'O': 'output',
             'B': 'bidirectional', 'T': 'tri_state', 'P': 'passive',
             'F': 'free', 'U': 'unspecified', 'C': 'open_collector',
             'E': 'open_emitter', 'X': 'unconnected', '-': 'space'}

    # Need to clean up LS Map!
    LSMap = {'L': 'line', 'I': 'inverted', 'CL': 'clock',
             'ICL': 'inverted_clock', 'IL': 'input_low',
             'KL': 'clock_low', 'OL': 'output_low', 'EC': 'edge_clock_high',
             'NL': 'non logic'}

    # Mappings:
    # print('GPD-I-' + str(so) + ' row:' + str(OFN))
    PD[2] = PTMap[PD[2]]            # Pin Type Mapping

    if PD[3] is None:
        PD[3] = 'line'              # Pin Shape Default
    else:
        PD[3] = LSMap[PD[3]]

    if PD[7] is None:
        PD[7] = '2.54'              # Pin Length

    DefValue = 1.27                 # 0.05"
    for dval in (8, 9, 10, 11):     # Text Sizes
        if PD[dval] is None:
            PD[dval] = str(DefValue)

    #    print('PinData-D-: ',PD)
    return PD


def PlacePin(PDat, OFile):
    # Generate Pin description in outfile
    OFile.write('    (pin ')

    OFile.write(PDat[2] + ' ' + PDat[3])
    OFile.write(' (at ')
    OFile.write(str(PDat[4]) + ' ' + str(PDat[5]) + ' ' + str(PDat[6]))
    OFile.write(')(length ')
    OFile.write(str(PDat[7]) + ')\n')
    OFile.write('     (name "' + PDat[0] + '"')
    OFile.write(' (effects (font (size ')
    OFile.write(str(PDat[8]) + ' ' + str(PDat[9]) + '))))\n')
    OFile.write('     (number "' + str(PDat[1]) + '"')
    OFile.write(' (effects (font (size ')
    OFile.write(str(PDat[10]) + ' ' + str(PDat[11]) + '))))\n')
    OFile.write('    )\n')
    return


def PlacePinsInBody(CWS, MaxH, OFile):
    # CWS: Current Input Body Worksheet
    # OFile: Output File
    # MaxH: Height measurement driven by maximum # of pins
    # DPL: Pin Labels List - To be Returned
    #
    DPL = []
    UnitScale = 1.27  # Converts to 0.05 inch
    [LYBord, RYBord, TXBord, BXBord] = GetBorders(CWS, MaxH)
    PinData = {"L": 3, "R": 17, "T": 31, "B": 45}

    # Start with the Left Border
    # Start Pin 0.1" right of body and 0.1" below top border
    PXLoc = round(float(LYBord) - (2 * UnitScale), 2)  # Pin Initial X Location
    PYLoc = round(float(TXBord) - (2 * UnitScale), 2)  # Pin Initial Y Location
    SSRow = 11  # First Row of Data

    while CWS.cell(row=SSRow, column=PinData["L"]).value is not None:
        PDat = GetPinData("L", (SSRow - 10), CWS)
        # print('PPIB-DL-: ', PDat)
        if PDat[4] is None: PDat[4] = PXLoc
        if PDat[5] is None: PDat[5] = PYLoc
        if PDat[6] is None: PDat[6] = '0'           # Pin Rotation
        PYLoc = round((PYLoc - (2 * UnitScale)), 2)
        SSRow += 1
        if PDat[2] != 'space':
            PlacePin(PDat, OFile)
            DPL.append(str(PDat[1]))  # For Pin Count and Dup Check

    # Right Border
    # Start Pin 0.1" left of body and 0.1" below top border
    PXLoc = round(float(RYBord) + (2 * UnitScale), 2)  # Pin Initial X Location
    PYLoc = round(float(TXBord) - (2 * UnitScale), 2)  # Pin Initial Y Location
    SSRow = 11  # First Row of Data

    while CWS.cell(row=SSRow, column=PinData["R"]).value is not None:
        PDat = GetPinData("R", (SSRow - 10), CWS)
        # print('PPIB-DR-: ', PDat)
        if PDat[4] is None: PDat[4] = PXLoc
        if PDat[5] is None: PDat[5] = PYLoc
        if PDat[6] is None: PDat[6] = '180'  # Pin Rotation
        PYLoc = round((PYLoc - (2 * UnitScale)), 2)
        SSRow += 1
        if PDat[2] != 'space':
            PlacePin(PDat, OFile)
            DPL.append(str(PDat[1]))  # For Pin Count and Dup Check

    # Top Border
    # Start Pin 0.1" ABOVE body and 0.1" RIGHT of Left border
    PXLoc = round(float(LYBord) + (2 * UnitScale), 2)  # Pin Initial X Location
    PYLoc = round(float(TXBord) + (2 * UnitScale), 2)  # Pin Initial Y Location
    SSRow = 11  # First Row of Data

    while CWS.cell(row=SSRow, column=PinData["T"]).value is not None:
        PDat = GetPinData("T", (SSRow - 10), CWS)
        # print('PPIB-DT-: ', PDat)
        if PDat[4] is None: PDat[4] = PXLoc
        if PDat[5] is None: PDat[5] = PYLoc
        if PDat[6] is None: PDat[6] = '270'  # Pin Rotation

        # Pins start at top left (one indented) and move to right
        PXLoc = round((PXLoc + (2 * UnitScale)), 2)
        SSRow += 1
        if PDat[2] != 'space':
            PlacePin(PDat, OFile)
            DPL.append(str(PDat[1]))  # For Pin Count and Dup Check

    # Bottom Border
    # Start Pin 0.1" BELOW body and 0.1" RIGHT of Left border
    PXLoc = round(float(LYBord) + (2 * UnitScale), 2)  # Pin Initial X Location
    PYLoc = round(float(BXBord) - (2 * UnitScale), 2)  # Pin Initial Y Location
    SSRow = 11  # First Row of Data

    while CWS.cell(row=SSRow, column=PinData["B"]).value is not None:
        PDat = GetPinData("B", (SSRow - 10), CWS)
        # print('PPIB-DB-: ', PDat)
        if PDat[4] is None: PDat[4] = PXLoc
        if PDat[5] is None: PDat[5] = PYLoc
        if PDat[6] is None: PDat[6] = '90'  # Pin Rotation

        # Pins start at top left (one indented) and move to right
        PXLoc = round((PXLoc + (2 * UnitScale)), 2)
        SSRow += 1
        if PDat[2] != 'space':
            PlacePin(PDat, OFile)
            DPL.append(str(PDat[1]))  # For Pin Count and Dup Check

    return DPL


def CheckForDupPins(DPL):
    # Check for Duplicate Pin Numbers
    DPL.sort()
    for i in range(1, len(DPL)):
        if DPL[i] == DPL[i - 1]:
            print('Duplicate Pin Name - E - ' + DPL[i])
            return True
    return False


def CheckTopSheet(CWS):
    # Error Check for Top Sheet

    row = 10            # Determine if the Property is populated
    col = 1
    while CWS.cell(row=row, column=col).value is not None:
        Props = GetTopArgs(CWS, row)
        row += 1
        # Checks here?
        # Check fo justify valid entry (L, R, T, or B)

    return False


def CheckDataSheet(CWS):
# Current throws errors - this is why commented out at top...
    # Error Check for all body Sheets
    rowlimit = 200                              # Maximum Row

    for border in ('L', 'R', 'T', 'B'):
        row = 11                                # Initial row
        PD = GetPinData(border, row, CWS)
        # Check for Type of pin correct
        # Verify no blanks in type column

    return False
