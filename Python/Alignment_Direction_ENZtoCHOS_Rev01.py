# Topic; Alignment Direction - Eaasting, Northing, Elevation to Chainage, H.Offset, V.Offset
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 02/04/2024

# Import Module
import math # for General Survey Function
import numpy as np
import pandas as pd
import time

t0 = time.time() # Start time

#---------------------------- All Function ----------------------------#
## General Survey ##
# Convert Degrees to Radians
def DegtoRad(d):
    ang = d * math.pi / 180.0
    return ang

# Convert Radians to Degrees
def RadtoDeg(d):
    ang = d * 180 / math.pi
    return ang

# Compute Distance and Azimuth from 2 Points
def DirecAziDist(EStart, NStart, EEnd, NEnd):
    dE = EEnd - EStart
    dN = NEnd - NStart
    Dist = math.sqrt(dE**2 + dN**2)

    if dN == 0:
        if dN == 0 and dE > 0:
            Azi = 90
        elif dN < 0 and dE == 0:
            Azi = 180
        elif dN == 0 and dE < 0:
            Azi = 270
        else:
            Azi = False

    elif dN != 0:
        ang = math.atan2(dE, dN)
        if ang >= 0:
            Azi = RadtoDeg(ang)
        else:
            Azi = RadtoDeg(ang) + 360

    return Dist, Azi

# Check Azimuth
def ModAzi(AziA):
    k = int(AziA / 360)
    if AziA >= 360 * k:
        FinalAZ = AziA - (360 * k)
    elif AziA < 0:
        FinalAZ = AziA + (360 * k)
    else:
        False
    return FinalAZ

# Convert Degrees to dd-mm-ss (String)
def DegtoDMSStr1(deg):
    d = abs(deg)
    mm, ss = divmod(d * 3600, 60)
    dd, mm = divmod(mm, 60)
    return '{:.0f}-{:.0f}-{:.2f}'.format(dd, mm, ss)

## Spiral Curve (Clothoid) ##
def SpiralXY(L, LS, Radius):
    C = LS * abs(Radius)
    LengthX = L - (L ** 5 / 40 / C ** 2) + (L ** 9 / 3456 / C ** 4) - (L ** 13 / 599040 / C ** 6) + (L ** 17 / 175472640 / C ** 8) - (L ** 21 / 78033715000 / C ** 10)
    LengthY = (L ** 3 / 6 / C) - (L ** 7 / 336 / C ** 3) + (L ** 11 / 42240 / C ** 5)  - (L ** 15 / 9676800 / C ** 7) + (L ** 19 / 3530096640 / C ** 9) - (L ** 23 / 1880240947000 / C ** 11)
    AngleXY = RadtoDeg(np.arctan(LengthY / LengthX))
    Length = LengthY / np.sin(DegtoRad(AngleXY))    
    return AngleXY, Length
#---------------------------- End All Function ----------------------------#

#---------------------------- Main Alignment Direction Computation ----------------------------#

# Path files
Import_enz_data_path = "Import Alignment Direction Data.xlsx"
Import_hor_array_path = "Export Hor-Alignment.xlsx"
Import_ver_array_path = "Export Ver-Alignment.xlsx"
Export_data_path = "Export Alignment-ENZ to CHOS.xlsx"

# Import data excel file to Data frame 
df_ENZ_DATA = pd.read_excel(Import_enz_data_path, "3D-COORDINATE DATA")
df_HOR_ARRAY = pd.read_excel(Import_hor_array_path,"HOR-ARRAY")
df_VER_ARRAY = pd.read_excel(Import_ver_array_path,"VER-ARRAY")

# Count data
totalENZ = df_ENZ_DATA["EASTHING (M.)"].count() - 1
totalHAR = df_HOR_ARRAY["LOOP NO."].count() - 1
totalVAR = df_VER_ARRAY["LOOP NO."].count() - 1

# When compute alignment finish then record to Data Frame
ColumnNames_REULT = ["NO.", "CHAINAGE (M.)", "HOR. OFFSET (M.)", "VER. OFFSET (M.)", "EASTHING (M.)", "NORTHING (M.)", "ELEVETION (M.)", "AZI (Deg.)", "AZI (DD-MM-SS)", "HOR. ELEMENT", "VER. ELEMENT","REMARK"]
df_ALIGN_RESULT = pd.DataFrame(columns= ColumnNames_REULT)

for i in range(totalENZ + 1):

    PntNO = df_ENZ_DATA["NO."][i]
    EFind = df_ENZ_DATA["EASTHING (M.)"][i]
    NFind = df_ENZ_DATA["NORTHING (M.)"][i]
    ZFind = df_ENZ_DATA["ELEVETION (M.)"][i]

    StaFirst = df_HOR_ARRAY["CH.START (m.)"][0]
    EFirst = df_HOR_ARRAY["E.START (m.)"][0]
    NFirst = df_HOR_ARRAY["N.START (m.)"][0]
    Azimuth1 = df_HOR_ARRAY["AZIMUTH (deg.)"][0]
    LL, Azimuth2 = DirecAziDist(EFirst, NFirst, EFind, NFind)
    Angle12 = ModAzi(Azimuth1 - Azimuth2)
    Li = StaFirst + LL * np.cos(DegtoRad(Angle12))
    #Finish = False
    LErr = 0

    while True:

        Li = Li - LErr
        ## Horizontal alignment computation ##
        # Hor. Array data
        for u in range(totalHAR + 1):
            CHStart = df_HOR_ARRAY["CH.START (m.)"][u]
            CHEnd = df_HOR_ARRAY["CH.END (m.)"][u]
            if Li >= CHStart and Li < CHEnd:
                EStart = df_HOR_ARRAY["E.START (m.)"][u]
                NStart = df_HOR_ARRAY["N.START (m.)"][u]
                Azi = df_HOR_ARRAY["AZIMUTH (deg.)"][u]
                Radius = df_HOR_ARRAY["RADIUS (m.)"][u]
                TypeofHC = df_HOR_ARRAY["CURVE TYPE"][u]
                break # exit the loop
        
        # Case Tangent
        if TypeofHC == "T":
            Lengtha = Li - CHStart
            EML = EStart + Lengtha * np.sin(DegtoRad(Azi))
            NML = NStart + Lengtha * np.cos(DegtoRad(Azi))
            WCB = Azi
            HorType = "Linear"

        # Case Circular Curve
        elif TypeofHC == "C":
            Lengtha = Li - CHStart
            Angled = RadtoDeg(Lengtha / Radius)
            Cord = 2 * Radius * np.sin(DegtoRad(Angled / 2))
            AziStartToFind = ModAzi(Azi + Angled / 2)
            EML = EStart + Cord * np.sin(DegtoRad(AziStartToFind))
            NML = NStart + Cord * np.cos(DegtoRad(AziStartToFind))
            WCB = ModAzi(Azi + Angled)
            HorType = "Circular"
        
        # Case Spiral Curve
        else:
            if Li == CHStart:
                EML = EStart
                NML = NStart
                WCB = Azi
                HorType = "Clothoid"
            else:
                LS = CHEnd - CHStart
                if TypeofHC == "SPIN": # Spiral In
                    L = Li - CHStart
                else: # Spiral Out
                    L = CHEnd - Li

                    Qs = 90 * LS / math.pi / abs(Radius)
                    Qx, Lx = SpiralXY(LS, LS, Radius)
                    Qa = Qs - Qx
                    AziA = ModAzi(Azi + Qa * np.sign(Radius))
                    EStart = EStart + Lx * np.sin(DegtoRad(AziA))
                    NStart = NStart + Lx * np.cos(DegtoRad(AziA))
                    Azi = ModAzi(AziA + Qx * np.sign(Radius) + 180)
                    Radius = Radius * -1

                AngleXY, Cord  = SpiralXY(L, LS, Radius)
                AziStartToFind = ModAzi(Azi + AngleXY * np.sign(Radius))
                EML = EStart + Cord * np.sin(DegtoRad(AziStartToFind))
                NML = NStart + Cord * np.cos(DegtoRad(AziStartToFind))
                AngleQ = 90 * L ** 2 / LS / math.pi / Radius

                if TypeofHC == "SPIN": # Spiral In
                    WCB = ModAzi(Azi + AngleQ)
                else: # Spiral Out
                    WCB = ModAzi(Azi + AngleQ + 180)
                
                HorType = "Clothoid"

        Direction = ModAzi(WCB + 90 * np.sign(0))
        EMLLi = EML + abs(0) * np.sin(DegtoRad(Direction))
        NMLLi = NML + abs(0) * np.cos(DegtoRad(Direction))
        Azimuth3  = WCB
        Azimuth4 = ModAzi(Azimuth3 - 90)
        LMLtoF, Azimuth5 = DirecAziDist(EMLLi, NMLLi, EFind, NFind)
        AngleErr = Azimuth4 - Azimuth5
        LErr = LMLtoF * np.sin(DegtoRad(AngleErr))
        N = 1000
        if int(LErr * N) == 0:
            #Finish = True
            break

    # Result : Hor.Alignment
    CH_FINAL = Li - LErr
    HOS_FINAL = LMLtoF * np.sign(np.sin(Azimuth5 - Azimuth3))
    WCB_FINAL = Azimuth3
    HOR_GEOMETRY = HorType

    ## Vertical alignment computation ##
    # Ver. Array data
    for v in range(totalVAR + 1):
        CHStart = df_VER_ARRAY["CH.START (m.)"][v]
        CHEnd = df_VER_ARRAY["CH.END (m.)"][v]
        if CH_FINAL >= CHStart and CH_FINAL < CHEnd:
            LevelStart = df_VER_ARRAY["ELEVATION (m.)"][v]
            g1 = df_VER_ARRAY["GRADIENT 1 (%)"][v]
            g2 = df_VER_ARRAY["GRADIENT 2 (%)"][v]
            LVC = df_VER_ARRAY["LVC (m.)"][v]
            LVCL = df_VER_ARRAY["LVC 1 (m.)"][v]
            LVCR = df_VER_ARRAY["LVC 2 (m.)"][v]
            TypeofVC = df_VER_ARRAY["CURVE TYPE"][v]
            break # exit the loop
        else:
            continue
    
    # Case Tangent
    if TypeofVC == "T":
       LFind = CH_FINAL - CHStart
       LevelFind = LevelStart + (g1 / 100 * LFind)
       VerType = "Linear"

    # Case Symmetric Curve
    elif TypeofVC == "S":
        LFind = CH_FINAL - CHStart
        LevelFind = LevelStart + (g1 / 100 * LFind)
        Y = (g2 - g1) * LFind ** 2 / LVC / 200
        LevelFind = LevelFind + Y
        VerType = "Parabola"
    
    # Case Unsymmetric Curve
    elif TypeofVC == "U":
        LFind = CH_FINAL - CHStart
        if CH_FINAL <= CHStart + LVCL:
           Y = (g2 - g1) * LVCL * LVCR / 200 / (LVCL + LVCR) * (LFind / LVCL) ** 2
           LevelFind = LevelStart + (g1 / 100 * LFind)
        else:
           Y = (g2 - g1) * LVCL * LVCR / 200 / (LVCL + LVCR) * ((CHEnd - CH_FINAL) / LVCR) ** 2
           LevelFind = LevelStart + (g1 / 100 * LVCL) + (g2 / 100 * (CH_FINAL - CHStart - LVCL))
        LevelFind = LevelFind + Y
        VerType = "Parabola"

    # Result : Ver.Alignment
    VOS_FINAL = ZFind - LevelFind
    VER_GEOMETRY = VerType

    # Add aligment data to data frame df_ALIGN_RESULT
    df_ALIGN = pd.DataFrame([[PntNO, CH_FINAL, HOS_FINAL, VOS_FINAL, EFind, NFind, ZFind, WCB_FINAL, DegtoDMSStr1(WCB_FINAL), HOR_GEOMETRY , VER_GEOMETRY , ""]], columns=ColumnNames_REULT)
    df_ALIGN_RESULT = df_ALIGN_RESULT._append(df_ALIGN, ignore_index=True)

# Export Alignment Result
with pd.ExcelWriter(Export_data_path) as writer:
    df_ALIGN_RESULT.to_excel(writer, sheet_name="ALIGN.RESULT", index = False)

t = time.time() # End time
print("Alignment was computed completely!, {:.3f}sec.".format(t-t0))
#---------------------------- End Alignment Direction Computation ----------------------------#