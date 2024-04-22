# Topic; Alignment Direction - Chainage, H.Offset, V.Offset to Eaasting, Northing, Elevation
# Created By; Suben Mukem (SBM) as Survey Engineer.
# Updated; 04/04/2024

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

    if dN != 0:
        ang = math.atan2(dE, dN)
    else:
        Azi = False

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
Import_chos_data_path = "Import Alignment Direction Data.xlsx"
Import_hor_array_path = "Export Hor-Alignment.xlsx"
Import_ver_array_path = "Export Ver-Alignment.xlsx"
Export_data_path = "Export Alignment-CHOS to ENZ.xlsx"

# Import data excel file to Data frame 
df_CHOS_DATA = pd.read_excel(Import_chos_data_path, "CHAINAGE&OFFSET DATA")
df_HOR_ARRAY = pd.read_excel(Import_hor_array_path,"HOR-ARRAY")
df_VER_ARRAY = pd.read_excel(Import_ver_array_path,"VER-ARRAY")

# Count data
totalCHOS = df_CHOS_DATA["CHAINAGE (M.)"].count() - 1
totalHAR = df_HOR_ARRAY["LOOP NO."].count() - 1
totalVAR = df_VER_ARRAY["LOOP NO."].count() - 1

# When compute alignment finish then record to Data Frame
ColumnNames_REULT = ["NO.", "CHAINAGE (M.)", "HOR. OFFSET (M.)", "VER. OFFSET (M.)", "EASTHING (M.)", "NORTHING (M.)", "ELEVETION (M.)", "AZI (Deg.)", "AZI (DD-MM-SS)", "HOR. ELEMENT", "VER. ELEMENT","REMARK"]
df_ALIGN_RESULT = pd.DataFrame(columns= ColumnNames_REULT)

for i in range(totalCHOS + 1):
    # Chainage, Hor.Offset, Ver.Offset data
    PntNO = df_CHOS_DATA["NO."][i]
    CHFind = df_CHOS_DATA["CHAINAGE (M.)"][i]
    HOS = df_CHOS_DATA["HOR. OFFSET (M.)"][i]
    VOS = df_CHOS_DATA["VER. OFFSET (M.)"][i]

    ## Horizontal alignment computation ##
    # Hor. Array data
    for u in range(totalHAR + 1):
        CHStart = df_HOR_ARRAY["CH.START (m.)"][u]
        CHEnd = df_HOR_ARRAY["CH.END (m.)"][u]
        if CHFind >= CHStart and CHFind < CHEnd:
            EStart = df_HOR_ARRAY["E.START (m.)"][u]
            NStart = df_HOR_ARRAY["N.START (m.)"][u]
            Azi = df_HOR_ARRAY["AZIMUTH (deg.)"][u]
            Radius = df_HOR_ARRAY["RADIUS (m.)"][u]
            TypeofHC = df_HOR_ARRAY["CURVE TYPE"][u]
            break # exit the loop
    
    # Case Tangent
    if TypeofHC == "T":
        Lengtha = CHFind - CHStart
        EML = EStart + Lengtha * np.sin(DegtoRad(Azi))
        NML = NStart + Lengtha * np.cos(DegtoRad(Azi))
        WCB = Azi
        HorType = "Linear"

    # Case Circular Curve
    elif TypeofHC == "C":
        Lengtha = CHFind - CHStart
        Angled = RadtoDeg(Lengtha / Radius)
        Cord = 2 * Radius * np.sin(DegtoRad(Angled / 2))
        AziStartToFind = ModAzi(Azi + Angled / 2)
        EML = EStart + Cord * np.sin(DegtoRad(AziStartToFind))
        NML = NStart + Cord * np.cos(DegtoRad(AziStartToFind))
        WCB = ModAzi(Azi + Angled)
        HorType = "Circular"
    
    # Case Spiral Curve
    else:
        if CHFind == CHStart:
            EML = EStart
            NML = NStart
            WCB = Azi
            HorType = "Clothoid"
        else:
            LS = CHEnd - CHStart
            if TypeofHC == "SPIN": # Spiral In
                L = CHFind - CHStart
            else: # Spiral Out
                L = CHEnd - CHFind

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

    # Result : Hor.Alignment
    Direction = ModAzi(WCB + 90 * np.sign(HOS))
    EML_FINAL = EML + abs(HOS) * np.sin(DegtoRad(Direction))
    NML_FINAL = NML + abs(HOS) * np.cos(DegtoRad(Direction))
    WCB_FINAL = WCB
    HOR_GEOMETRY = HorType

    ## Vertical alignment computation ##
    # Ver. Array data
    for v in range(totalVAR + 1):
        CHStart = df_VER_ARRAY["CH.START (m.)"][v]
        CHEnd = df_VER_ARRAY["CH.END (m.)"][v]
        if CHFind >= CHStart and CHFind < CHEnd:
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
       LFind = CHFind - CHStart
       LevelFind = LevelStart + (g1 / 100 * LFind)
       VerType = "Linear"

    # Case Symmetric Curve
    elif TypeofVC == "S":
        LFind = CHFind - CHStart
        LevelFind = LevelStart + (g1 / 100 * LFind)
        Y = (g2 - g1) * LFind ** 2 / LVC / 200
        LevelFind = LevelFind + Y
        VerType = "Parabola"
    
    # Case Unsymmetric Curve
    elif TypeofVC == "U":
        LFind = CHFind - CHStart
        if CHFind <= CHStart + LVCL:
           Y = (g2 - g1) * LVCL * LVCR / 200 / (LVCL + LVCR) * (LFind / LVCL) ** 2
           LevelFind = LevelStart + (g1 / 100 * LFind)
        else:
           Y = (g2 - g1) * LVCL * LVCR / 200 / (LVCL + LVCR) * ((CHEnd - CHFind) / LVCR) ** 2
           LevelFind = LevelStart + (g1 / 100 * LVCL) + (g2 / 100 * (CHFind - CHStart - LVCL))
        LevelFind = LevelFind + Y
        VerType = "Parabola"

    # Result : Ver.Alignment
    EL_FINAL = LevelFind + VOS
    VER_GEOMETRY = VerType

    # Add aligment data to data frame df_ALIGN_RESULT
    df_ALIGN = pd.DataFrame([[PntNO, CHFind, HOS, VOS, EML_FINAL, NML_FINAL, EL_FINAL, WCB_FINAL, DegtoDMSStr1(WCB_FINAL), HOR_GEOMETRY, VER_GEOMETRY, ""]], columns=ColumnNames_REULT)
    df_ALIGN_RESULT = df_ALIGN_RESULT._append(df_ALIGN, ignore_index=True)

# Export Alignment Result
with pd.ExcelWriter(Export_data_path) as writer:
    df_ALIGN_RESULT.to_excel(writer, sheet_name="ALIGN.RESULT", index = False)

t = time.time() # End time
print("Alignment was computed completely!, {:.3f}sec.".format(t-t0))
#---------------------------- End Alignment Direction Computation ----------------------------#