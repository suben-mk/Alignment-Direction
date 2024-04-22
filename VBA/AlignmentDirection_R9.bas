Attribute VB_Name = "AlignmentDirection_R9"
 ' Topic; Alignment Direction
 ' Created By; Unknown
 ' Dated; 31/05/1997
 ' Updated; 19/10/2023
 
 Option Base 1
 Const Pi As Single = 3.141592654

'Compute Coordinate and Azimuth of Horizontal Alignment.
'By Chainage and Offset.
'Result; E=Easting, N=Northing, W=Azimuth on alignment

 Function Hor(StaFind, Offset, HorArray, ENWD)
   Dim TypeofC As String
   Dim NoOfData As Integer
   Dim I As Integer
     
   NoOfData = HorArray(1, 1)
   StaMin = HorArray(1, 2)
   StaMax = HorArray(NoOfData, 3)
   
   If StaFind < StaMin Or StaFind > StaMax Then
       Hor = "Out Of Range "
       Exit Function
   End If
   
'Code is resvised by SBM updated 11.01.2023
   For I = 1 To NoOfData
      StaStart = HorArray(I, 2)
      StaEnd = HorArray(I, 3)
    If StaFind >= StaStart And StaFind < StaEnd Then
      EStart = HorArray(I, 4)
      NStart = HorArray(I, 5)
      Azi = HorArray(I, 6)
      Radius = HorArray(I, 7)
      TypeofC = UCase$(HorArray(I, 8))
      Exit For
    End If
   Next
   
   Select Case TypeofC
    Case "T"   'Tangent
        Lengtha = StaFind - StaStart
        EML = EStart + Lengtha * Sin(Azi * Pi / 180)
        NML = NStart + Lengtha * Cos(Azi * Pi / 180)
        WCB = Azi
        
    Case "C"   'Circular
        Lengtha = StaFind - StaStart
        Angled = 180 * Lengtha / Pi / Radius
        Cord = 2 * Radius * Sin(Angled / 2 * Pi / 180)
        AziStartToFind = ModAzi(Azi + Angled / 2)
        EML = EStart + Cord * Sin(AziStartToFind * Pi / 180)
        NML = NStart + Cord * Cos(AziStartToFind * Pi / 180)
        WCB = ModAzi(Azi + Angled)
        
    Case Else     'Spiral
     If StaFind = StaStart Then
        EML = EStart
        NML = NStart
        WCB = Azi
     Else
        LS = StaEnd - StaStart
        If TypeofC = "SPIN" Then     'Spiral-In
         L = StaFind - StaStart
        Else                         'Spiral-Out
         L = StaEnd - StaFind
         Qs = 90 * LS / Pi / Abs(Radius)
         Qx = SpiralXY(LS, LS, Radius, "A")
         Lx = SpiralXY(LS, LS, Radius, "L")
         Qa = Qs - Qx
         AziA = ModAzi(Azi + Qa * Sgn(Radius))
         EStart = EStart + Lx * Sin(AziA * Pi / 180)
         NStart = NStart + Lx * Cos(AziA * Pi / 180)
         Azi = ModAzi(AziA + Qx * Sgn(Radius) + 180)
         Radius = Radius * -1
        End If
        
        AngleXY = SpiralXY(L, LS, Radius, "A")
        Cord = SpiralXY(L, LS, Radius, "L")
        AziStartToFind = ModAzi(Azi + AngleXY * Sgn(Radius))
        EML = EStart + Cord * Sin(AziStartToFind * Pi / 180)
        NML = NStart + Cord * Cos(AziStartToFind * Pi / 180)
        AngleQ = 90 * L ^ 2 / LS / Pi / Radius
        If TypeofC = "SPIN" Then
          WCB = ModAzi(Azi + AngleQ)
        Else
          WCB = ModAzi(Azi + AngleQ + 180)
        End If
     End If
    End Select 'Case TypeOfC
             
    Direction = ModAzi(WCB + 90 * Sgn(Offset))
   
   Select Case UCase$(ENWD)
    Case "E"
            Hor = EML + Abs(Offset) * Sin(Direction * Pi / 180)
    Case "N"
            Hor = NML + Abs(Offset) * Cos(Direction * Pi / 180)
    Case "W"
            Hor = WCB
    Case "D"
            Hor = Direction
   End Select
 End Function 'Hor Funtion

'Check Azimuth for Hor Function.

 Private Function ModAzi(AziA)
      Dim k As Integer
      k = Int(AziA / 360)
      If AziA >= 360 * k Then
         ModAzi = AziA - (360 * k)
      ElseIf AziA < 0 Then
         ModAzi = (360 * k) + AziA
      End If
 End Function 'ModAZi
            
'Spiral Curve by Clothoid formular for Hor Function.

 Private Function SpiralXY(L, LS, Radius, AL)
   Dim C As Double
   Dim LengthX As Double
   Dim LengthY As Double
   Dim AngleXY As Double
   Dim Length As Double
   
   C = LS * Abs(Radius)
   LengthX = L - (L ^ 5 / 40 / C ^ 2) + (L ^ 9 / 3456 / C ^ 4) _
              - (L ^ 13 / 599040 / C ^ 6) + (L ^ 17 / 175472640 / C ^ 8) _
              - (L ^ 21 / 78033715000# / C ^ 10)
   LengthY = (L ^ 3 / 6 / C) - (L ^ 7 / 336 / C ^ 3) + (L ^ 11 / 42240 / C ^ 5) _
              - (L ^ 15 / 9676800 / C ^ 7) + (L ^ 19 / 3530096640# / C ^ 9) _
              - (L ^ 23 / 1880240947000# / C ^ 11)
   
   AngleXY = Atn(LengthY / LengthX) * 180 / Pi
   Length = LengthY / Sin(AngleXY * Pi / 180)
   Select Case UCase$(AL)
    Case "A"
       SpiralXY = AngleXY
    Case "L"
       SpiralXY = Length
    Case Else
       SpiralXY = False
   End Select
 End Function 'SpiralXY
 
'Azimuth for Hor Function.
 Private Function Azimuth(EStart, NStart, EEnd, NEnd)
     Dim dE As Single
     Dim DN As Single
     Dim dd As Single
     
     dE = EEnd - EStart: DN = NEnd - NStart
     If DN <> 0 Then dd = Atn(dE / DN) * 180 / Pi
      If DN = 0 Then
           If dE > 0 Then
              Azimuth = 90
           ElseIf dE < 0 Then
              Azimuth = 270
           Else
              Azimuth = False
           End If
      ElseIf DN > 0 Then
           If dE > 0 Then
              Azimuth = dd
           ElseIf dE < 0 Then
              Azimuth = 360 + dd
           End If
      ElseIf DN < 0 Then
              Azimuth = 180 + dd
      End If
 End Function 'Azimuth

'Distance for Hor Function.

 Private Function Dist(X1, Y1, X2, Y2)
      Dist = Sqr((X2 - X1) ^ 2 + (Y2 - Y1) ^ 2)
 End Function 'Dist

'Compute Chainage and Offset of Horizontal Alignment.
'By Coordinate; Easting and Northing
'Result; C=Chainage and O=Offset

 Function HorP(EFind, NFind, HorArray, CO, Start)
  Dim Finish As Boolean
  Dim N As Integer
  Dim k As Integer
  If Start > HorArray(1, 1) Then Start = HorArray(1, 1)
  If Start <= 0 Or Start = "" Then
   k = 1
  Else
   k = Start
  End If
  StaFirst = HorArray(k, 2)
  EFirst = HorArray(k, 4)
  NFirst = HorArray(k, 5)
  Azimuth1 = HorArray(k, 6)
  Azimuth2 = Azimuth(EFirst, NFirst, EFind, NFind)
  Angle12 = ModAzi(Azimuth1 - Azimuth2)
  LL = Dist(EFirst, NFirst, EFind, NFind)
  Lx = StaFirst + LL * Cos(Angle12 * Pi / 180)
  Finish = False
  LErr = 0
  
  Do
    Lx = Lx - LErr
    EMLLx = Hor(Lx, 0, HorArray, "E")
    NMLLx = Hor(Lx, 0, HorArray, "N")
    Azimuth3 = Hor(Lx, 0, HorArray, "W")
    Azimuth4 = ModAzi(Azimuth3 - 90)
    Azimuth5 = Azimuth(EMLLx, NMLLx, EFind, NFind)
    LMLtoF = Dist(EMLLx, NMLLx, EFind, NFind)
    AngleErr = Azimuth4 - Azimuth5
    LErr = LMLtoF * Sin(AngleErr * Pi / 180)
    N = 1000
    If Int(LErr * N) = 0 Then
      Finish = True
    End If
  Loop Until Finish
    
  OffsetF = LMLtoF * Sgn(Sin(Azimuth5 - Azimuth3))
       
  Select Case UCase$(CO)
     Case "C"
             HorP = Lx - LErr
     Case "O"
             HorP = OffsetF
     Case Else
             HorP = False
  End Select
 End Function 'HorP Funtion
 
'Compute Elevation (PG.) of Vertical Alignment (Symmetry and Unsymmetry).
'By Chainage.
'Result; Elevation (PG.)

'Code is resvised by SBM updated 01.11.2023
 Function Ver(StaFind, VerArray)
   Dim I As Integer
   Dim NoData As Integer
   
   NoData = VerArray(1, 1)
   StaMin = VerArray(1, 2)
   StaMax = VerArray(NoData, 3)
  
   If StaFind < StaMin Or StaFind > StaMax Then
     Ver = "Out Of Range"
     Exit Function
   End If
   
   For I = 1 To NoData
     StaStart = VerArray(I, 2)
     StaEnd = VerArray(I, 3)
      
     If StaFind >= StaStart And StaFind < StaEnd Then
       LevelStart = VerArray(I, 4)
       g1 = VerArray(I, 5)
       g2 = VerArray(I, 6)
       LVC = VerArray(I, 7)
       LVCL = VerArray(I, 8)
       LVCR = VerArray(I, 9)
       TypeofC = UCase$(VerArray(I, 10))
       Exit For
     End If
   Next
   
   Select Case TypeofC
    Case "T"   'Tangent
       LFind = StaFind - StaStart
       LevelFind = LevelStart + (g1 / 100 * LFind)
       Ver = LevelFind
       
    Case "S"   'Symmetric Curve
       LFind = StaFind - StaStart
       LevelFind = LevelStart + (g1 / 100 * LFind)
       Y = (g2 - g1) * LFind ^ 2 / LVC / 200
       Ver = LevelFind + Y
    
    Case "U"   'Unsymmetric Curve
        LFind = StaFind - StaStart
        If StaFind <= StaStart + LVCL Then
            Y = (g2 - g1) * LVCL * LVCR / 200 / (LVCL + LVCR) * (LFind / LVCL) ^ 2
            LevelFind = LevelStart + (g1 / 100 * LFind)
        Else
            Y = (g2 - g1) * LVCL * LVCR / 200 / (LVCL + LVCR) * ((StaEnd - StaFind) / LVCR) ^ 2
            LevelFind = LevelStart + (g1 / 100 * LVCL) + (g2 / 100 * (StaFind - StaStart - LVCL))
        End If
        Ver = LevelFind + Y
    End Select
    
 End Function 'Ver Function

'Compute Superelevation (Cross fall%) by Chainage.
'Result; Cross fall%

 Function XFall(StaFind, XFallArray)
   Dim I As Integer
   Dim NoData As Integer
   
   NoData = XFallArray(1, 1)
   StaMin = XFallArray(1, 2)
   StaMax = XFallArray(NoData, 3)
   If StaFind < StaMin Or StaFind > StaMax Then
     XFall = "Out Of Range"
     Exit Function
   End If
   For I = 1 To NoData
     StaStart = XFallArray(I, 2)
     StaEnd = XFallArray(I, 3)
     
     If StaFind >= StaStart And StaFind < StaEnd Then
       XF1 = XFallArray(I, 4)
       XF2 = XFallArray(I, 5)
       TypeX = UCase$(XFallArray(I, 6))
       Exit For
     End If
   Next
   
   If TypeX = "N" Or XF1 = XF2 Then
     XFall = XF1
   Else
     L = StaEnd - StaStart
     LFind = StaFind - StaStart
     If TypeX = "V" Then
       Y = (XF2 - XF1) * LFind / L
     Else
       Y = (XF2 - XF1) * (LFind / L) ^ 2 * (3 - 2 * LFind / L)
     End If
     XFall = XF1 + Y
   End If
 End Function 'XFall Function
