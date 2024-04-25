# Alignment Direction
Alignment direction is computed from setting-out alignment data for tunnel survey work such as tunnel segment plan, tunnel boring mechine (TBM), setting-out point at site etc. I created the code to compute 3d-coordinate by known chainage &amp; offset and compute chainage &amp; offset by known 3d-coordinate. The code was created 2 languages which're python and vba excel.

## Workflow
### Python
  1. Prepare Alignment data as [Import Alignment Direction Data.xlsx](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/Python/Import%20Data/Import%20Alignment%20Direction%20Data.xlsx)
  2. Set path file
     
     [**Alignment_Direction_CHOStoENZ_Rev01.py**](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/Python/Alignment_Direction_CHOStoENZ_Rev01.py)
      ```py
      # Path files
      Import_chos_data_path = "Import Alignment Direction Data.xlsx"
      Import_hor_array_path = "Export Hor-Alignment.xlsx"
      Import_ver_array_path = "Export Ver-Alignment.xlsx"
      Export_data_path = "Export Alignment-CHOS to ENZ.xlsx"
      ```
     [**Alignment_Direction_ENZtoCHOS_Rev01.py**](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/Python/Alignment_Direction_ENZtoCHOS_Rev01.py)
      ```py
      # Path files
      Import_enz_data_path = "Import Alignment Direction Data.xlsx"
      Import_hor_array_path = "Export Hor-Alignment.xlsx"
      Import_ver_array_path = "Export Ver-Alignment.xlsx"
      Export_data_path = "Export Alignment-ENZ to CHOS.xlsx"
      ```
  3. Run python file
### VBA
  1. Open file [**VBA - Alignment Direction & Survey Program Rev.09.xlsm**](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/VBA/VBA%20-%20Alignment%20Direction%20%26%20Survey%20Program%20Rev.09.xlsm)
  2. Add setting-out alignmet array sheet to **VBA - Alignment Direction & Survey Program Rev.09.xlsm** and set array by name manager
     ![image](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/assets/89971741/85bc228a-6343-4c94-a9dc-a76b3b17181d)
  3. Create a table for computation and prepare data of Chainage, Hor.Offset, Ver.Offset or 3d-coordinate (Easting, Northing, Elevation)
  4. VBA was created as **Function** below.\
       4.1 Known Chainage, Hor.Offset, Hor.Array name compute Easting, Northing, Azimuth : **Hor(Chainage, Hor.Offset, Hor.Array name, "E or N or W")** ***Note."E" = Easting, "N" = Northing, "W" = Azimuth***\
       4.2 Known Easting, Northing, Hor.Array name compute Chainage, Hor.Offset : **HorP(Easting, Northing, Hor.Array name "C or O", 0)** ***Note."C" = Chainage, "O" = Hor.Offset***\
       4.3 Know Chainage, Ver.Array name compute Elevation : **Ver(Chainage, Ver.Array name)**\
       4.4 Horizontal and Vertical Type : HorType(Chainage, Hor.Array name), **VerType(Chainage, Ver.Array name)**\
       4.5 Convert Azimuth from Degrees to Degree-Minute-Second (DD-MM-SS) : **DegtoDMSStr2(deg)**
     ![image](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/assets/89971741/d838b509-dd08-4f2d-a657-c739228f50ed)
