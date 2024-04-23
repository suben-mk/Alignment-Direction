# Alignment Direction
Alignment direction is computed from setting-out alignment data for tunnel survey work such as tunnel segment plan, tunnel boring mechine (TBM), setting-out point at site etc. I created the coding to compute 3d-coordinate by known chainage &amp; offset and compute chainage &amp; offset by known 3d-coordinate. The coding was created 2 languages which're python and vba excel.

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
  2. 
