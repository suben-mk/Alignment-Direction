# Alignment Direction
**Alignment direction** เป็นการคำนวณหาค่าพิกัด 3 มิติ (3D-Coordinate) ที่ระยะ Chainage และระยะ Offset ใดๆ หรือการคำนวณหาระยะ Chainage และระยะ Offset ที่ค่าพิกัด 3 มิติ (3D-Coordinate) ใดๆ จากข้อมูล Setting-out Alignment (ARRAY) สำหรับการทำงานสำรวจ เช่น วางตำแหน่งแนวอุโมงทุกๆระยะ Chainage 25 เมตร ในโครงการ, นำข้อมูลแนวอุโมงค์ทุกๆระยะ Chainage 50 เซนติเมตร ใส่ในหัวเจาะอุโมงค์ (Tunnel Boring Machine) เป็นต้น

ผู้เขียนได้ขียนโค้ดสำหรับการคำนวณ Alignment Direction ไว้ 2 ภาษา คือภาษา Python และภาษา VBA Excel

## Workflow
### Python
  **_Python libraries :_** Numpy, Pandas
  1. เตรียมข้อมูล Alignment ที่ตำแหน่งใดๆ ตาม Format [Import Alignment Direction Data.xlsx](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/Python/Import%20Data/Import%20Alignment%20Direction%20Data.xlsx)
  2. เตรียมข้อมูล Setting-out Alignment (ARRAY) ตาม Format [Export Hor-Alignment.xlsx](https://github.com/suben-mk/Alignment-Direction-for-Tunnel-Project/blob/main/Python/Import%20Data/Export%20Hor-Alignment.xlsx) และ [Export Ver-Alignment.xlsx](https://github.com/suben-mk/Alignment-Direction-for-Tunnel-Project/blob/main/Python/Import%20Data/Export%20Ver-Alignment.xlsx)
  3. ตั้งไฟล์ Path
     
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
  4. รันไฟล์ Python
### VBA
  1. เปิดไฟล์ [**VBA - Alignment Direction & Survey Program Rev.09.xlsm**](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/VBA/VBA%20-%20Alignment%20Direction%20%26%20Survey%20Program%20Rev.09.xlsm)
  2. เพิ่มข้อมูล Setting-out Alignmet (Array) sheet ไปที่ **VBA - Alignment Direction & Survey Program Rev.09.xlsm** และตั้งเป็น Array โดยไปที่ Fomulars --> Name Manager
     
     ![image](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/assets/89971741/85bc228a-6343-4c94-a9dc-a76b3b17181d)
     
  3. สร้างตารางสำหรับการคำนวณและเตรียมข้อมูล Chainage, Hor.Offset, Ver.Offset หรือพิกัด 3 มิติ (3D-Coordinate; Easting, Northing, Elevation)
  4. ผู้เขียนได้ขียนโค้ด VBA แบบ **Function** ตามชื่อด้านล่าง\
       4.1 _รู้ค่า Chainage, Hor.Offset, Hor.Array name ต้องการคำนวณ Easting, Northing, Azimuth :_ **Hor(Chainage, Hor.Offset, Hor.Array name, "E or N or W")** ***Note."E" = Easting, "N" = Northing, "W" = Azimuth***\
       4.2 _รู้ค่า Easting, Northing, Hor.Array name ต้องการคำนวณ Chainage, Hor.Offset :_ **HorP(Easting, Northing, Hor.Array name "C or O", 0)** ***Note."C" = Chainage, "O" = Hor.Offset***\
       4.3 _รู้ค่า Chainage, Ver.Array name ต้องการคำนวณ Elevation :_ **Ver(Chainage, Ver.Array name)**\
       4.4 _ต้องการหา Horizontal and Vertical Type :_ **HorType(Chainage, Hor.Array name)**, **VerType(Chainage, Ver.Array name)**\
       4.5 _ต้องการแปลง Azimuth องศา (Degrees) เป็น องศา-ลิปดา-พิลิปดา (Degree-Minute-Second; DD-MM-SS) :_ **DegtoDMSStr2(deg)**
     
     ![image](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/assets/89971741/d838b509-dd08-4f2d-a657-c739228f50ed)

## Output
### Python
  [Export Alignment-CHOS to ENZ.xlsx](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/Python/Export%20Data/Export%20Alignment-CHOS%20to%20ENZ.xlsx)\
  [Export Alignment-ENZ to CHOS.xlsx](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/Python/Export%20Data/Export%20Alignment-ENZ%20to%20CHOS.xlsx)
### VBA
  [**VBA - Alignment Direction & Survey Program Rev.09.xlsm**](https://github.com/suben-mk/Alignment-Direction-for-Metro-Line/blob/main/VBA/VBA%20-%20Alignment%20Direction%20%26%20Survey%20Program%20Rev.09.xlsm)\
 สามารถดู VBA function ที่ Main Pro.-CHOS to ENZ sheet และ Main Pro.-ENZ to CHOS sheet 
