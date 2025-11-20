# PVSD
PVSD Repository is GIS Data Reviewer.

Specifically for my friends at the Psychosis Veiled in Sacred Dreams.

## ในนี้มีอะไร
รวบรวมไพธอนสคริปต์ที่ใช้ในด้าน GIS ทั้งที่อยู่ใน GDB และเป็นเชปไฟล์

## ในนี้มีอะไร
ในโปรเจ็กต์นี้ ประกอบด้วยไพธอนสคริปต์ดังต่อไปนี้

### check_required_featureclass.py

อันนี้จะใช้หรือไม่ใช้ก็ได้ ใช้ตรวจสอบว่าใน แต่ละ gdb มีฟีเจอร์คลาสที่กำหนดหรือไม่ ถ้ามี เป็นจำนวนเท่าไหร่ 

(สร้างสคริปต์นี้ เนื่องจากพบว่าการสำเนา gdb อาจจะไม่สมบูรณ์ จึงเขียนเพื่อตรวจสอบว่ามีฟีเจอร์คลาสที่ต้องการหรือเปล่า)

### gdb_data_reviewer.py

ใช้สำหรับตรวจสอบว่า ข้อมูลที่อยู่ในฟิลด์ตามที่กำหนดในแต่ละฟีเจอร์คลาส ครบตามที่กำหนดหรือไม่ 

ข้อกำหนดต่าง ๆ ที่ตรวจสอบ จะอยู่ใน [GDB Data Standard](https://github.com/ApiratRepublic/PVSD/wiki/Logic_For_GDB_Data_Standard) 

วิธีใช้งานและตั้งค่าจะอยู่ใน [gdb_data_reviewer](https://github.com/ApiratRepublic/PVSD/wiki/gdb_data_reviewer)

สคริปต์นี้ ใช้ Arcpy หมายความว่าต้องรันใน ENV ของ ArcGIS Pro

## อภิ Need you!
หากมีคำแนะนำว่าควรตรวจสอบอะไรเพิ่ม หรือตรรกะในการตรวจสอบผิดพลาด กรุณาแจ้งโดยตรงนะจ๊ะ

## โค้ดในนี้ ดูแลปรับแต่งโดย
Apirat Rattanapaiboon ชายไทยวัยใกล้ฝั่งผู้แสนจะอบอุ่นกึ่งหัวร้อน
