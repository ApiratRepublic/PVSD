########################################################################
#### สำหรับเช็คว่า GDB มี Featureclass ที่กำหนด ครบหรือไม่ ใช่หรือมั่ว ชัวร์หรือไม่ ###
#### โดยจะนับจำนวนฟีเจอร์คลาสที่ตรงกับ Pattern ที่กำหนดไว้
#### สคริปต์นี้สร้างขึ้นเพื่อช่วยตรวจสอบความสมบูรณ์ของข้อมูล GDB
#### ที่อาจจะเกิดจากการคัดลอกข้อมูลที่ไม่สมบูรณ์
#### ผู้ใช้ต้องติดตั้ง arcpy และ pandas ก่อนใช้งานสคริปต์นี้
#### สามารถรันได้ใน Python environment ที่มี arcpy (เช่น ArcGIS Pro)  
#### ผู้ใช้สามารถปรับแต่ง root_dir และ output_excel ตามต้องการ
#### รวมถึงสามารถเพิ่มหรือลด Pattern ของฟีเจอร์คลาสที่ต้องการตรวจสอบได้
#### สคริปต์นี้จะสแกนหาไฟล์ .gdb ทั้งหมดใน root_dir และตรวจสอบฟีเจอร์คลาสภายใน
#### ผลลัพธ์จะถูกบันทึกในไฟล์ Excel ที่ระบุ  
#### version: 1.00
#### ผู้เขียน: Apirat Rattanapaiboon
########################################################################

import os
import re
import arcpy
import pandas as pd

# === ตั้งค่า ===
root_dir = r"D:\A02-Projects\WarRoom\GDB"
output_excel = r"D:\A02-Projects\WarRoom\Report\check_gdb.xlsx"

# === กำหนดฟีเจอร์คลาสที่จะตรวจสอบ === #
patterns = {
    "PARCEL": re.compile(r"^PARCEL_\d{2}_\d{2}$", re.IGNORECASE),
    "PARCEL_NS3K": re.compile(r"^PARCEL_\d{2}_NS3K_\d{2}$", re.IGNORECASE),
    "ROAD": re.compile(r"^ROAD_\d{2}$", re.IGNORECASE),
    "BLOCK_FIX": re.compile(r"^BLOCK_FIX_\d{2}$", re.IGNORECASE),
    "BLOCK_PRICE": re.compile(r"^BLOCK_PRICE_\d{2}$", re.IGNORECASE),
    "BLOCK_BLUE": re.compile(r"^BLOCK_BLUE_\d{2}$", re.IGNORECASE),
    "PARCEL_REL": re.compile(r"^PARCEL_REL_\d{2}$", re.IGNORECASE),
    "NS3K_REL": re.compile(r"^NS3K_REL_\d{2}$", re.IGNORECASE),
}

# === เก็บข้อมูล === #
results = []

for dirpath, dirnames, filenames in os.walk(root_dir):
    for dirname in dirnames:
        if dirname.lower().endswith(".gdb"):
            gdb_path = os.path.join(dirpath, dirname)
            print(f"Checking: {gdb_path}")
            arcpy.env.workspace = gdb_path

            counts = {k: 0 for k in patterns.keys()}

            try:
                feature_classes = arcpy.ListFeatureClasses()
                if feature_classes:
                    for fc in feature_classes:
                        for key, pattern in patterns.items():
                            if pattern.match(fc):
                                counts[key] += 1
                else:
                    print(f"  No feature classes found in {gdb_path}")
            except Exception as e:
                print(f"  Error reading {gdb_path}: {e}")

            row = {"Full Path": gdb_path}
            row.update(counts)
            results.append(row)

# === รายงานเป็นเอกเซล ===
df = pd.DataFrame(results)
df = df[
    ["Full Path", "PARCEL", "PARCEL_NS3K", "ROAD", "BLOCK_FIX", 
     "BLOCK_PRICE", "BLOCK_BLUE", "PARCEL_REL", "NS3K_REL"]
]
os.makedirs(os.path.dirname(output_excel), exist_ok=True)
df.to_excel(output_excel, index=False)

print(f"\n✅ Done! Report saved to: {output_excel}")