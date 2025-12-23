# =============================================================================
# - ตรวจสอบรูปแบบข้อมูลของฟีเจอร์คลาสใน GDB ตามมาตรฐานที่กำหนด
# - โค้ดนี้เป็นการรีแฟคเตอร์จากสคริปต์เดิมที่ใช้ `arcpy` มาเป็นโค้ดที่ใช้ไลบรารีโอเพนซอร์ส
# - แทนที่ `arcpy` ด้วย `geopandas`, `fiona`, and `pandas`.

# - Version 2.0.0 (Open-Source Refactor)
# - release date: 2025-12-18
# - ผู้เขียน: Apirat Rattanapaiboon
# - issue: ยังมีปัญหา performance เมื่อเทียบกับ `arcpy` โดยเฉพาะกับ GDB ขนาดใหญ่
# - issue: การอ่าน GDB ด้วย `fiona` อาจมีข้อจำกัดบางอย่าง ขึ้นกับไดรเวอร์ที่ติดตั้ง
# - หมายเหตุ: โค้ดนี้ต้องการไดรเวอร์ "OpenFileGDB" ที่มาพร้อมกับ `fiona`
# - issue: ยังไม่ได้ฟิลเตอร์ geometry ที่ไม่ถูกต้องก่อนตรวจสอบทับซ้อน
# - ปรับปรุง:
#   - ยกเลิก `arcpy` dependencies.
#   - ใช้ `geopandas` และ `fiona` สำหรับการอ่าน GDB และจัดการเรขาคณิต
#   - Re-implemented `FindIdentical` using WKB comparison.
#   - ยังใช้ logic แบบเดิมสำหรับการตรวจสอบข้อมูล
# =============================================================================

import os
import re
import datetime
import uuid
from collections import defaultdict
import pandas as pd
import geopandas as gpd
import fiona
import numpy as np
from openpyxl import load_workbook

# =============================================================================
# หมายเหตุสำคัญโดยย่อนะจ๊ะ:
#
# 1.  การติดตั้ง:
#     ในการอ่าน File Geodatabase (.gdb) ไลบรารี `geopandas` และ `fiona`
#     จำเป็นต้องมีไดรเวอร์ "OpenFileGDB"
#     วิธีที่ง่ายที่สุดในการติดตั้งคือใช้ `conda`:
#
#     conda install -c conda-forge geopandas
#
# 2.  ประสิทธิภาพ:
#     สคริปต์นี้จะอ่านฟีเจอร์คลาสทั้งหมดลงในหน่วยความจำ (RAM) สำหรับการตรวจสอบ
#     (เช่น `gdf = gpd.read_file(...)`)
#     ซึ่งอาจใช้หน่วยความจำจำนวนมากหากฟีเจอร์คลาสมีขนาดใหญ่มาก
#     (แตกต่างจาก `arcpy.da.SearchCursor` ที่อ่านทีละแถว)
# =============================================================================


###############################################
#----------------- ที่ตั้งไฟล์
###############################################

ROOT_DIR = r"D:\A02-Projects\WarRoom\GDB"
REPORT_ROOT = r"D:\A02-Projects\Clinix\Report"
OVERLAP_ROOT = r"D:\A02-Projects\Clinix\Overlaping"
SUMMARY_EXCEL_PATH = os.path.join(REPORT_ROOT,"Summary_Report.xlsx")

# --------------------------------------------
#   จัดการค่าต่าง ๆ รวมทั้งฟังก์ชัน ตัวแปร ที่ใช้ร่วมกัน
# --------------------------------------------

# ค่าที่กำหนดตายตัว
ROAD_LAND_USE_DOMAIN = {"เกษตรกรรม", "ที่อยู่อาศัย", "พาณิชยกรรม",
                        "พาณิชยกรรมและที่อยู่อาศัย", "ที่อยู่อาศัยและเกษตรกรรม",
                        "อุตสาหกรรม", "ส่วนราชการ",  "พื้นที่ป่าสงวน", "พื้นที่อุทยาน"}
ROAD_STREET_TYPE_DOMAIN = {"คอนกรีต", "ลาดยาง", "หินคลุก", "ลูกรัง", "ดิน", "น้ำ", "ไม้", "ทางไม่มีสภาพ"}
ROAD_REQ_NAME_TD_CODES = {1, 2, 3, 4, 5, 6, 8, 9}
REL_TABLE_NO_DOMAIN = {1, 2, 3, 41, 42, 5, 6, 7}
REL_SUB_TABLE_NO_RANGE = range(0, 7) #คือค่า 0-6 นั่นแหละไม่ได้พิมพ์ผิด

# Duplicate detection tolerance (meters)
# Set to 0.0 for exact match, or 0.001 for near-duplicates
GEOMETRY_TOLERANCE = 0.0

def safe_to_none(val):
    """เปลี่ยน pandas NA/NaN/None เป็น None"""
    if val is None:
        return None
    if pd.isna(val):
        return None
    return val

def write_error_report(error_list, gdb_path, fc_name, check_type, oid, field, value, message):
    """
    รวบรวมข้อผิดพลาดลงใน List
    """
    error_list.append([
        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        gdb_path,
        fc_name,
        check_type,
        oid,
        field,
        value,
        message
    ])

# ค้นหา GDBs
def find_gdb_paths(root_dir):
    gdb_paths = []
    for root, dirs, _ in os.walk(root_dir):
        
        # 1. ค้นหา GDB ใน Dirs ปัจจุบัน
        found_gdbs = []
        for d in dirs:
            if d.lower().endswith(".gdb"):
                gdb_paths.append(os.path.join(root, d))
                found_gdbs.append(d)

        # 2. (สำคัญ) ลบ GDB ที่พบออกจาก Dirs เพื่อป้องกันปัญหา GDB ซ้อนกัน
        for gdb_dir in found_gdbs:
            try:
                dirs.remove(gdb_dir)
            except ValueError:
                pass        
    if not gdb_paths:        
        print(f"คำเตือน: ไม่พบ .gdb ใน {root_dir}")
    else:
        print(f"พบ {len(gdb_paths)} GDB(s) สำหรับดำเนินการต่อ")
    return gdb_paths


# *** ฟังก์ชันแปลง GDB Path ***
def get_short_gdb_path(full_gdb_path):
    try:
        parent = os.path.basename(os.path.dirname(full_gdb_path))
        grandparent = os.path.basename(os.path.dirname(os.path.dirname(full_gdb_path)))
        return f"{grandparent}{os.path.sep}{parent}"
    except Exception:
        return full_gdb_path
    
def extract_province(gdb_path_str):
    try:
        match = re.search(r"^\d+[\d_]*[-_](.*?)\\", str(gdb_path_str))
        if match:
            return match.group(1)
    except Exception:
        pass
    return "Unknown"

def categorize_featureclass(fc_name_str):
    fc = str(fc_name_str).upper()
    if re.match(r"^PARCEL_\d+_\d+$", fc):
        return "PARCEL"
    elif re.match(r"^PARCEL_\d+_NS3K_\d+$", fc):
        return "NS3K"
    elif re.match(r"^ROAD_\d+$", fc):
        return "ROAD"
    elif re.match(r"^BLOCK_FIX_\d+$", fc):
        return "BLOCK_FIX"
    elif re.match(r"^BLOCK_PRICE_\d+$", fc):
        return "BLOCK_PRICE"
    elif re.match(r"^BLOCK_BLUE_\d+$", fc):
        return "BLOCK_BLUE"
    elif re.match(r"^PARCEL_REL_\d+$", fc):
        return "PARCEL_REL"
    elif re.match(r"^NS3K_REL_\d+$", fc):
        return "NS3K_REL"
    else:
        return None

# --- Helper Functions ---

def get_fiona_schema(gdb_path, layer_name):
    """
    [ทำเพิ่ม] ฟังก์ชันช่วยอ่าน Schema โดยใช้ Fiona
    แทนที่ `arcpy.ListFields`
    """
    try:
        with fiona.open(gdb_path, layer=layer_name) as src:
            return {k.upper(): v for k, v in src.schema['properties'].items()}
    except Exception as e:
        print(f"    ⚠️ Warning: Cannot read schema for {layer_name}: {e}")
        return {}

def is_numeric_field_type(fiona_type_str):
    """
    ตรวจสอบประเภทข้อมูลจาก Fiona schema string
    แทนที่การตรวจสอบ `arcpy` types
    """
    if not fiona_type_str:
        return False
    base_type = str(fiona_type_str).lower().split(':')[0].strip()  #แก้ปัญหา "int:10"
    # e.g., 'int', 'float', 'double', 'int:10', 'float:19.11'
    # ถ้าเป็น 'str', 'string', 'date', 'datetime' จะส่งกลับค่า False
    numeric_types = {'int', 'integer', 'float', 'double', 'real', 'numeric'}
    return base_type in numeric_types

# (Helper functions คงเดิม)
def can_be_number(val):
    """
    Enhanced to handle numpy scalar types
    Check if value can be converted to number
    """
    if val is None:
        return False
    if pd.isna(val):
        return False
    if isinstance(val, (int, float, np.integer, np.floating)):
        # Check for NaN/Inf in numpy types
        if isinstance(val, (float, np.floating)):
            if not np.isfinite(val):
                return False
        return True
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False
        
def safe_value_is_int_like(val):
    if val is None: return False
    try:
        if isinstance(val, (int, float, np.integer, np.floating)):
            # Check for NaN/Inf
            if not np.isfinite(val):
                return False
            return float(val).is_integer()
        if isinstance(val, str) and val.isdigit():
            return True
        return False
    except Exception:
        return False

########################################
# ฟังก์ชันตรวจสอบทับซ้อน
########################################

def check_for_exact_overlaps(gdb_path,
                             fc_name,
                             error_list,
                             output_dir,
                             output_basename,
                             return_layer_path=False,
                             verbose=True,
                             tolerance=0.0):
    """
    ตรวจสอบโพลีกอนที่ทับกันสนิท (exact overlap)
    โดยใช้ Geopandas และการเปรียบเทียบ Well-Known Binary (WKB)
    แทนที่ `arcpy.management.FindIdentical`
    
    Parameters
    ----------
    gdb_path : str
        Full path ของ GDB
    fc_name : str
        ชื่อของ feature class
    error_list : list
        รายการ error ที่จะถูกเขียนเพิ่มผ่าน write_error_report()
    output_dir : str
        โฟลเดอร์สำหรับเก็บ shapefile ที่เป็นผลลัพธ์
    output_basename : str
        ชื่อ prefix สำหรับไฟล์ผลลัพธ์
    ... (other params) ...

    Returns
    -------
    str | None
        คืน path ของ shapefile ที่พบโพลีกอนซ้ำ ถ้ามี
    """

    if verbose:
        print(f"    ▶ ตรวจสอบการซ้อนทับ (Exact Overlap): {fc_name}")
    
    output_shp = os.path.join(output_dir, f"{output_basename}_{fc_name}_duplicates.shp")

    try:
        # 1. อ่าน Feature Class ด้วย Geopandas
        # `gdb_path` คือไดเรกทอรี .gdb, `layer` คือชื่อ FC
        try:
            gdf = gpd.read_file(gdb_path, layer=fc_name)
        except Exception as e:
            msg = f"ไม่สามารถอ่าน Feature Class {fc_name} ด้วย Geopandas/Fiona: {e}"
            if verbose: print(f"      ❌ {msg}")
            write_error_report(error_list, gdb_path, fc_name, "Geometry Error", -1, "Shape", "", msg)
            return None

        if gdf.empty:
            if verbose: print("      ✓ Feature Class ว่างเปล่า, ข้ามการตรวจสอบทับซ้อน")
            return None
        

        # 2. ตรวจสอบเรขาคณิตที่ถูกต้อง
        # `FindIdentical` ทำงานกับเรขาคณิตที่ถูกต้องเท่านั้น
        # เราจะกรองเอาเฉพาะเรขาคณิตที่ถูกต้องและไม่ว่างเปล่า
        valid_gdf = gdf[gdf.geometry.is_valid & ~gdf.geometry.is_empty]
        if len(valid_gdf) != len(gdf):
            invalid_count = len(gdf) - len(valid_gdf)
            if verbose: print(f"      ⚠ พบ {invalid_count} invalid/empty geometries (จะถูกข้าม)")
            # หมายเหตุ: เราสามารถรายงาน FIDs/OIDs ที่ไม่ถูกต้องได้ถ้าต้องการ

        # 3. แปลงเรขาคณิตเป็น Well-Known Binary (WKB)
        # คิดว่านี่เป็นวิธีที่รวดเร็วในการหา "duplicate" geometries
        if tolerance == 0.0:
            wkb_series = valid_gdf.geometry.to_wkb()
            is_duplicate = wkb_series.duplicated(keep=False)
        else:
            # Tolerance-based comparison
            is_duplicate = pd.Series([False] * len(valid_gdf), index=valid_gdf.index)
            
            for i in range(len(valid_gdf)):
                if is_duplicate.iloc[i]:
                    continue
                
                geom_i = valid_gdf.geometry.iloc[i]
                for j in range(i + 1, len(valid_gdf)):
                    if is_duplicate.iloc[j]:
                        continue
                    
                    geom_j = valid_gdf.geometry.iloc[j]
                    
                    try:
                        if geom_i.equals_exact(geom_j, tolerance=tolerance):
                            is_duplicate.iloc[i] = True
                            is_duplicate.iloc[j] = True
                    except Exception:
                        # Fallback to exact WKB comparison
                        if geom_i.wkb == geom_j.wkb:
                            is_duplicate.iloc[i] = True
                            is_duplicate.iloc[j] = True
    
        # 4. ค้นหา WKB ที่ซ้ำกัน
        # `duplicated(keep=False)` จะทำเครื่องหมาย *ทุกแถว* ที่มีค่าซ้ำ
        is_duplicate = wkb_series.duplicated(keep=False)
        
        # 5. กรอง GeoDataFrame ให้เหลือเฉพาะแถวที่ซ้ำ
        dup_gdf = valid_gdf[is_duplicate]

        if dup_gdf.empty:
            if verbose: print("      ✓ ไม่พบโพลีกอนทับกันสนิท")
            return None
        
        # 6. รวบรวม FIDs/OIDs ที่ซ้ำ
        # ใน `geopandas`, FID (Object ID) คือ `gdf.index`
        dup_fids = sorted(dup_gdf.index.tolist())
        count = len(dup_fids)
        
        msg = f"พบโพลีกอนทับกันสนิท {count} รูปแปลง (FIDs: {dup_fids[:20]}{'...' if count > 20 else ''})"
        if verbose: print(f"      ⚠ {msg}")

        write_error_report(
            error_list,
            gdb_path,
            fc_name,
            "Duplicated Polygon",
            str(dup_fids),
            "Shape",
            count,
            msg
        )

        # 7. สร้าง Shapefile Output
        os.makedirs(output_dir, exist_ok=True)
        # บันทึกเฉพาะแถวที่ซ้ำ
        dup_gdf.to_file(output_shp, driver='ESRI Shapefile')

        if verbose: print(f"      → บันทึก shapefile: {output_shp}")
        return output_shp if return_layer_path else None

    except Exception as e:
        # `fiona` อาจมีปัญหาในการเขียน shapefile ถ้า schema ซับซ้อน
        msg = f"เกิดข้อผิดพลาดระหว่างตรวจสอบทับซ้อน (Geopandas): {e}"
        if verbose: print(f"      ❌ {msg}")
        write_error_report(error_list, gdb_path, fc_name, "Geometry Error", -1, "Shape", "", msg)
        return None
    finally:
        if verbose: print("      • ตรวจสอบทับซ้อนเสร็จสิ้น\n")
        # Explicit cleanup
        try:
            if 'gdf' in locals():
                del gdf
            if 'valid_gdf' in locals():
                del valid_gdf
            if 'dup_gdf' in locals():
                del dup_gdf
        except Exception:
            pass        


# ----------------------------------------
# [REFACTORED] ตรวจสอบประเภทข้อมูลและค่าต่าง ๆ
# ----------------------------------------

# --- [HELPER] สำหรับอ่านข้อมูล (จัดการ Table vs Feature Class) ---
def read_layer_data(gdb_path, fc_name, is_spatial):
    """
    อ่านข้อมูลจาก GDB layer
    ถ้า `is_spatial`=True, อ่านด้วย Geopandas (ได้ geometry)
    ถ้า `is_spatial`=False, อ่านด้วย Fiona (เร็วกว่าสำหรับ table) และแปลงเป็น DataFrame
    """
    try:
        if is_spatial:
            return gpd.read_file(gdb_path, layer=fc_name)
        else:
            records = []
            with fiona.open(gdb_path, layer=fc_name) as src:
                for feat in src:
                    rec = feat['properties']
                    rec['OID@'] = feat.get('id', -1)
                    records.append(rec)
            
            if not records:
                return pd.DataFrame(columns=['OID@'])
            
            df = pd.DataFrame(records)
            df = df.set_index('OID@')
            return df
    except Exception as e:
        print(f"    ❌ Error reading layer {fc_name}: {e}")
        raise

################################################
#--------------------- 1) PARCEL
################################################

def validate_parcel(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `gpd.read_file().iterrows()` """
    
    print(f"  กำลังตรวจสอบ PARCEL: {fc_name}")
    
    # [MODIFIED] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return

    required = ["UTMMAP1","UTMMAP2","UTMMAP3","UTMMAP4","UTMSCALE","LAND_NO",
                "PARCEL_TYPE","CHANGWAT_CODE","BRANCH_CODE","PARCEL_RN"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์นี้")
# ----------------------------------------    
    # [แก้ไข] ตรวจสอบประเภทข้อมูล
# ----------------------------------------
    if "UTMMAP1" in schema and schema["UTMMAP1"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP1", schema["UTMMAP1"], "ต้องเป็น String")
    if "UTMMAP2" in schema and not is_numeric_field_type(schema["UTMMAP2"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP2", schema["UTMMAP2"], "ต้องเป็น Number")
    if "UTMMAP3" in schema and schema["UTMMAP3"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP3", schema["UTMMAP3"], "ต้องเป็น String")
    if "UTMMAP4" in schema and schema["UTMMAP4"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP4", schema["UTMMAP4"], "ต้องเป็น String")
    if "UTMSCALE" in schema and not is_numeric_field_type(schema["UTMSCALE"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMSCALE", schema["UTMSCALE"], "ต้องเป็น Number")
    if "LAND_NO" in schema and not is_numeric_field_type(schema["LAND_NO"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "LAND_NO", schema["LAND_NO"], "ต้องเป็น Number") 
    if "PARCEL_TYPE" in schema and not is_numeric_field_type(schema["PARCEL_TYPE"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "PARCEL_TYPE", schema["PARCEL_TYPE"], "ต้องเป็น Number")       
    if "CHANGWAT_CODE" in schema and schema["CHANGWAT_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "CHANGWAT_CODE", schema["CHANGWAT_CODE"], "ต้องเป็น String")
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")
 
    # ----------------------------------------
    # [แก้ไข] ตรวจสอบ ความถูกต้องของข้อมูล
    # ----------------------------------------
    utm_key = defaultdict(list)
    branch_parcel_rn = defaultdict(list)

    try:
        # [แก้ไข] อ่านข้อมูลด้วย Geopandas
        gdf = read_layer_data(gdb_path, fc_name, is_spatial=True)

        # [แก้ไข] วนลูปด้วย `iterrows()`
        # `oid` คือ FID/Object ID (จาก GDF index)
        # `row` คือ Pandas Series (เหมือน dict)
        for oid, row in gdf.iterrows():
            
            # [เพิ่มใหม่] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            # เพื่อให้โค้ดตรวจสอบเดิมทำงานได้
            rec = {k.upper(): v for k, v in row.items()}
            # `oid` จาก `iterrows()` คือ FID/OID อยู่แล้ว
            rec["OID@"] = oid 

            # ================================================================
            # --- [ของเดิม] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            utm1 = safe_to_none(rec.get("UTMMAP1"));utm2 = safe_to_none(rec.get("UTMMAP2")); utm3=safe_to_none(rec.get("UTMMAP3")); utm4=safe_to_none(rec.get("UTMMAP4"))
            scale = safe_to_none(rec.get("UTMSCALE")); land_no=safe_to_none(rec.get("LAND_NO")); parcel_type=safe_to_none(rec.get("PARCEL_TYPE"))
            cwt = safe_to_none(rec.get("CHANGWAT_CODE")); branch = safe_to_none(rec.get("BRANCH_CODE")); parcel_rn = safe_to_none(rec.get("PARCEL_RN"))

            # 1.1.1. UTMMAP1 ต้องเป็น String และเป็น 4 หลักเท่านั้น เช่น "5042"
            if not (isinstance(utm1, str) and utm1.isdigit() and len(utm1)==4):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP1", utm1, "UTMMAP1 ต้องเป็น 4 หลัก")
            
            # 1.1.2.UTMMAP2	ต้องเป็น  Number  และต้องเป็น 1 หรือ 2 หรือ 3 หรือ 4 เท่านั้น 
            if not (isinstance(utm2,(int,float, np.integer, np.floating)) or (isinstance(utm2,str) and utm2.isdigit())):
                write_error_report(error_list, gdb_path, fc_name, "Field Type", oid, "UTMMAP2", utm2, "ประเภทข้อมูลต้องเป็น Number และไม่ควรว่าง")
            else:
                try:
                    if int(float(utm2)) not in (1,2,3,4):
                        write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP2", utm2, "UTMMAP2 ต้องเป็น 1 - 4 ")
                except:
                    pass
            
            # 1.1.3.UTMMAP3	ต้องเป็น String  และเป็น 4 หลักเท่านั้น เช่น "0016"
            if not (isinstance(utm3,str) and utm3.isdigit() and len(utm3)==4):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP3", utm3, "UTMMAP3 ต้องเป็น 4 หลัก")
            
            # 1.1.4.UTMMAP4	ต้องเป็น String  และเป็น 2 หลักเท่านั้น เช่น "02"
            if not (isinstance(utm4,str) and utm4.isdigit() and len(utm4)==2):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP4", utm4, "UTMMAP4 ของชั้น PARCEL ต้องเป็น 2 หลัก")
            else:
                try:
                    scale_i = int(float(scale)) if scale is not None and can_be_number(scale) else None
                except:
                    scale_i = None
                if scale_i == 4000 and utm4 != '00':
                    write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMMAP4", utm4, "UTMMAP4 ต้องเป็น '00' เนื่องจาก UTMSCALE=4000")
                elif scale_i == 2000:
                    try:
                        if not (1 <= int(utm4) <= 4):
                            write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMMAP4", utm4, "UTMMAP4 ต้องอยู่ระหว่าง '01'-'04' เนื่องจาก UTMSCALE=2000")
                    except:
                        pass
                elif scale_i == 1000:
                    try:
                        if not (1 <= int(utm4) <= 16):
                            write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMMAP4", utm4, "UTMMAP4 ต้องอยู่ระหว่าง '01'-'16' เนื่องจาก UTMSCALE=1000")
                    except:
                        pass
                elif scale_i == 500:
                    try:
                        if not (1 <= int(utm4) <= 64):
                            write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMMAP4", utm4, "UTMMAP4 ต้องอยู่ระหว่าง '01'-'64' เนื่องจาก UTMSCALE=500")
                    except:
                        pass

            # 1.1.5.UTMSCALE  ต้องเป็น Number และเป็น  4000 หรือ 2000 หรือ 1000 หรือ 500 เท่านั้น
            if scale is None or not can_be_number(scale) or int(float(scale)) not in (4000,2000,1000,500):
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMSCALE", scale, "UTMSCALE ของฟีเจอร์คลาส PARCEL จะต้องเป็น 4000,2000,1000 หรือ 500")

            # 1.1.8.CHANGWAT_CODE ต้องเป็น String และเป็น 2 หลัก เช่น "66"
            if not (isinstance(cwt,str) and len(cwt)==2 and cwt.isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "CHANGWAT_CODE", cwt, "CHANGWAT_CODE ต้องเป็น 2 หลัก")

            # 1.1.9.BRANCH_CODE ต้องเป็น String และเป็น 8 หลัก และสองหลักแรก จะต้องตรงกับ CHANGWAT_CODE
            if not (isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")
            else:
                if isinstance(cwt,str) and not branch.startswith(cwt):
                    write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "BRANCH_CODE", branch, f"2 หลักแรกของ BRANCH_CODE ไม่ตรงกับ CHANGWAT_CODE {cwt}")
            
            # 1.1.10.PARCEL_RN ต้องเป็น Number และใน BRANCH_CODE เดียวกัน จะต้องไม่มีค่าซ้ำ
            if parcel_rn is None or not can_be_number(parcel_rn):
                write_error_report(error_list, gdb_path, fc_name, "Field Type", oid, "PARCEL_RN", parcel_rn, "ต้องเป็น Number และไม่ควรว่าง")
            else:
                branch_parcel_rn[(branch.strip() if branch else "NULL", int(float(parcel_rn)))].append(oid)

            # 1.2. ถ้า LAND_NO ไม่ใช่ค่าว่าง หรือ 0 
            is_land_no_valid = False
            if land_no is not None and can_be_number(land_no):
                if int(float(land_no)) != 0:
                    is_land_no_valid = True

            if is_land_no_valid:
                scale_i = int(float(scale)) if can_be_number(scale) else scale
                check_key = (branch.strip() if branch else "NULL", utm1, utm2, utm3, utm4, scale_i, land_no)
                utm_key[check_key].append(oid)
            # ================================================================
            # --- สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================

        # Logic สำหรับตรวจสอบค่าซ้ำ
        for primery_key, oids in utm_key.items():
            if len(oids) > 1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate UTM", str(oids), "PRIMERY_KEY", primery_key, "BRANCH_CODE+UTMMAP1+UTMMAP2+UTMMAP3+UTMMAP4+UTMSCALE+LAND_NO มีค่าซ้ำ")

        for k, oids in branch_parcel_rn.items():
            if len(oids) > 1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "PARCEL_RN", k, "PARCEL_RN มีค่าซ้ำภายใน BRANCH_CODE เดียวกัน")
    
    except Exception as ex:
        # Catch errors during file read or iteration
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Geopandas read/loop) {ex}")
    
    # 1.3. [ปรับแต่ง] ตรวจสอบโพลีกอนที่ซ้อนทับกันสนิท
    check_for_exact_overlaps(gdb_path, fc_name, error_list, os.path.join(OVERLAP_ROOT,"PARCEL"), basename or "PARCEL")


################################################
#----------------2) PARCEL_NS3K
################################################

def validate_parcel_ns3k(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `gpd.read_file().iterrows()` """
    
    print(f"  กำลังตรวจสอบ PARCEL_NS3K: {fc_name}")

    # [แก้ไข] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return

    required = ["UTMMAP1","UTMMAP2","UTMMAP3","UTMMAP4","UTMSCALE","LAND_NO","PARCEL_TYPE","CHANGWAT_CODE","BRANCH_CODE","NS3K_RN"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์นี้")

    # [แก้ไข] ตรวจสอบประเภทข้อมูล
    if "UTMMAP1" in schema and schema["UTMMAP1"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP1", schema["UTMMAP1"], "ต้องเป็น String")
    if "UTMMAP2" in schema and not is_numeric_field_type(schema["UTMMAP2"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP2", schema["UTMMAP2"], "ต้องเป็น Number")
    if "UTMMAP3" in schema and schema["UTMMAP3"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP3", schema["UTMMAP3"], "ต้องเป็น String")
    if "UTMMAP4" in schema and schema["UTMMAP4"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMMAP4", schema["UTMMAP4"], "ต้องเป็น String")
    if "UTMSCALE" in schema and not is_numeric_field_type(schema["UTMSCALE"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "UTMSCALE", schema["UTMSCALE"], "ต้องเป็น Number")
    if "LAND_NO" in schema and not is_numeric_field_type(schema["LAND_NO"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "LAND_NO", schema["LAND_NO"], "ต้องเป็น Number")
    if "PARCEL_TYPE" in schema and not is_numeric_field_type(schema["PARCEL_TYPE"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "PARCEL_TYPE", schema["PARCEL_TYPE"], "ต้องเป็น Number")
    if "CHANGWAT_CODE" in schema and schema["CHANGWAT_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "CHANGWAT_CODE", schema["CHANGWAT_CODE"], "ต้องเป็น String")
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")
    
    # ----------------------------------------
    # [ปรับแต่ง] ตรวจสอบ ความถูกต้องของข้อมูล
    # ----------------------------------------
    utm_key = defaultdict(list)
    branch_ns3k = defaultdict(list)

    try:
        # [แก้ไข] อ่านข้อมูลด้วย Geopandas
        gdf = read_layer_data(gdb_path, fc_name, is_spatial=True)

        # [แก้ไข] วนลูปด้วย `iterrows()`
        for oid, row in gdf.iterrows():
            # [เพิ่มใหม่] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 
            
            # ================================================================
            # --- [ของเดิม] โค้ดตรวจสอบเดิม ---
            # ================================================================
            utm1 = safe_to_none(rec.get("UTMMAP1")); utm2=safe_to_none(rec.get("UTMMAP2")); utm3=safe_to_none(rec.get("UTMMAP3")); utm4=safe_to_none(rec.get("UTMMAP4"))
            scale = safe_to_none(rec.get("UTMSCALE")); land_no = safe_to_none(rec.get("LAND_NO")); parcel_type = safe_to_none(rec.get("PARCEL_TYPE"))
            cwt = safe_to_none(rec.get("CHANGWAT_CODE")); branch = safe_to_none(rec.get("BRANCH_CODE")); ns3k_rn = safe_to_none(rec.get("NS3K_RN"))
            
            # 2.1.1. UTMMAP1 ต้องเป็น String และเป็น 4 หลักเท่านั้น เช่น "5042"
            if not (isinstance(utm1,str) and utm1.isdigit() and len(utm1)==4):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP1", utm1, "UTMMAP1 ต้องมี 4 หลัก")
            
            # 2.1.2. UTMMAP2 ต้องเป็น  Number  และต้องเป็น 1 หรือ 2 หรือ 3 หรือ 4 เท่านั้น
            if not (isinstance(utm2,(int,float, np.integer, np.floating)) or (isinstance(utm2,str) and utm2.isdigit())):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP2", utm2, "รูปแบบข้อมูลต้องเป็น Number")
            else:
                try:
                    if int(float(utm2)) not in (1,2,3,4):
                        write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP2", utm2, "UTMMAP2 ต้องอยู่ระหว่าง 1-4")
                except:
                    pass
            
            # 2.1.3. UTMMAP3 ต้องเป็น String และต้องเป็น '0000' เท่านั้น
            if not (isinstance(utm3,str) and utm3 == "0000"):
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMMAP3", utm3, "UTMMAP3 ของ NS3K ต้องเป็น '0000'")
            
            # 2.1.4. UTMMAP4  ต้องเป็น String และเป็น 3 หลักเท่านั้น เช่น "002"
            if not (isinstance(utm4,str) and utm4.isdigit() and len(utm4)==3):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "UTMMAP4", utm4, "ต้องเป็น 3 หลัก")
            
            # 2.1.5. UTMSCALE  ต้องเป็น Number และต้องเป็น 5000 เท่านั้น
            if scale is None or not can_be_number(scale) or int(float(scale)) != 5000:
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "UTMSCALE", scale, "UTMSCALE ของ NS3K ต้องเป็น 5000")
            
            # 2.1.7. PARCEL_TYPE  ต้องเป็น Number และต้องเป็น 3 เท่านั้น
            if not can_be_number(parcel_type) or int(float(parcel_type)) != 3:
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "PARCEL_TYPE", parcel_type, "PARCEL_TYPE ของ NS3K ต้องเป็น 3")
            
            # 2.1.8. CHANGWAT_CODE ต้องเป็น String และเป็น 2 หลัก เช่น "66"
            if not (isinstance(cwt,str) and len(cwt)==2 and cwt.isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "CHANGWAT_CODE", cwt, "ต้องเป็น 2 หลัก")
            
            # 2.1.9. BRANCH_CODE  ต้องเป็น String และเป็น 8 หลัก และสองหลักแรก จะต้องตรงกับ CHANGWAT_CODE
            if not (isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "ต้องเป็น 8 หลัก")
            else:
                if isinstance(cwt,str) and not branch.startswith(cwt):
                    write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "BRANCH_CODE", branch, f"2 หลักแรกของ BRANCH_CODE ไม่ตรงกับ CHANGWAT_CODE {cwt}")
            
            # 2.1.10. NS3K_RN ต้องเป็น Number และใน BRANCH_CODE เดียวกัน จะต้องไม่มีค่าซ้ำ
            if ns3k_rn is None or not can_be_number(ns3k_rn):
                write_error_report(error_list, gdb_path, fc_name, "Field Type", oid, "NS3K_RN", ns3k_rn, "ต้องเป็น Number")
            else:
                branch_ns3k[(branch.strip() if branch else "NULL", int(float(ns3k_rn)))].append(oid)
            
            # 2.2. ถ้า LAND_NO ไม่ใช่ค่าว่าง หรือ 0 
            is_land_no_valid = False
            if land_no is not None and can_be_number(land_no):
                if int(float(land_no)) != 0:
                    is_land_no_valid = True

            if is_land_no_valid:
                scale_i = int(float(scale)) if can_be_number(scale) else scale
                check_key = (branch.strip() if branch else "NULL", utm1, utm2, utm3, utm4, scale_i, land_no)
                utm_key[check_key].append(oid)
            # ================================================================
            # --- สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================

        # (ของเดิม) Logic ตรวจสอบค่าซ้ำ
        for primery_key,oids in utm_key.items():
            if len(oids)>1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate UTM", str(oids), "PRIMERY_KEY", primery_key, "BRANCH_CODE+UTMMAP1+UTMMAP2+UTMMAP3+UTMMAP4+UTMSCALE+LAND_NO มีค่าซ้ำ")
        for k,oids in branch_ns3k.items():
            if len(oids)>1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "NS3K_RN", k, "NS3K_RN ซ้ำภายใน BRANCH_CODE เดียวกัน")
    
    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Geopandas read/loop) {ex}")

    # 2.3. [ปรับแต่ง] ตรวจสอบโพลีกอนที่ซ้อนทับกันสนิท
    check_for_exact_overlaps(gdb_path, fc_name, error_list, os.path.join(OVERLAP_ROOT,"PARCEL"), basename or "PARCEL_NS3K")

################################################
# -------------------3) ROAD
################################################

def validate_road(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `gpd.read_file().iterrows()` """
    
    print(f"  กำลังตรวจสอบชั้นข้อมูล ROAD: {fc_name}")
    
    # [แก้ไข] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return

    VALID_LAND_USE = ROAD_LAND_USE_DOMAIN
    VALID_STREET_TYPE = ROAD_STREET_TYPE_DOMAIN
    VALID_TD_RP3 = ROAD_REQ_NAME_TD_CODES

    required = ["STREET_NAME","STREET_CODE","STREET_DEPTH",
                "LAND_USE","STREET_TYPE","STREET_WIDTH","STREET_AREA",
                "BRANCH_CODE","PARCEL_TYPE",
                "TD_RP3_TYPE_CODE","STREET_RN",
                "CHANGWAT_CODE","STREET_SMG"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์นี้")
    
    # ----------------------------------------
    # [แก้ไข] ตรวจสอบประเภทข้อมูล
    # ----------------------------------------
    if "STREET_NAME" in schema and schema["STREET_NAME"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_NAME", schema["STREET_NAME"], "ต้องเป็น String")
    if "STREET_CODE" in schema and schema["STREET_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_CODE", schema["STREET_CODE"], "ต้องเป็น String")
    if "STREET_DEPTH" in schema and not is_numeric_field_type(schema["STREET_DEPTH"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_DEPTH", schema["STREET_DEPTH"], "ต้องเป็น Number")
    if "LAND_USE" in schema and schema["LAND_USE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "LAND_USE", schema["LAND_USE"], "ต้องเป็น String")
    if "STREET_TYPE" in schema and schema["STREET_TYPE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_TYPE", schema["STREET_TYPE"], "ต้องเป็น String")
    if "STREET_WIDTH" in schema and not is_numeric_field_type(schema["STREET_WIDTH"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_WIDTH", schema["STREET_WIDTH"], "ต้องเป็น Number")
    if "STREET_AREA" in schema and not is_numeric_field_type(schema["STREET_AREA"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_AREA", schema["STREET_AREA"], "ต้องเป็น Number")
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")
    if "PARCEL_TYPE" in schema and not is_numeric_field_type(schema["PARCEL_TYPE"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "PARCEL_TYPE", schema["PARCEL_TYPE"], "ต้องเป็น Number")
    if "TD_RP3_TYPE_CODE" in schema and not is_numeric_field_type(schema["TD_RP3_TYPE_CODE"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "TD_RP3_TYPE_CODE", schema["TD_RP3_TYPE_CODE"], "ต้องเป็น Number")
    if "STREET_RN" in schema and not is_numeric_field_type(schema["STREET_RN"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_RN", schema["STREET_RN"], "ต้องเป็น Number")
    if "CHANGWAT_CODE" in schema and schema["CHANGWAT_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "CHANGWAT_CODE", schema["CHANGWAT_CODE"], "ต้องเป็น String")
    if "STREET_SMG" in schema and schema["STREET_SMG"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_SMG", schema["STREET_SMG"], "ต้องเป็น String")

    try:
        branch_street_rn_seen = defaultdict(list)
        name_code_branch_list = []

        # [แก้ไข] อ่านข้อมูลด้วย Geopandas
        gdf = read_layer_data(gdb_path, fc_name, is_spatial=True)

        # [แก้ไข] วนลูปด้วย `iterrows()`
        for oid, row in gdf.iterrows():
            # [เพิ่มใหม่] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 

            # ================================================================
            # --- [ของเดิม] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            branch = safe_to_none(rec.get("BRANCH_CODE")); cwt = safe_to_none(rec.get("CHANGWAT_CODE"))
            parcel_type = safe_to_none(rec.get("PARCEL_TYPE")); td_type = safe_to_none(rec.get("TD_RP3_TYPE_CODE"))
            street_rn = safe_to_none(rec.get("STREET_RN")); name = safe_to_none(rec.get("STREET_NAME")); code = safe_to_none(rec.get("STREET_CODE"))
            street_type = safe_to_none(rec.get("STREET_TYPE")); land_use = safe_to_none(rec.get("LAND_USE"))

            # (มีไว้จัดการค่า None จาก pandas/numpy)
            if pd.isna(name): name = None
            if pd.isna(code): code = None
            if pd.isna(branch): branch = None
            if pd.isna(cwt): cwt = None
            if pd.isna(street_type): street_type = None
            if pd.isna(land_use): land_use = None
            if pd.isna(td_type): td_type = None
            if pd.isna(street_rn): street_rn = None

            # ================================================================
            # ถ้า TD_RP3_TYPE_CODE เป็น 9 ไม่ต้องตรวจสอบใดใดทั้งสิ้น
            
            td_type_int_check = None
            if can_be_number(td_type):
                td_type_int_check = int(float(td_type))
            
            if td_type_int_check == 9:
                continue  # ข้ามการตรวจสอบทั้งหมดสำหรับแถวนี้

            # 3.1.1. STREET_NAME ต้องเป็น String ถ้า TD_RP3_TYPE_CODE เป็น 1 หรือ 2 หรือ 3 หรือ 4 หรือ 5 หรือ 6 หรือ 8 จะต้องไม่ใช่ค่าว่าง
            is_street_name_empty = (name is None or (isinstance(name, str) and name.strip() == ""))

            # 3.1.4. LAND_USE ต้องเป็น String
            if (not is_street_name_empty) and (land_use is None or str(land_use).strip() not in VALID_LAND_USE):
                write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "LAND_USE", land_use, f"LAND_USE จะต้องมีค่าดังต่อไปนี้ {VALID_LAND_USE} (เมื่อ STREET_NAME มีค่า)")
            
            # 3.1.5. ตรวจสอบ STREET_TYPE
            if (not is_street_name_empty) and (street_type is None or str(street_type).strip() not in VALID_STREET_TYPE):
                write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "STREET_TYPE", street_type, f"STREET_TYPE จะต้องมีค่าดังต่อไปนี้ {VALID_STREET_TYPE} (เมื่อ STREET_NAME มีค่า)")

            # 3.1.8. CHANGWAT_CODE ต้องเป็น String และเป็น 2 หลัก เช่น "66"
            if not (isinstance(cwt,str) and len(cwt)==2 and cwt.isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "CHANGWAT_CODE", cwt, "ต้องเป็น 2 หลัก")
            
            # 3.1.9. BRANCH_CODE ต้องเป็น String และเป็น 8 หลัก และสองหลักแรก จะต้องตรงกับ CHANGWAT_CODE
            if branch and cwt and (not (isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit())):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")
            if branch and cwt and isinstance(branch,str) and isinstance(cwt,str) and not branch.startswith(cwt):
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "BRANCH_CODE", branch, f"2 หลักแรกของ BRANCH_CODE ไม่ตรงกับ CHANGWAT_CODE {cwt}")

            # 3.1.11. TD_RP3_TYPE_CODE ต้องเป็น Number
            td_type_int = None
            is_td_type_valid_number = False

            if td_type is None:
                td_type_int = None
                is_td_type_valid_number = True
            elif can_be_number(td_type):
                td_type_int = int(float(td_type))
                is_td_type_valid_number = True
            else:
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "TD_RP3_TYPE_CODE", td_type, "TD_RP3_TYPE_CODE ต้องเป็นตัวเลขเท่านั้น")
                is_td_type_valid_number = False
            
            if is_td_type_valid_number:
                is_street_name_empty = (name is None or (isinstance(name, str) and name.strip() == ""))

                if not is_street_name_empty:
                    if td_type_int not in VALID_TD_RP3:
                        write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "TD_RP3_TYPE_CODE", td_type, f"TD_RP3_TYPE_CODE ต้องมีค่าเป็น {sorted(VALID_TD_RP3)} (เนื่องจาก STREET_NAME มีค่า)")
                else:
                    ALLOWED_VALUES_WHEN_NAME_IS_EMPTY = {0, None} 
                    ALLOWED_VALUES_WHEN_NAME_IS_EMPTY.update(VALID_TD_RP3)
                    if td_type_int not in ALLOWED_VALUES_WHEN_NAME_IS_EMPTY:
                        allowed_str = "{0, None} หรือ " + str(sorted(VALID_TD_RP3))
                        write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "TD_RP3_TYPE_CODE", td_type, f"TD_RP3_TYPE_CODE ต้องเป็น {allowed_str} (เนื่องจาก STREET_NAME ว่างเปล่า)")
            
            if (td_type_int in VALID_TD_RP3) and is_street_name_empty:
                write_error_report(
                    error_list, 
                    gdb_path, 
                    fc_name, 
                    "Data Required",
                    oid, 
                    "STREET_NAME", 
                    name, 
                    f"STREET_NAME ต้องไม่เป็นค่าว่าง เนื่องจาก TD_RP3_TYPE_CODE คือ {td_type_int}"
                )

            if (td_type_int == 0 or td_type_int is None) and not is_street_name_empty:
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "STREET_NAME", name, "TD_RP3_TYPE_CODE เป็น 0 หรือ NULL ดังนั้น STREET_NAME ต้องเป็นค่าว่าง")
            
            if is_street_name_empty and (land_use is not None and str(land_use).strip() != ""):
                write_error_report(error_list, gdb_path, fc_name, "Conditional Rule", oid, "LAND_USE", land_use, "STREET_NAME เป็นค่าว่าง ดังนั้น LAND_USE ต้องเป็นค่าว่างด้วย")

            # 3.1.12. STREET_RN ต้องเป็น Number และใน BRANCH_CODE เดียวกัน จะต้องไม่มีค่าซ้ำ
            if street_rn is None or not can_be_number(street_rn):
                    write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "STREET_RN", street_rn, "STREET_RN ต้องเป็น Number")
            else:
                key = (branch.strip() if isinstance(branch,str) else "NULL", int(float(street_rn)))
                branch_street_rn_seen[key].append(oid)

            if name and code:
                name_code_branch_list.append((branch, name, code))
            # ================================================================
            # --- [KEPT] สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================
    
        # (KEPT) Logic for 1-to-1 checks
        name_to_code_by_branch = defaultdict(dict)
        code_to_name_by_branch = defaultdict(dict)

        for branch, name, code in name_code_branch_list:
            branch_key = branch.strip() if branch else "NULL" 
            
            if name in name_to_code_by_branch[branch_key] and name_to_code_by_branch[branch_key][name] != code:
                write_error_report(error_list, gdb_path, fc_name, "OneToOne", "N/A", "STREET_NAME", name, f"{name} เชื่อมต่อกับ STREET_CODE มากกว่า 1 ค่า (ภายใน BRANCH_CODE '{branch_key}')")
            
            if code in code_to_name_by_branch[branch_key] and code_to_name_by_branch[branch_key][code] != name:
                write_error_report(error_list, gdb_path, fc_name, "OneToOne", "N/A", "STREET_CODE", code, f"{code} เชื่อมต่อกับ STREET_NAME มากกว่า 1 ค่า (ภายใน BRANCH_CODE '{branch_key}')")
            
            name_to_code_by_branch[branch_key][name] = code
            code_to_name_by_branch[branch_key][code] = name

        for key, oids in branch_street_rn_seen.items():
            if len(oids) > 1:
                branch_str, rn_str = key
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "STREET_RN", rn_str, f"STREET_RN ซ้ำ ภายใน BRANCH_CODE '{branch_str}'")
    
    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Geopandas read/loop) {ex}")
    
    #----- 3.3. [MODIFIED] ตรวจสอบโพลีกอนที่ซ้อนทับกันสนิท
    check_for_exact_overlaps(gdb_path, fc_name, error_list, os.path.join(OVERLAP_ROOT,"ROAD"), basename or "ROAD")

################################################
# ---------------4) BLOCK_FIX
################################################

def validate_block_fix(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `gpd.read_file().iterrows()` """
    
    print(f"  กำลังตรวจสอบ BLOCK FIX: {fc_name}")
    
    # [MODIFIED] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return
    
    required = ["STREET_NAME", "STREET_CODE", "BRANCH_CODE", "BLOCK_FIX_RN"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์นี้")
    
    # --- [MODIFIED] (ตรวจสอบประเภท) ---
    if "STREET_NAME" in schema and schema["STREET_NAME"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_NAME", schema["STREET_NAME"], "ต้องเป็น String")
    if "STREET_CODE" in schema and schema["STREET_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_CODE", schema["STREET_CODE"], "ต้องเป็น String")
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")
    if "BLOCK_FIX_RN" in schema and not is_numeric_field_type(schema["BLOCK_FIX_RN"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BLOCK_FIX_RN", schema["BLOCK_FIX_RN"], "ต้องเป็น Number")

    try:
        branch_rns = defaultdict(list)
        name_code_branch_list = []

        # [MODIFIED] อ่านข้อมูลด้วย Geopandas
        gdf = read_layer_data(gdb_path, fc_name, is_spatial=True)

        # [MODIFIED] วนลูปด้วย `iterrows()`
        for oid, row in gdf.iterrows():
            # [NEW] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 

            # ================================================================
            # --- [KEPT] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            branch = safe_to_none(rec.get("BRANCH_CODE"))
            rn = safe_to_none(rec.get("BLOCK_FIX_RN"))
            name = safe_to_none(rec.get("STREET_NAME"))
            code = safe_to_none(rec.get("STREET_CODE"))

            if pd.isna(name): name = None
            if pd.isna(code): code = None
            if pd.isna(branch): branch = None
            if pd.isna(rn): rn = None

            # 4.1.1. STREET_NAME ต้องเป็น String  และไม่ใช่ค่าว่าง (NULL) หรือ " " หรือขีดกลาง (-)
            if not name or (isinstance(name, str) and (name.strip() == "" or name.strip() == "-")):
                write_error_report(error_list, gdb_path, fc_name, "Data Required", oid, "STREET_NAME", name, "STREET_NAME ต้องไม่เป็นค่าว่าง, ช่องว่าง หรือ '-'")

            # 4.1.3: BRANCH_CODE ต้องเป็น String และมี 8 หลักเท่านั้น
            if not (branch and isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")
            
            # BLOCK_FIX_RN ต้องเป็น Number และใน BRANCH_CODE เดียวกัน ต้องไม่ซ้ำ
            if rn is None or not can_be_number(rn):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BLOCK_FIX_RN", rn, "ต้องเป็น Number")
            else:
                branch_rns[(branch.strip() if branch else "NULL", int(float(rn)))].append(oid)
            
            # 4.2. STREET_NAME กับ STREET_CODE ต้องจับคู่กันแบบ 1 ต่อ 1
            if name and code:
                name_code_branch_list.append((branch, name, code))
            # ================================================================
            # --- [KEPT] สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================

        # (KEPT) Logic for 1-to-1 checks
        name_to_code_by_branch = defaultdict(dict)
        code_to_name_by_branch = defaultdict(dict)

        for branch, name, code in name_code_branch_list:
            branch_key = branch.strip() if branch else "NULL"
            
            if name in name_to_code_by_branch[branch_key] and name_to_code_by_branch[branch_key][name] != code:
                write_error_report(error_list, gdb_path, fc_name, "OneToOne", "N/A", "STREET_NAME", name, f"{name} เชื่อมต่อกับ STREET_CODE มากกว่า 1 ค่า (ภายใน BRANCH_CODE '{branch_key}')")
            
            if code in code_to_name_by_branch[branch_key] and code_to_name_by_branch[branch_key][code] != name:
                write_error_report(error_list, gdb_path, fc_name, "OneToOne", "N/A", "STREET_CODE", code, f"{code} เชื่อมต่อกับ STREET_NAME มากกว่า 1 ค่า (ภายใน BRANCH_CODE '{branch_key}')")
            
            name_to_code_by_branch[branch_key][name] = code
            code_to_name_by_branch[branch_key][code] = name
        
        # (KEPT) Check RN duplicates
        for key, oids in branch_rns.items():
            if len(oids) > 1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "BLOCK_FIX_RN", key[1], f"BLOCK_FIX_RN ซ้ำใน BRANCH_CODE '{key[0]}'")

    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Geopandas read/loop) {ex}")

    # 4.3. [MODIFIED] ตรวจสอบโพลีกอนที่ซ้อนทับกันสนิท
    check_for_exact_overlaps(gdb_path, fc_name, error_list, os.path.join(OVERLAP_ROOT,"BLOCK"), basename or "BLOCK_FIX")

############################################
###----- 5) BLOCK_PRICE
############################################

def validate_block_price(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `gpd.read_file().iterrows()` """
    
    print(f"  กำลังตรวจสอบ BLOCK PRICE: {fc_name}")
    
    # [MODIFIED] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return
    
    required = ["STREET_NAME", "STREET_CODE", "BRANCH_CODE", "BLOCK_PRICE_RN"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์นี้")

    # [MODIFIED] (ตรวจสอบประเภทข้อมูล)
    if "STREET_NAME" in schema and schema["STREET_NAME"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_NAME", schema["STREET_NAME"], "ต้องเป็น String")
    if "STREET_CODE" in schema and schema["STREET_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "STREET_CODE", schema["STREET_CODE"], "ต้องเป็น String")
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")
    if "BLOCK_PRICE_RN" in schema and not is_numeric_field_type(schema["BLOCK_PRICE_RN"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BLOCK_PRICE_RN", schema["BLOCK_PRICE_RN"], "ต้องเป็น Number")

    try:
        branch_rns = defaultdict(list)
        
        # [MODIFIED] อ่านข้อมูลด้วย Geopandas
        gdf = read_layer_data(gdb_path, fc_name, is_spatial=True)
        
        # [MODIFIED] วนลูปด้วย `iterrows()`
        for oid, row in gdf.iterrows():
            # [NEW] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 
            
            # ================================================================
            # --- [KEPT] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            branch = safe_to_none(rec.get("BRANCH_CODE"))
            rn = safe_to_none(rec.get("BLOCK_PRICE_RN"))
            name = safe_to_none(rec.get("STREET_NAME"))

            if pd.isna(name): name = None
            if pd.isna(branch): branch = None
            if pd.isna(rn): rn = None

            # 5.1.1. STREET_NAME ต้องเป็น String และไม่ใช่ค่าว่าง (NULL) หรือ " " หรือขีดกลาง (-)E
            if not name or (isinstance(name, str) and (name.strip() == "" or name.strip() == "-")):
                write_error_report(error_list, gdb_path, fc_name, "Data Required", oid, "STREET_NAME", name, "STREET_NAME ต้องไม่เป็นค่าว่าง, ช่องว่าง หรือ '-'")
            
            # 5.1.3. BRANCH_CODE ต้องเป็น String  และมี 8 หลักเท่านั้น
            if not (branch and isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")
            
            # 5.1.4. BLOCK_PRICE_RN ต้องเป็น Number และใน BRANCH_CODE เดียวกัน ต้องไม่ซ้ำกัน
            if rn is None or not can_be_number(rn):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BLOCK_PRICE_RN", rn, "ต้องเป็น Number")
            else:
                branch_rns[(branch.strip() if branch else "NULL", int(float(rn)))].append(oid)
            # ================================================================
            # --- [KEPT] สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================
        
        # (KEPT) Check RN duplicates
        for key, oids in branch_rns.items():
            if len(oids) > 1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "BLOCK_PRICE_RN", key[1], f"BLOCK_PRICE_RN ซ้ำใน BRANCH_CODE '{key[0]}'")

    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Geopandas read/loop) {ex}")
    
    # 5.3. [MODIFIED] ตรวจสอบโพลีกอนที่ซ้อนทับกันสนิท
    check_for_exact_overlaps(gdb_path, fc_name, error_list, os.path.join(OVERLAP_ROOT,"BLOCK"), basename or "BLOCK_PRICE")

##############################################
#----------------- 6) BLOCK_BLUE
##############################################

def validate_block_blue(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `gpd.read_file().iterrows()` """
    
    print(f"  กำลังตรวจสอบ BLOCK_BLUE: {fc_name}")
    
    # [MODIFIED] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return

    required = ["BRANCH_CODE","BLOCK_BLUE_RN","BLOCK_TYPE_ID"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์นี้")
    
    # [MODIFIED] (ตรวจสอบประเภทข้อมูล - เพิ่ม)
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")
    if "BLOCK_BLUE_RN" in schema and not is_numeric_field_type(schema["BLOCK_BLUE_RN"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BLOCK_BLUE_RN", schema["BLOCK_BLUE_RN"], "ต้องเป็น Number")
    if "BLOCK_TYPE_ID" in schema and not is_numeric_field_type(schema["BLOCK_TYPE_ID"]):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BLOCK_TYPE_ID", schema["BLOCK_TYPE_ID"], "ต้องเป็น Number")

    try:
        branch_vals = defaultdict(list)
        
        # [MODIFIED] อ่านข้อมูลด้วย Geopandas
        gdf = read_layer_data(gdb_path, fc_name, is_spatial=True)
        
        # [MODIFIED] วนลูปด้วย `iterrows()`
        for oid, row in gdf.iterrows():
            # [NEW] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 

            # ================================================================
            # --- [KEPT] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            branch = safe_to_none(rec.get("BRANCH_CODE")); rn = safe_to_none(rec.get("BLOCK_BLUE_RN")); bt = safe_to_none(rec.get("BLOCK_TYPE_ID"))

            if pd.isna(branch): branch = None
            if pd.isna(rn): rn = None
            if pd.isna(bt): bt = None
            
            # 6.1.1. BRANCH_CODE ต้องเป็น String  และมี 8 หลักเท่านั้น
            if not (branch and isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")
            
            # BLOCK_BLUE_RN  ต้องเป็น Number และใน BRANCH_CODE เดียวกัน ต้องไม่ซ้ำกัน
            if rn is None or not can_be_number(rn):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BLOCK_BLUE_RN", rn, "ต้องเป็น Number")
            else:
                branch_vals[(branch.strip() if branch else "NULL", int(float(rn)))].append(oid)
            
            # 6.1.3. BLOCK_TYPE_ID ต้องเป็น Number และต้องเป็น 1 หรือ 2 หรือ 3 เท่านั้น
            if not can_be_number(bt) or int(float(bt)) not in (1,2,3):
                write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "BLOCK_TYPE_ID", bt, "BLOCK_TYPE_ID ต้องเป็น 1 หรือ 2 หรือ 3")
            # ================================================================
            # --- [KEPT] สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================
        
        # (KEPT) Check RN duplicates
        for k,oids in branch_vals.items():
            if len(oids)>1:
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "BLOCK_BLUE_RN", k[1], f"พบค่าซ้ำใน BRANCH_CODE '{k[0]}'")
    
    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Geopandas read/loop) {ex}")
    
    # 6.2. [MODIFIED] ตรวจสอบโพลีกอนที่ซ้อนทับกันสนิท
    check_for_exact_overlaps(gdb_path, fc_name, error_list, os.path.join(OVERLAP_ROOT,"BLOCK"), basename or "BLOCK_BLUE")

##############################################
#----------------- 7) PARCEL_REL (Table)
##############################################

def validate_parcel_rel(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `read_layer_data().iterrows()` """
    
    print(f"  กำลังตรวจสอบ PARCEL_REL: {fc_name}")
    
    # [MODIFIED] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return

    required = ["BRANCH_CODE","REL_RN","PARCEL_RN","STREET_RN","BLOCK_FIX_RN","BLOCK_BLUE_RN","BLOCK_PRICE_RN","TABLE_NO","SUB_TABLE_NO","DEPTH_R","DEPTH_GROUP","START_X","START_Y","END_X","END_Y"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบฟิลด์")

    # [MODIFIED] ตรวจสอบประเภทข้อมูลที่เป็น Number
    numeric_fields = ["REL_RN","PARCEL_RN","STREET_RN","BLOCK_FIX_RN","BLOCK_BLUE_RN","BLOCK_PRICE_RN","TABLE_NO","SUB_TABLE_NO","DEPTH_R","DEPTH_GROUP","START_X","START_Y","END_X","END_Y"]
    for nf in numeric_fields:
        if nf.upper() in schema and not is_numeric_field_type(schema[nf.upper()]):
            write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, nf, schema[nf.upper()], "ต้องเป็น Number")
    
    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")


    VALID_TABLE_NO = REL_TABLE_NO_DOMAIN
    VALID_SUB_TABLE_NO = REL_SUB_TABLE_NO_RANGE
    branch_rel_rn = defaultdict(list)
    
    try:
        # [MODIFIED] อ่านข้อมูล (Table, no geometry)
        df = read_layer_data(gdb_path, fc_name, is_spatial=False)
        
        # [MODIFIED] วนลูปด้วย `iterrows()`
        for oid, row in df.iterrows():
            # [NEW] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 
            
            # ================================================================
            # --- [KEPT] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            branch = safe_to_none(rec.get("BRANCH_CODE")); rel_rn = safe_to_none(rec.get("REL_RN"))
            
            if pd.isna(branch): branch = None
            if pd.isna(rel_rn): rel_rn = None

            # 7.1.1. BRANCH_CODE ต้องเป็น String  และมี 8 หลักเท่านั้น
            if not (branch and isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")
            
            # 7.1.2. REL_RN ต้องเป็น Number และใน BRANCH_CODE เดียวกันจะต้องไม่ซ้ำกัน
            if rel_rn is None or not can_be_number(rel_rn):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "REL_RN", rel_rn, "REL_RN ต้องเป็น Number")
            else:
                branch_key = branch.strip() if branch else "NULL"
                branch_rel_rn[(branch_key, int(float(rel_rn)))].append(oid)
            
            # 7.1.8. TABLE_NO ต้องเป็น Number
            table_no = safe_to_none(rec.get("TABLE_NO")); sub_no = safe_to_none(rec.get("SUB_TABLE_NO"))
            if pd.isna(table_no): table_no = None
            if pd.isna(sub_no): sub_no = None

            if not can_be_number(table_no) or int(float(table_no)) not in VALID_TABLE_NO:
                write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "TABLE_NO", table_no, f"ต้องเป็น {sorted(VALID_TABLE_NO)}")
            
            # 7.1.9. SUB_TABLE_NO ต้องเป็น Number และต้องมีค่าระหว่าง 0 - 6 หรือค่าว่าง เท่านั้น
            if sub_no is not None:
                if not can_be_number(sub_no) or int(float(sub_no)) not in VALID_SUB_TABLE_NO:
                    write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "SUB_TABLE_NO", sub_no, "ต้องเป็น 0-6 ")
            
            #   7.1.10. - 7.1.15.
            for fld in ("DEPTH_R","START_X","START_Y","END_X","END_Y"):
                val = safe_to_none(rec.get(fld))
                if pd.isna(val): val = None
                if val is None or (can_be_number(val) and float(val)==0.0):
                    write_error_report(error_list, gdb_path, fc_name, "Data Required", oid, fld, val, f"{fld} ต้องไม่ใช่ 0 หรือค่าว่าง")
            # ================================================================
            # --- [KEPT] สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================

        # (KEPT) Check RN duplicates
        for key, oids in branch_rel_rn.items():
            if len(oids) > 1:
                branch_key, rn_val = key
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "REL_RN", rn_val, f"ซ้ำภายใน BRANCH_CODE '{branch_key}'")

    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Fiona read/loop) {ex}")
    
    # (หมายเหตุ: ไม่มีการตรวจสอบทับซ้อนสำหรับ Table)

##############################################  
#---------------- 8) NS3K_REL (Table)
##############################################
def validate_ns3k_rel(gdb_path, fc_name, error_list, basename=None):
    """ [REFACTORED] `arcpy.da.SearchCursor` replaced with `read_layer_data().iterrows()` """

    print(f"  กำลังตรวจสอบ NS3K_REL: {fc_name}")
    
    # [MODIFIED] อ่าน Schema ด้วย Fiona
    schema = get_fiona_schema(gdb_path, fc_name)
    if not schema:
        write_error_report(error_list, gdb_path, fc_name, "Read Error", -1, "", "", "ไม่สามารถอ่าน Schema จากไฟล์ได้")
        return
    
    #   8.1. ตรวจสอบฟิลด์ที่จำเป็น
    required = ["BRANCH_CODE","REL_RN","NS3K_RN","STREET_RN","BLOCK_FIX_RN","BLOCK_BLUE_RN","BLOCK_PRICE_RN","TABLE_NO","SUB_TABLE_NO","DEPTH_R","DEPTH_GROUP","START_X","START_Y","END_X","END_Y"]
    for f in required:
        if f.upper() not in schema:
            write_error_report(error_list, gdb_path, fc_name, "Field Check", -1, f, "", "ไม่พบ field")
    
    # [MODIFIED] (ตรวจสอบประเภทข้อมูล)
    numeric_fields = ["REL_RN","NS3K_RN","STREET_RN","BLOCK_FIX_RN","BLOCK_BLUE_RN","BLOCK_PRICE_RN","TABLE_NO","SUB_TABLE_NO","DEPTH_R","DEPTH_GROUP","START_X","START_Y","END_X","END_Y"]
    for nf in numeric_fields:
        if nf.upper() in schema and not is_numeric_field_type(schema[nf.upper()]):
            write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, nf, schema[nf.upper()], "ต้องเป็น Number")

    if "BRANCH_CODE" in schema and schema["BRANCH_CODE"].lower() not in ("str", "string"):
        write_error_report(error_list, gdb_path, fc_name, "Field Type", -1, "BRANCH_CODE", schema["BRANCH_CODE"], "ต้องเป็น String")

    chk_fields = [f for f in required if f.upper() in schema]
    VALID_TABLE_NO = REL_TABLE_NO_DOMAIN
    VALID_SUB_TABLE_NO = REL_SUB_TABLE_NO_RANGE
    
    branch_rel_rn = defaultdict(list)
    
    try:
        # [MODIFIED] อ่านข้อมูล (Table, no geometry)
        df = read_layer_data(gdb_path, fc_name, is_spatial=False)

        # [MODIFIED] วนลูปด้วย `iterrows()`
        for oid, row in df.iterrows():
            # [NEW] สร้าง `rec` dict (ด้วย keys ตัวพิมพ์ใหญ่)
            rec = {k.upper(): v for k, v in row.items()}
            rec["OID@"] = oid 

            # ================================================================
            # --- [KEPT] โค้ดตรวจสอบเดิม (ไม่เปลี่ยนแปลง) ---
            # ================================================================
            branch = safe_to_none(rec.get("BRANCH_CODE")); 
            rel_rn = safe_to_none(rec.get("REL_RN"))
            ns3k_rn = safe_to_none(rec.get("NS3K_RN"))

            if pd.isna(branch): branch = None
            if pd.isna(rel_rn): rel_rn = None
            if pd.isna(ns3k_rn): ns3k_rn = None
            
            # 8.1.1. BRANCH_CODE ต้องเป็น String และมี 8 หลักเท่านั้น
            if not (branch and isinstance(branch,str) and len(branch.strip())==8 and branch.strip().isdigit()):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "BRANCH_CODE", branch, "BRANCH_CODE ต้องเป็น 8 หลัก")

            # 8.1.2. REL_RN ต้องเป็น Number และภายใน BRANCH_CODE เดียวกันจะต้องไม่ซ้ำกัน
            if rel_rn is None or not can_be_number(rel_rn):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "REL_RN",rel_rn , "ต้องเป็น Number")
            else:
                branch_key = branch.strip() if branch else "NULL"
                branch_rel_rn[(branch_key, int(float(rel_rn)))].append(oid)
            
            # 8.1.3 NS3K_RN ต้องเป็น Number
            if ns3k_rn is None or not can_be_number(ns3k_rn):
                write_error_report(error_list, gdb_path, fc_name, "Data Format", oid, "NS3K_RN", ns3k_rn, "NS3K_RN ต้องเป็น Number")
    
            # 8.1.8. TABLE_NO ต้องเป็น Number
            table_no = safe_to_none(rec.get("TABLE_NO")); sub_no = safe_to_none(rec.get("SUB_TABLE_NO"))
            if pd.isna(table_no): table_no = None
            if pd.isna(sub_no): sub_no = None

            if not can_be_number(table_no) or int(float(table_no)) not in VALID_TABLE_NO:
                write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "TABLE_NO", table_no, f"ต้องเป็น {sorted(VALID_TABLE_NO)} ")
            
            #  SUB_TABLE_NO ต้องเป็น Number และต้องมีค่าระหว่าง 0 – 6 หรือค่าว่าง เท่านั้น
            if sub_no is not None:
                if not can_be_number(sub_no) or int(float(sub_no)) not in VALID_SUB_TABLE_NO:
                    write_error_report(error_list, gdb_path, fc_name, "Data Specified", oid, "SUB_TABLE_NO", sub_no, "ต้องเป็น 0 หรือ 1-6 ")
            
            # 8.1.10. - 8.1.15.
            for fld in ("DEPTH_R","START_X","START_Y","END_X","END_Y"):
                val = safe_to_none(rec.get(fld))
                if pd.isna(val): val = None
                if val is None or (can_be_number(val) and float(val)==0.0):
                    write_error_report(error_list, gdb_path, fc_name, "Data Required", oid, fld, val, f"{fld} จะต้องไม่ใช่ 0 หรือค่าว่าง")
            # ================================================================
            # --- [KEPT] สิ้นสุดโค้ดตรวจสอบเดิม ---
            # ================================================================
        
        # (KEPT) Check RN duplicates
        for key, oids in branch_rel_rn.items():
            if len(oids) > 1:
                branch_key, rn_val = key
                write_error_report(error_list, gdb_path, fc_name, "Duplicate Value", str(oids), "REL_RN", rn_val, f"REL_RN ซ้ำภายใน BRANCH_CODE '{branch_key}'")

    except Exception as ex:
        write_error_report(error_list, gdb_path, fc_name, "Cursor Error", -1, "", "", f"(Fiona read/loop) {ex}")
    
    # (หมายเหตุ: ไม่มีการตรวจสอบทับซ้อนสำหรับ Table)

################################################
# --------------- [REFACTORED] MAIN
################################################

def main():
    print("เริ่มต้นกระบวนการตรวจสอบมาตรฐาน (เวอร์ชัน Geopandas/Fiona)...")

    # (KEPT) Validation map
    validation_map = {
        # Spatial (Feature Classes)
        "PARCEL": {"pattern": re.compile(r'^PARCEL_\d{2}_\d{2}$', re.IGNORECASE), "func": validate_parcel, "is_spatial": True},
        "PARCEL_NS3K": {"pattern": re.compile(r'^PARCEL_\d{2}_NS3K_\d{2}$', re.IGNORECASE), "func": validate_parcel_ns3k, "is_spatial": True},
        "ROAD": {"pattern": re.compile(r'^ROAD_\d{2}$', re.IGNORECASE), "func": validate_road, "is_spatial": True},
        "BLOCK_FIX": {"pattern": re.compile(r'^BLOCK_FIX_\d{2}$', re.IGNORECASE), "func": validate_block_fix, "is_spatial": True},
        "BLOCK_PRICE": {"pattern": re.compile(r'^BLOCK_PRICE_\d{2}$', re.IGNORECASE), "func": validate_block_price, "is_spatial": True},
        "BLOCK_BLUE": {"pattern": re.compile(r'^BLOCK_BLUE_\d{2}$', re.IGNORECASE), "func": validate_block_blue, "is_spatial": True},
        # Non-Spatial (Tables)
        "PARCEL_REL": {"pattern": re.compile(r'^PARCEL_REL_\d{2}$', re.IGNORECASE), "func": validate_parcel_rel, "is_spatial": False},
        "NS3K_REL": {"pattern": re.compile(r'^NS3K_REL_\d{2}$', re.IGNORECASE), "func": validate_ns3k_rel, "is_spatial": False}
    }

    gdb_paths = find_gdb_paths(ROOT_DIR)
    if not gdb_paths:
        print("ไม่พบ GDBs ยกเลิกการดำเนินการ.")
        return

    today_str = datetime.datetime.now().strftime('%Y-%m-%d')
    run_timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    
    gdb_report_dir = os.path.join(REPORT_ROOT, today_str)
    os.makedirs(gdb_report_dir, exist_ok=True)
    
    all_data_records = []
    error_summary_records = []


    for gdb in gdb_paths:
        print(f"\nกำลังดำเนินการ: {gdb}")
        
        gdb_error_list = []
        
        try:
            # (KEPT) Basename logic
            parent = os.path.basename(os.path.dirname(gdb))
            grandparent = os.path.basename(os.path.dirname(os.path.dirname(gdb)))
            basename = f"{grandparent}_{parent}"
           
            # [MODIFIED] - `arcpy.List...` replaced with `fiona.listlayers`
            try:
                # `fiona.listlayers` คืนค่ารายชื่อ Layer ทั้งหมดใน GDB
                fcs_and_tables = fiona.listlayers(gdb)
            except Exception as e:
                print(f"  !! ไม่สามารถอ่าน GDB (list layers) ได้: {gdb}. ข้อผิดพลาด: {e}")
                print("  !! (อาจเป็นเพราะ GDB เสียหาย หรือไม่พบไดรเวอร์ OpenFileGDB)")
                continue

            if not fcs_and_tables:
                print("  ไม่พบฟิเจอร์คลาสหรือตารางใน GDB.")
                continue
            
            for fc in fcs_and_tables:
                fc_upper = fc.upper()
                for key,meta in validation_map.items():
                    if meta["pattern"].match(fc_upper): 
                        # `key` is category (e.g., "PARCEL")
                        # `meta` is the dict with "func" and "is_spatial"
                        
                        print(f"  >> ตรวจสอบ {fc} ด้วย {key} Validator...")

                        # --- [REFACTORED] Counting Logic ---
                        try:
                            if meta["is_spatial"]:
                                # --- 1. นี่คือ Feature Class (ใช้ Geopandas) ---
                                # อ่านไฟล์เพื่อดำเนินการนับ
                                gdf = gpd.read_file(gdb, layer=fc)
                                total_count = len(gdf)
                                all_data_records.append([
                                    run_timestamp, gdb, fc, total_count, key
                                ])
                                
                                # 2. Get conditional counts
                                if key == "PARCEL":
                                    # Coerce types for safe filtering
                                    gdf['LAND_NO'] = pd.to_numeric(gdf.get('LAND_NO'), errors='coerce')
                                    gdf['PARCEL_TYPE'] = pd.to_numeric(gdf.get('PARCEL_TYPE'), errors='coerce')
                                    
                                    cln_df = gdf[
                                        (gdf['LAND_NO'].notna()) & (gdf['LAND_NO'] != 0) & 
                                        (gdf['PARCEL_TYPE'].isin([1, 4, 5]))
                                    ]
                                    all_data_records.append([run_timestamp, gdb, fc, len(cln_df), "PARCEL_CLN"])
                                
                                elif key == "PARCEL_NS3K":
                                    # Coerce types
                                    gdf['LAND_NO'] = pd.to_numeric(gdf.get('LAND_NO'), errors='coerce')
                                    gdf['PARCEL_TYPE'] = pd.to_numeric(gdf.get('PARCEL_TYPE'), errors='coerce')
                                    
                                    cln_df = gdf[
                                        (gdf['LAND_NO'].notna()) & (gdf['LAND_NO'] != 0) & 
                                        (gdf['PARCEL_TYPE'] == 3)
                                    ]
                                    all_data_records.append([run_timestamp, gdb, fc, len(cln_df), "NS3K_CLN"])
                                
                                # ล้าง GDF ออกจากหน่วยความจำ
                                del gdf 

                            else:
                                # --- 2. นี่คือ Table (ใช้ Fiona) ---
                                # ใช้วิธีที่เร็วกว่าในการนับ (ไม่โหลด geometry)
                                with fiona.open(gdb, layer=fc) as src:
                                    total_count = len(src)
                                all_data_records.append([
                                    run_timestamp, gdb, fc, total_count, key
                                ])
                                
                        except Exception as e:
                            print(f"  !! ไม่สามารถนับจำนวน (Pivot) {fc} ได้: {e}")
                            all_data_records.append([
                                run_timestamp, gdb, fc, "Error", key
                            ])
                        # --- End of Counting Logic ---


                        # [MODIFIED] รัน Validator 
                        # (ส่ง gdb path และ fc name แทน fc_path)
                        try:
                            meta["func"](gdb, fc, gdb_error_list, basename)
                        except Exception as e:
                            write_error_report(gdb_error_list, gdb, fc, "Validator Error", -1, "", "", str(e))
                        break  # Exit the inner loop once matched
            
            # ================================================================
            # --- [KEPT] โค้ดรายงานผล (ไม่เปลี่ยนแปลง) ---
            # (ส่วนนี้ใช้ Pandas อยู่แล้ว จึงไม่ต้องแก้ไข)
            # ================================================================
            if gdb_error_list:
                report_path = os.path.join(gdb_report_dir, f"{basename}_error_report.xlsx")
                try:
                    headers = [
                        'Timestamp', 'GDB_Path', 'Featureclass', 'Check_Type',
                        'Object_ID(s)', 'Field_Name', 'Invalid_Value', 'Message'
                    ]
                    error_df_gdb = pd.DataFrame(gdb_error_list, columns=headers)
                    error_df_gdb['GDB_Path'] = error_df_gdb['GDB_Path'].apply(get_short_gdb_path)
                    error_df_gdb['GDB'] = error_df_gdb['GDB_Path'].apply(lambda x: os.path.normpath(x).split('GDB')[-2] + 'GDB' + x.split('GDB')[-1] if 'GDB' in x else x)
                    error_df_gdb = error_df_gdb[['Timestamp','GDB','Featureclass','Check_Type','Object_ID(s)','Field_Name','Invalid_Value','Message']]

                    groups = {
                        "PARCEL": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^PARCEL_\d+_\d+$', case=False, na=False)],
                        "PARCEL_NS3K": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^PARCEL_\d+_NS3K_\d+$', case=False, na=False)],
                        "ROAD": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^ROAD_\d+$', case=False, na=False)],
                        "BLOCK_FIX": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^BLOCK_FIX_\d+$', case=False, na=False)],
                        "BLOCK_PRICE": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^BLOCK_PRICE_\d+$', case=False, na=False)],
                        "BLOCK_BLUE": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^BLOCK_BLUE_\d+$', case=False, na=False)],
                        "PARCEL_REL": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^PARCEL_REL_\d+$', case=False, na=False)],
                        "NS3K_REL": error_df_gdb[error_df_gdb['Featureclass'].str.match(r'^NS3K_REL_\d+$', case=False, na=False)],
                    }
                    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                        for sheet_name, df in groups.items():
                            if not df.empty:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)

                    print(f"  ✅ บันทึก Error Report แยกตามประเภท Featureclass เรียบร้อย: {report_path}")

                except Exception as e:
                    print(f"  ⚠️ เกิดข้อผิดพลาดในการสร้าง Error Report: {e}")
                    # Fallback to single sheet
                    try:
                        report_path_fallback = os.path.join(gdb_report_dir, f"{basename}_error_report_single_sheet.xlsx")
                        error_df_gdb.to_excel(report_path_fallback, sheet_name='Errors', index=False)
                        print(f"  -> รายงาน Excel (Fallback) ถูกบันทึก: {report_path_fallback} (พบ {len(gdb_error_list)} errors)")
                    except Exception as e2:
                        print(f"  !! ไม่สามารถเขียนรายงาน Excel (Fallback) ได้ {report_path_fallback}: {e2}")
                
                # สรุป Error สำหรับ Sheet 2
                try:
                    error_df = pd.DataFrame(gdb_error_list, columns=['Timestamp', 'GDB_Path', 'Featureclass', 'Check_Type', 'Object_ID(s)', 'Field_Name', 'Invalid_Value', 'Message'])
                    summary_df = error_df.groupby(['GDB_Path', 'Featureclass', 'Check_Type']).size().reset_index(name='Count of Errors')
                    summary_df['Timestamp'] = run_timestamp
                    summary_df = summary_df[['Timestamp', 'GDB_Path', 'Featureclass', 'Check_Type', 'Count of Errors']]
                    error_summary_records.extend(summary_df.values.tolist())
                    
                except Exception as e:
                    print(f"  !! ไม่สามารถสรุป Error GDB นี้ได้: {e}")

            else:
                print(f"  -> ไม่พบข้อผิดพลาด (ไม่ต้องสร้างไฟล์สำหรับ {basename})")
            
            # (No `in_memory` to delete)

        except Exception as e:
            print(f"  Failed processing {gdb}: {e}")
            # (No `in_memory` to delete)

    # ================================================================
    # --- [KEPT] โค้ดเขียนรายงานสรุป (ไม่เปลี่ยนแปลง) ---
    # (ส่วนนี้ใช้ Pandas อยู่แล้ว จึงไม่ต้องแก้ไข)
    # ================================================================
    print(f"\nกำลังเขียนรายงานสรุป Excel ที่: {SUMMARY_EXCEL_PATH}")
    try:
        # [FIX] Explicitly create the summary report's parent directory
        # This fixes Error 3 and may help with Error 1 & 2
        os.makedirs(os.path.dirname(SUMMARY_EXCEL_PATH), exist_ok=True)
        
        with pd.ExcelWriter(SUMMARY_EXCEL_PATH, engine='openpyxl') as writer:
            # Sheet 1: All_DATA
            if all_data_records:
                all_data_df = pd.DataFrame(all_data_records, columns=['Timestamp', 'GDB_Path', 'Featureclass', 'Count', 'Category'])
                all_data_df['GDB_Path'] = all_data_df['GDB_Path'].apply(get_short_gdb_path)
                all_data_df["Province"] = all_data_df["GDB_Path"].apply(extract_province)
                all_data_df['Category'] = all_data_df['Category'].replace('PARCEL_NS3K', 'NS3K')
                all_data_df = all_data_df[all_data_df['Count'] != 'Error']
                all_data_df['Count'] = pd.to_numeric(all_data_df['Count'])
                grouped_data_df = all_data_df.groupby(["Province", "Category"])["Count"].sum().reset_index()
                pivot_data_df = grouped_data_df.pivot(
                    index="Province",
                    columns="Category",
                    values="Count"
                ).fillna(0).astype(int)
                required_cols = ["BLOCK_BLUE", "BLOCK_FIX", "BLOCK_PRICE", "NS3K", "NS3K_CLN", "NS3K_REL", "PARCEL", "PARCEL_CLN", "PARCEL_REL", "ROAD"]
                for col in required_cols:
                    if col not in pivot_data_df.columns:
                        pivot_data_df[col] = 0
                final_cols = required_cols + [col for col in pivot_data_df.columns if col not in required_cols]
                pivot_data_df = pivot_data_df[final_cols]
                pivot_data_df.to_excel(writer, sheet_name='All_DATA', index=True) 
                print(f"  -> เขียน Sheet 'All_DATA' (Pivoted) ({len(pivot_data_df)} แถว)")
            else:
                print("  -> ไม่มีข้อมูลสำหรับ 'All_DATA'")

            # Sheet 2: Error SUM
            if error_summary_records:
                error_sum_df = pd.DataFrame(error_summary_records, columns=['Timestamp', 'GDB_Path', 'Featureclass', 'Check_Type', 'Count of Errors'])
                error_sum_df['GDB_Path'] = error_sum_df['GDB_Path'].apply(get_short_gdb_path)
                error_sum_df.to_excel(writer, sheet_name='Error SUM', index=False)
                print(f"  -> เขียน Sheet 'Error SUM' ({len(error_sum_df)} แถว)")
                
                # Sheet 3: Errors_by_Province
                try:
                    df_report = error_sum_df.copy()
                    df_report["Province"] = df_report["GDB_Path"].apply(extract_province)
                    df_report["Category"] = df_report["Featureclass"].apply(categorize_featureclass)
                    df_report = df_report.dropna(subset=["Category"])
                    df_report["Category"] = df_report["Category"].replace('PARCEL_NS3K', 'NS3K')
                    grouped_df = df_report.groupby(["Province", "Category"])["Count of Errors"].sum().reset_index()
                    pivot_df = grouped_df.pivot(
                        index="Province", 
                        columns="Category", 
                        values="Count of Errors"
                    ).fillna(0).astype(int)
                    err_report_cols = ["BLOCK_BLUE", "BLOCK_FIX", "BLOCK_PRICE",
                                       "NS3K", "NS3K_REL",
                                       "PARCEL", "PARCEL_REL", "ROAD"]
                    for col in err_report_cols:
                        if col not in pivot_df.columns:
                            pivot_df[col] = 0
                    pivot_df = pivot_df[err_report_cols]
                    pivot_df.to_excel(writer, sheet_name='Errors_by_Province', index=True)
                    print(f"  -> เขียน Sheet 'Errors_by_Province' ({len(pivot_df)} แถว)")
                except Exception as e:
                    print(f"  !! ล้มเหลวในการสร้าง Sheet 'Errors_by_Province': {e}")
            else:
                print("  -> ไม่มีข้อมูลสำหรับ 'Error SUM'")
        print("  -> บันทึกไฟล์สรุป Excel เรียบร้อยแล้ว")
    except Exception as e:
        print(f"  !! ล้มเหลวในการเขียนไฟล์สรุป Excel: {e}")
        print("  !! (โปรดตรวจสอบว่าไฟล์ Excel ปิดอยู่ และ/หรือ มีสิทธิ์เขียนทับ)")

    print("\nเสร็จแล้วจ้า ดูผลลัพธ์ได้เลยจ้า")

if __name__ == "__main__":
    main()