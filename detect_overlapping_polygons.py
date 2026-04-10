"""
detect_overlapping_polygons.py
==================
จุดประสงค์: ตรวจจับโพลีกอนทับซ้อนกัน
ส่งรายงานเป็น Excel หนึ่งไฟล์ต่อ GDB หนึ่งไฟล์ (สรุป + แผ่นงานต่อประเภท)
ถ้า เปิด ADD_FEATURE_CLASS = True ก็จะเขียนฟีเจอรร์คลาสตำแหน่งที่ทับซ้อนกันลง GDB ผลลัพธ์ด้วย 

- Version 1.0
- release date: 2026-04-08
- ผู้เขียน: Apirat Rattanapaiboon

==========================
ขั้นตอนการทำงานหลัก

1) สแกนโครงสร้างไดเร็กทอรีเพื่อหาไฟล์ data.gdb
ตัวอย่างเก็บข้อมูลในโครงสร้างแบบนี้นะจ๊ะ
"10_กรุงเทพฯ
|----GDB_10_1
    |----data.gdb"
|----GDB_10_2
    |----data.gdb"

* gdb ชื่อ data.gdb เหมือนกันทุกไฟล์ เก็บอยู่ในซับไดเร็กทอรี่ที่มีชื่อแตกต่างกัน (เช่น GDB_10_1, GDB_10_2) 
ดังนั้นผมจะใช้ชื่อไดเร็กทอรี่ที่เก็บ data.gdb (เช่น GDB_10_1) เป็นตัวระบุแต่ละชุดข้อมูล

=======================
2) ค้นหาคลาสฟีเจอร์ใน gdb ที่ตั้งชื่อตรงกับรูปแบบ PARCEL_{zone}_{province}
เช่น PARCEL_47_10 หรือ PARCEL_48_44 เท่านั้น

==============================

3) ตรวจจับการซ้อนทับกันของโพลีกอนในแต่ละฟีเจอร์คลาสที่พบ
ทั้งนี้เป็นการตรวจรูปร่างโพลีกอนเท่านั้น ไม่ได้ตรวจข้อมูลแอตทริบิวต์อื่น ๆ เช่น ระวาง แผ่น เลขที่ดิน มาตราส่วน ฯลฯ
โดยกำหนดนัยยะสำคัญของค่าปัดเศษในการตรวจสอบว่าไม่ซ้อนทับกัน ไว้ที่ 0.01 (MIN_OVERLAP_AREA)
ค่านี้ปรับเปลี่ยนได้ตามความเหมาะสมเพื่อกรองการซ้อนทับที่มีพื้นที่เล็กมากจนไม่ถือว่ามีความสำคัญ (ค่านี้เป็นตารางเมตรนะจ๊ะ)
(หมายเหตุ ค่า 0.01 ผมใช้สำหรับการทำงานพื้นพาณิชยกรรมในกรุงเทพฯ
ถ้าเป็นพื้นที่ที่มีแปลงใหญ่ ๆ อาจต้องปรับเพิ่มมากกว่านั้นเพื่อกรองการซ้อนทับที่เล็กมาก ๆ ออกไป)
ซึ่งจะกำหนดการซ้อนทับกันเป็น 3 ประเภทหลัก ๆ ตามอัตราส่วนพื้นที่ซ้อนทับ (overlap ratio) ดังนี้
─────────────
  DUPLICATE  เป็นแปลงที่ซ้อนกัน เกือบทั้งหมด โดยใช้ค่า CONTAINMENT_THRESHOLD เป็นตัวกำหนด
                (ตั้งต้นให้ที่ ≥ 90% ของพื้นที่ซ้อนทับกันทั้งสองโพลีกอน แต่สามารถปรับได้ตามความเหมาะสม)

  CONTAINED  ในกรณีที่ไม่ใช่ DUPLICATE แต่โพลีกอนหนึ่งซ้อนทับมากกว่า CONTAINMENT_THRESHOLD
             ซึ่งเป็นลักษณะของโพลีกอนที่อยู่ข้างในอีกอันหนึ่งอย่างชัดเจน
             (นั่นคือกรณี แปลงแบ่งแยก หรือ การรวมแปลง)
             แต่เนื่องจากเป็นชั้นฟีเจอร์คลาสเดียวกัน อาจมีการเปลี่ยนแปลงข้อมูลระหว่างทำงาน
             จึงไม่สามารถแยกได้ว่า เป็นกรณีรวมแปลง หรือ แยกแปลง 
             เพราะไม่เชื่อมั่นใน Date Update หรือลำดับ OID ในการตัดสินใจ
             จึงพิจารณาแค่ว่า โพลีกอนไหนมีอัตราส่วนการซ้อนทับมากกว่าเกณฑ์ที่กำหนด
             (เช่น ถ้า A ซ้อนทับ B มากกว่าเกณฑ์ แต่ B ซ้อนทับ A ไม่ถึงเกณฑ์ ก็จะถือว่า A เป็น CHECK และ B เป็น COMPARE)

  PARTIAL    กลุ่มนี้คือโพลีกอนที่ซ้อนทับกันแบบเหลื่อมกันบางส่วน แต่ไม่เข้าเกณฑ์ DUPLICATE หรือ CONTAINED


สิ่งที่ส่งออกในรายงาน Excel
──────────────────────
  Sheet — SUMMARY    : สรุปจำนวน
  Sheet — DUPLICATE  : สำหรับข้อมูลประเภท DUPLICATE
  Sheet — CONTAINED  : สำหรับข้อมูลประเภท CONTAINED
  Sheet — PARTIAL    : สำหรับข้อมูลประเภท PARTIAL

การตั้งค่า
──────────────────────
สิ่งที่ต้องตั้งค่าในส่วน "ตั้งค่า" คือ
ROOT_DIR  ใส่พาร์ธที่โฟลเดอร์ที่เก็บ gdb อย่างเช่น r"D:\A02-Projects\WarRoom\TestDetect\GDB" 

REPORT_DIR ใส่พาร์ธที่โฟลเดอร์ที่เก็บรายงาน Excel ที่สร้างขึ้น เช่น r"D:\A02-Projects\WarRoom\TestDetect\Reports"

RESULT_GDB ใส่พาร์ธและชื่อ gdb ที่จะใช้เก็บผลลัพธ์ในรูปแบบ gdb เช่น r"D:\A02-Projects\WarRoom\TestDetect\OverlapResults.gdb"
หมายเหตุ
ชื่อไฟล์ต้องลงท้ายด้วย .gdb 
ถ้าตั้งค่า ADD_FEATURE_CLASS = False จะไม่สร้างฟีเจอร์คลาสและไม่เขียนผลลัพธ์ลง GDB
ถ้า True จะเขียนฟีเจอร์คลาสตำแหน่งที่ทับซ้อนกันลง RESULT_GDB
ถ้า False จะไม่เขียนฟีเจอร์คลาส แต่ยังสร้างรายงาน Excel ตามปกติ

CONTAINMENT_THRESHOLD คือกำหนดเกณฑ์การพิจารณาว่าโพลีกอนหนึ่งถูกซ้อนทับอยู่ในอีกอันหนึ่งหรือไม่
ซึ่งตั้งค่าต้นให้ที่ 0.90 (90%) ของพื้นที่ซ้อนทับกันทั้งสองโพลีกอน แต่สามารถปรับได้ตามความเหมาะสม
ตามปกติเวลาทำเชปไฟล์จากกรมที่ดินสำหรับรอบส่งแปลงปรับปรุงแต่ละรอบ ผมจะตั้งค่าที่ 0.98 เพื่อให้แน่ใจว่าโพลีกอนที่จัดว่าเป็น DUPLICATE นั้นซ้อนทับกันเกือบทั้งหมดจริง ๆ

MIN_OVERLAP_AREA เป็นค่าที่กำหนดว่า ถ้ามีพื้นที่ซ้อนทับน้อยกว่าที่กำหนด จะไม่ถือว่าเป็นการทับซ้อนที่มีนัยสำคัญ
ค่า 0.01 ที่ตั้งต้นไว้เป็นตารางเมตร (sqm) ซึ่งเหมาะสำหรับการทำงานกับพื้นที่ที่มีแปลงขนาดเล็ก เช่น พื้นที่ในกรุงเทพฯ
ถ้าเป็นพื้นที่ที่มีแปลงใหญ่ ๆ อาจต้องปรับเพิ่มมากกว่านั้นเพื่อกรองการซ้อนทับที่เล็กมาก ๆ ออกไป
ขึ้นอยู่กับลักษณะของพื้นที่และความต้องการในการใช้งาน

"""

import os
import re
import sys
import logging
import traceback
from collections import Counter
from pathlib import Path
from datetime import datetime as dt

import arcpy
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════════
# ตั้งค่า — แก้ไขเฉพาะส่วนนี้เท่านั้นนะจ๊ะ
# ══════════════════════════════════════════════════════════════════════════════

# ROOT_DIR คือโฟลเดอร์หลักที่เก็บ GDB ย่อย ๆ ไว้ เช่น D:\Test\GDB_10_01\data.gdb
ROOT_DIR   = r"D:\A02-Projects\WarRoom\TestDetect\GDB"
# REPORT_DIR คือโฟลเดอร์ที่เก็บรายงาน Excel ที่สร้างขึ้น
REPORT_DIR = r"D:\A02-Projects\WarRoom\TestDetect\Reports"
# RESULT_GDB คือไฟล์ GDB ที่จะเขียนฟีเจอร์คลาสตำแหน่งที่ทับซ้อนกันลงไป 
# (จะสร้างเมื่อ ADD_FEATURE_CLASS = True)
RESULT_GDB = r"D:\A02-Projects\WarRoom\TestDetect\OverlapResults.gdb"

# ถ้า True จะเขียนฟีเจอร์คลาสตำแหน่งที่ทับซ้อนกันลง RESULT_GDB
# ถ้า False จะไม่เขียนฟีเจอร์คลาส แต่ยังสร้างรายงาน Excel ตามปกติ
ADD_FEATURE_CLASS = True

# กำหนดเกณฑ์การพิจารณาว่าโพลีกอนหนึ่งถูกซ้อนทับอยู่ในอีกอันหนึ่งหรือไม่
# 0.90 → 90 % ของพื้นที่
# ปรับได้ตามความเหมาะสม
CONTAINMENT_THRESHOLD = 0.90

# กำหนดว่า ถ้า เล็กกว่าจำนวนตารางเมตรนี้ จะไม่ถือว่าเป็นการทับซ้อนที่มีนัยสำคัญ
# หน่วยเป็น ตารางเมตร (sqm) นะจ๊ะ
MIN_OVERLAP_AREA = 0.01

# ══════════════════════════════════════════════════════════════════════════════
# หลังจากบรรทัดนี้ ถ้าไม่รู้ว่ามันคืออะไร ก็อย่าแก้ไขเลยนะจ๊ะ
# ══════════════════════════════════════════════════════════════════════════════

# กำหนดรูปแบบชื่อฟีเจอร์คลาสที่ต้องการตรวจสอบ (PARCEL_{zone}_{province})
FC_PATTERN = re.compile(r"^PARCEL_(47|48)_\d{2}$", re.IGNORECASE)

# กำหนดชื่อเลเยอร์ชั่วคราวสำหรับการทำ Intersect ใน detect_overlaps()
_LYR_A = "_overlap_lyr_a"
_LYR_B = "_overlap_lyr_b"

# ══════════════════════════════════════════════════════════════════════════════
# รายงานเอ็กเซล
# ══════════════════════════════════════════════════════════════════════════════

_PAIR_EXCEL_COLS = [
    "CHECK_OID",
    "COMPARE_OID",
    "OVERLAP_TYPE",
    "REASON",
    "GDB_FOLDER",
    "BRANCH_CODE",
    "FC_Name",
    "Area_Check_sqm",
    "Area_Compare_sqm",
    "Overlap_Pct_Check",
    "Overlap_Pct_Compare",
    "Overlap_Area_sqm",
]

_PARTIAL_EXCEL_COLS = [
    "CHECK_OID",
    "OVERLAP_COUNT",
    "COMPARE_OIDS",
    "OVERLAP_TYPE",
    "GDB_FOLDER",
    "BRANCH_CODE",
    "FC_Name",
    "Area_Check_sqm",
    "Total_Overlap_Area_sqm",
]


# ══════════════════════════════════════════════════════════════════════════════
# กำหนด GDB FIELDS ที่จะใช้ในฟีเจอร์คลาสผลลัพธ์
# ══════════════════════════════════════════════════════════════════════════════

_FC_FIELDS = [
    ("CHECK_OID",          "LONG",   "Check OID",          None),
    ("COMPARE_OID",        "LONG",   "Compare OID",        None),
    ("OVERLAP_TYPE",       "TEXT",   "Overlap Type",       30),
    ("REASON",             "TEXT",   "Reason",             500),
    ("BRANCH_CODE",        "TEXT",   "Branch Code",        50),
    ("FC_Name",            "TEXT",   "Feature Class Name", 40),
    ("Area_Check_sqm",     "DOUBLE", "Area Check (m²)",    None),
    ("Area_Compare_sqm",   "DOUBLE", "Area Compare (m²)",  None),
    ("Overlap_Pct_Check",  "DOUBLE", "Overlap % Check",    None),
    ("Overlap_Pct_Compare","DOUBLE", "Overlap % Compare",  None),
    ("Overlap_Area_sqm",   "DOUBLE", "Overlap Area (m²)",  None),
]

# อันนี้มีไว้จัดการเรื่องขีดจำกัดความยาวฟิลด์ TEXT ที่คำนวณไว้ล่วงหน้าสำหรับตัวป้องกันการตัดทอนใน write_fc()
_FC_TEXT_LIMITS = {
    f[0]: f[3]
    for f in _FC_FIELDS
    if f[1] == "TEXT" and f[3] is not None
}


# ══════════════════════════════════════════════════════════════════════════════
# LOGGING
# ══════════════════════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════════════════════
# ตั้งค่าระบบและสร้างโครงสร้างผลลัพธ์ (output directories, GDB) ก่อนเริ่มประมวลผล
# ══════════════════════════════════════════════════════════════════════════════

def setup():
    """Create output directories and the result GDB if they do not already exist."""
    Path(REPORT_DIR).mkdir(parents=True, exist_ok=True)

    if ADD_FEATURE_CLASS and not arcpy.Exists(RESULT_GDB):
        arcpy.management.CreateFileGDB(
            str(Path(RESULT_GDB).parent),
            Path(RESULT_GDB).name,
        )
        log.info(f"Created result GDB: {RESULT_GDB}")
    arcpy.env.overwriteOutput = True

# ══════════════════════════════════════════════════════════════════════════════
# DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════

def find_gdbs(root: str) -> list[str]:
    gdbs = []
    for dirpath, dirnames, _ in os.walk(root):
        for dirname in dirnames:
            if dirname.lower().endswith(".gdb"):
                full = os.path.join(dirpath, dirname)
                log.info(f"Found GDB: {full}")
                gdbs.append(full)
    return gdbs


def find_fcs(gdb: str) -> list[str]:
    fcs = []
    for dirpath, _, names in arcpy.da.Walk(gdb, datatype="FeatureClass"):
        for name in names:
            if FC_PATTERN.match(name):
                fcs.append(os.path.join(dirpath, name))
    return fcs


# ══════════════════════════════════════════════════════════════════════════════
# OVERLAP DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def _cleanup_temp_layers():
    for lyr in (_LYR_A, _LYR_B):
        if arcpy.Exists(lyr):
            arcpy.Delete_management(lyr)

def detect_overlaps(fc: str) -> list[dict]:
    temp = os.path.join(arcpy.env.scratchGDB, "tmp_isect")

    _cleanup_temp_layers()
    if arcpy.Exists(temp):
        arcpy.Delete_management(temp)

    arcpy.MakeFeatureLayer_management(fc, _LYR_A)
    arcpy.MakeFeatureLayer_management(fc, _LYR_B)

    arcpy.analysis.Intersect([_LYR_A, _LYR_B], temp)

    fields     = [f.name for f in arcpy.ListFields(temp)]
    fid_fields = sorted(f for f in fields if f.startswith("FID_"))

    if len(fid_fields) < 2:
        log.warning("    Intersect produced fewer than 2 FID_ fields — skipping.")
        return []

    fid_a, fid_b = fid_fields[0], fid_fields[1]

    pairs = []
    seen  = set()

    with arcpy.da.SearchCursor(temp, [fid_a, fid_b, "SHAPE@AREA", "SHAPE@"]) as cur:
        for oid_a, oid_b, area, geom in cur:
            if oid_a >= oid_b:
                continue
            if area < MIN_OVERLAP_AREA:
                continue
            key = (oid_a, oid_b)
            if key in seen:
                continue
            seen.add(key)
            pairs.append({"a": oid_a, "b": oid_b, "area": area, "geom": geom})

    return pairs

# ══════════════════════════════════════════════════════════════════════════════
# CLASSIFICATION
# ══════════════════════════════════════════════════════════════════════════════

def classify_pair(
    ov: float,
    area_a: float,
    area_b: float,
    threshold: float,
) -> dict:
    """
    ตรรกะการตัดสินใจ
    Classification rules  
    ────────────────────
    ra ≥ T AND rb ≥ T  →  DUPLICATE  (สองโพลีกอนซ้อนทับกันเกือบทั้งหมด /ตัดเปอร์เซ็นต์ตาม threshold/)
    ra ≥ T OR  rb ≥ T  →  CONTAINED  (โพลีกอนซ้อนทับกันมากกว่า T% แต่ไม่ถึงกับเหมือนกัน — อาจเป็นกรณีรวมแปลงหรือแยกแปลง)
    neither            →  PARTIAL    (มีการซ้อนทับกันบางส่วน)

        Returns
    ───────
    Dict with keys:
      overlap_type  — "DUPLICATE" | "CONTAINED" | "PARTIAL"
      check_key     — "a" or "b"  (never None)
      compare_key   — "a" or "b"  (never None)
      ra, rb        — computed overlap ratios (floats)
      reason_tmpl   — format string for the REASON column
    """
    ra = ov / area_a if area_a else 0.0
    rb = ov / area_b if area_b else 0.0

    # ── DUPLICATE ─────────────────────────────────────────────────────────────
    if ra >= threshold and rb >= threshold:
        return {
            "overlap_type": "DUPLICATE",
            "check_key":    "a",
            "compare_key":  "b",
            "ra": ra, "rb": rb,
            "reason_tmpl": (
                "OID {CHECK_OID} กับ {COMPARE_OID} ซ้อนทับกัน"
                "({pct_check:.1f}% กับ {pct_compare:.1f}%)"
            ),
        }

    # ── CONTAINED ─────────────────────────────────────────────────────────────
    if ra >= threshold or rb >= threshold:
        if ra >= rb:
            check_key, compare_key = "a", "b"
        else:
            check_key, compare_key = "b", "a"

        return {
            "overlap_type": "CONTAINED",
            "check_key":    check_key,
            "compare_key":  compare_key,
            "ra": ra, "rb": rb,
            "reason_tmpl": (
                "เนื้อที่ {pct_check:.1f}% ของ OID {CHECK_OID} ซ้อนทับเนื้อที่ {pct_compare:.1f}% ของ OID {COMPARE_OID}"
            ),
        }

    # ── PARTIAL ───────────────────────────────────────────────────────────────
    return {
        "overlap_type": "PARTIAL",
        "check_key":    "a",   
        "compare_key":  "b",   
        "ra": ra, "rb": rb,
        "reason_tmpl": (
            "OID {CHECK_OID} และ {COMPARE_OID} ซ้อนทับกัน {overlap_area:.2f} ตารางเมตร"
            "({pct_a:.1f}% ของ {CHECK_OID} ทับกับ {pct_b:.1f}% ของ {COMPARE_OID})"
        ),
    }


# ══════════════════════════════════════════════════════════════════════════════
# BUILD RECORDS
# ══════════════════════════════════════════════════════════════════════════════

def build_records(gdb: str, fc: str, pairs: list[dict]) -> list[dict]:
    if not pairs:
        return []

    fc_name  = Path(fc).name
    gdb_name = Path(gdb).parent.name

    field_names = [f.name for f in arcpy.ListFields(fc)]
    has_branch  = "BRANCH_CODE" in field_names
    read_fields = ["OID@", "SHAPE@AREA"] + (["BRANCH_CODE"] if has_branch else [])

    areas  = {}
    branch = {}

    with arcpy.da.SearchCursor(fc, read_fields) as cur:
        for row in cur:
            oid         = row[0]
            areas[oid]  = row[1]
            branch[oid] = row[2] if has_branch else "UNKNOWN"

    records = []

    for p in pairs:
        a, b   = p["a"], p["b"]
        area_a = areas.get(a, 0.0)
        area_b = areas.get(b, 0.0)
        ov     = p["area"]

        cls = classify_pair(ov, area_a, area_b, CONTAINMENT_THRESHOLD)

        oid_map  = {"a": a,             "b": b}
        area_map = {"a": area_a,        "b": area_b}
        pct_map  = {"a": cls["ra"]*100, "b": cls["rb"]*100}

        chk_key     = cls["check_key"]
        compare_key = cls["compare_key"]

        check_oid   = oid_map[chk_key]
        compare_oid = oid_map[compare_key]

        fmt = {
            "CHECK_OID":    check_oid,
            "COMPARE_OID":  compare_oid,
            "pct_check":    pct_map[chk_key],
            "pct_compare":  pct_map[compare_key],
            "overlap_area": ov,
            "pct_a":        cls["ra"] * 100,
            "pct_b":        cls["rb"] * 100,
        }
        reason = cls["reason_tmpl"].format(**fmt)

        branch_val = branch.get(check_oid, "UNKNOWN")

        records.append({
            "CHECK_OID":           check_oid,
            "COMPARE_OID":         compare_oid,
            "OVERLAP_TYPE":        cls["overlap_type"],
            "REASON":              reason,
            "GDB_FOLDER":          gdb_name,
            "BRANCH_CODE":         branch_val,
            "FC_Name":             fc_name,
            "Area_Check_sqm":      area_map[chk_key],
            "Area_Compare_sqm":    area_map[compare_key],
            "Overlap_Pct_Check":   pct_map[chk_key],
            "Overlap_Pct_Compare": pct_map[compare_key],
            "Overlap_Area_sqm":    ov,
            "_geom":               p["geom"],
        })

    return records


# ══════════════════════════════════════════════════════════════════════════════
# PARTIAL CONSOLIDATION
# ══════════════════════════════════════════════════════════════════════════════

def _consolidate_partial(records: list[dict]) -> pd.DataFrame:

    partial = [r for r in records if r["OVERLAP_TYPE"] == "PARTIAL"]

    if not partial:
        return pd.DataFrame(columns=_PARTIAL_EXCEL_COLS)

    df = pd.DataFrame(partial)

    def _join_partner_oids(series) -> str:
        """Sort partner OIDs numerically and join as a comma-separated string."""
        sorted_oids = sorted(int(v) for v in series if v is not None)
        return ", ".join(str(v) for v in sorted_oids)

    grouped = (
        df
        .groupby(["GDB_FOLDER", "FC_Name", "CHECK_OID"], sort=False)
        .agg(
            OVERLAP_COUNT          = ("COMPARE_OID",      "count"),
            COMPARE_OIDS           = ("COMPARE_OID",      _join_partner_oids),
            OVERLAP_TYPE           = ("OVERLAP_TYPE",     "first"),
            BRANCH_CODE            = ("BRANCH_CODE",      "first"),
            Area_Check_sqm         = ("Area_Check_sqm",   "first"),
            Total_Overlap_Area_sqm = ("Overlap_Area_sqm", "sum"),
        )
        .reset_index()
    )

    return grouped[_PARTIAL_EXCEL_COLS]


# ══════════════════════════════════════════════════════════════════════════════
# Report
# ══════════════════════════════════════════════════════════════════════════════

def write_excel(gdb: str, records: list[dict]) -> None:

    ts   = dt.now().strftime("%Y%m%d_%H%M%S")
    name = Path(gdb).parent.name
    path = os.path.join(REPORT_DIR, f"{name}_{ts}.xlsx")

    with pd.ExcelWriter(path, engine="openpyxl") as writer:

        if not records:
            pd.DataFrame([{"Message": "No overlaps found in this GDB."}]).to_excel(
                writer, sheet_name="Results", index=False
            )
            log.info(f"  No overlaps — Excel written → {path}")
            return

        df = pd.DataFrame(records)
        summary = (
            df.groupby(["GDB_FOLDER", "OVERLAP_TYPE"])
              .size()
              .reset_index(name="Count_Pairs")
        )
        partial_unique = (
            df[df["OVERLAP_TYPE"] == "PARTIAL"]
            .groupby("GDB_FOLDER")["CHECK_OID"]
            .nunique()
            .reset_index(name="PARTIAL_Unique_Polygons")
        )
        summary = summary.merge(partial_unique, on="GDB_FOLDER", how="left")
        summary.to_excel(writer, sheet_name="SUMMARY", index=False)

        # ── DUPLICATE ─────────────────────────────────────────────────────────
        dup = df[df["OVERLAP_TYPE"] == "DUPLICATE"][_PAIR_EXCEL_COLS]
        if not dup.empty:
            dup.to_excel(writer, sheet_name="DUPLICATE", index=False)

        # ── CONTAINED ─────────────────────────────────────────────────────────
        con = df[df["OVERLAP_TYPE"] == "CONTAINED"][_PAIR_EXCEL_COLS]
        if not con.empty:
            con.to_excel(writer, sheet_name="CONTAINED", index=False)

        # ── PARTIAL (consolidated) ────────────────────────────────────────────        
        partial_consolidated = _consolidate_partial(records)
        if not partial_consolidated.empty:
            partial_consolidated.to_excel(writer, sheet_name="PARTIAL", index=False)
            n_pairs = len(df[df["OVERLAP_TYPE"] == "PARTIAL"])
            n_poly  = len(partial_consolidated)
            log.info(
                f"  PARTIAL: {n_pairs} pairs → {n_poly} unique polygons "
                f"(avg {n_pairs/n_poly:.1f} partners per polygon)."
            )

    log.info(f"  Excel written → {path}")


# ══════════════════════════════════════════════════════════════════════════════
# FEATURE CLASS
# ══════════════════════════════════════════════════════════════════════════════

def ensure_fc(gdb: str, fc_name: str, sr: arcpy.SpatialReference) -> str:
    raw_name = f"O_{Path(gdb).parent.name}_{fc_name}"
    name     = raw_name[:60]

    if len(raw_name) > 60:
        log.warning(
            f"  Output FC name truncated: '{raw_name}' → '{name}'. "
            "Check for name collisions if multiple GDBs share a long prefix."
        )

    path = os.path.join(RESULT_GDB, name)

    if arcpy.Exists(path):
        return path

    arcpy.CreateFeatureclass_management(
        RESULT_GDB, name, "POLYGON", spatial_reference=sr
    )

    for field_name, field_type, field_alias, field_length in _FC_FIELDS:
        arcpy.AddField_management(
            path,
            field_name,
            field_type,
            field_alias=field_alias,
            field_length=field_length,
        )

    log.info(f"  Created output FC: {path}")
    return path


def _trunc(field: str, value) -> str | None:
    if value is None or not isinstance(value, str):
        return value
    max_len = _FC_TEXT_LIMITS.get(field)
    if max_len and len(value) > max_len:
        log.warning(
            f"    Field '{field}' truncated: {len(value)} chars → {max_len} chars."
        )
        return value[:max_len]
    return value


def write_fc(fc: str, gdb: str, fc_name: str, records: list[dict]) -> None:

    if not records:
        return

    sr  = arcpy.Describe(fc).spatialReference
    out = ensure_fc(gdb, fc_name, sr)

    insert_fields = ["SHAPE@"] + [f[0] for f in _FC_FIELDS]

    with arcpy.da.InsertCursor(out, insert_fields) as cur:
        for r in records:
            cur.insertRow([
                r["_geom"],
                r["CHECK_OID"],
                r["COMPARE_OID"],
                _trunc("OVERLAP_TYPE",      r["OVERLAP_TYPE"]),
                _trunc("REASON",            r["REASON"]),
                _trunc("BRANCH_CODE",       r["BRANCH_CODE"]),
                _trunc("FC_Name",           r["FC_Name"]),
                r["Area_Check_sqm"],
                r["Area_Compare_sqm"],
                r["Overlap_Pct_Check"],
                r["Overlap_Pct_Compare"],
                r["Overlap_Area_sqm"],
            ])


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    setup()

    gdbs = find_gdbs(ROOT_DIR)

    if not gdbs:
        log.warning(f"No .gdb folders found under {ROOT_DIR}")
        return

    for gdb in gdbs:
        log.info(f"═══ GDB: {gdb}")

        all_records: list[dict] = []

        fcs = find_fcs(gdb)

        if not fcs:
            log.info("  No matching feature classes found.")

        for fc in fcs:

            records: list[dict] = []

            # ── 1: Detection and classification ─────────────────────────
            try:
                log.info(f"  ── FC: {fc}")

                pairs = detect_overlaps(fc)
                log.info(f"     Pairs found   : {len(pairs)}")

                records = build_records(gdb, fc, pairs)
                log.info(f"     Records built : {len(records)}")

                if records:
                    counts = Counter(r["OVERLAP_TYPE"] for r in records)
                    for otype, cnt in sorted(counts.items()):
                        log.info(f"       {otype:<12} {cnt}")

            except Exception:
                log.error(
                    f"  ERROR during detection/classification for {fc}:\n"
                    f"{traceback.format_exc()}"
                )

            all_records.extend(records)

            # ── 2: GDB output ───────────────────────────────────────────
            if ADD_FEATURE_CLASS and records:
                try:
                    write_fc(fc, gdb, Path(fc).name, records)
                except Exception:
                    log.error(
                        f"  ERROR writing GDB output for {fc}:\n"
                        f"{traceback.format_exc()}"
                    )

        write_excel(gdb, all_records)
        log.info(f"═══ Finished: {gdb}\n")

    log.info("All GDBs processed.")

if __name__ == "__main__":
    main()