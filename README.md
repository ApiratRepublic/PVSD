# PVSD
PVSD Repository is GIS Data Reviewer specifically for my friends at the Psychosis Veiled in Sacred Dreams.

## What the project does
รวบรวมไพธอนสคริปต์ที่ใช้ในการบริหารจัดการด้าน GIS 
ตรวจสอบและจัดการข้อมูลเบื้องต้นตามข้อกำหนด

## Why the project is useful
การจัดการกับเชปไฟล์หรือข้อมูลที่อยู่ใน GDB ไม่ใช่เรื่องยาก แต่ค่อนข้างวุ่นวาย เมื่อคำนึงถึงปริมาณข้อมูลที่มี จึงต้องใช้ไพธอน ช่วยจัดการตรวจสอบหรือการจัดการอัตโนมัติมาช่วยบ้าง
ซึ่งจากการทดสอบ gdb จำนวน 160 gdb ข้อมูลทั้งประเทศใช้เวลาตรวจสอบจาก check_required_featureclass.py เพียง 75 นาที 

## How users can get started with the project
ในโปรเจ็กต์นี้ ประกอบด้วยสคริปต์ดังต่อไปนี้
	1. check_required_featureclass.py อันนี้จะใช้หรือไม่ใช้ก็ได้ เป็นการตรวจสอบว่าใน แต่ละ gdb มีฟีเจอร์คลาสที่กำหนดหรือไม่ ถ้ามี เป็นจำนวนเท่าไหร่ (สร้างสคริปต์นี้ เนื่องจากพบว่าการสำเนา gdb อาจจะไม่สมบูรณ์)
	2. gdb_data_reviewer.py ใช้สำหรับตรวจสอบว่ามีฟีเจอร์คลาสครบตามที่กำหนดหรือไม่ (ข้อกำหนดต่าง ๆ ที่ตรวจสอบ จะอยู่ใน GDB Data Standard.md) วิธีใช้งานและตั้งค่าจะอยู่ใน gdb_data_reviewer.md


## Where users can get help with your project
หากมีคำแนะนำว่าควรตรวจสอบอะไรเพิ่ม หรือตรรกะในการตรวจสอบผิดพลาด กรุณาแจ้ง

## Who maintains to the project
Apirat Rattanapaiboon
