# Results Approved Logo Upload Patch

เพิ่มระบบอัปโหลดโลโก้สำหรับเอกสาร Results Approved SEA Games Style

## เพิ่มในฐานข้อมูล
ตาราง `results_approved_setting` เพิ่มคอลัมน์:
- cover_main_logo_path
- cover_bottom_logo_1_path
- cover_bottom_logo_2_path
- cover_bottom_logo_3_path
- header_logo_1_path
- header_logo_2_path
- header_logo_3_path
- header_logo_4_path
- side_logo_path

ระบบมี migration สำหรับ SQLite และ PostgreSQL เพื่อเพิ่มคอลัมน์ในฐานข้อมูลเดิมอัตโนมัติ

## เพิ่มในหน้า Settings
หน้า `/events/<event_id>/results-approved/settings` เพิ่มช่องอัปโหลดโลโก้:
- โลโก้หลักบนปก
- โลโก้ล่างปก 1-3
- โลโก้หัวกระดาษ 1-4
- โลโก้มุมซ้ายในหน้าผล

รองรับไฟล์ png, jpg, jpeg, webp, gif

## การแสดงผล
หน้า Results Approved จะใช้โลโก้ที่อัปโหลด ถ้าไม่ได้อัปโหลดจะใช้โลโก้ค่าเริ่มต้นตามเอกสารตัวอย่าง

## DOCX
ไฟล์ DOCX จะใส่โลโก้หลักบนปกและโลโก้ล่างปก รวมถึงโลโก้หัวหน้า Officials ถ้าไฟล์เป็นชนิดที่ python-docx รองรับ
