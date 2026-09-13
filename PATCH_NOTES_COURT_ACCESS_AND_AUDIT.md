# Court Access + Audit patch

## เพิ่มในเวอร์ชันทดลองนี้

1. สร้างบัญชีสนามจากหน้า Dashboard โดย superadmin
   - ระบุเลขสนาม + รหัสผ่าน
   - username สร้างอัตโนมัติเป็น `court01`, `court02`, ...
   - สร้างซ้ำเลขเดิม = เปลี่ยนรหัสผ่านบัญชีนั้น

2. เพิ่ม role `court` และ `user.court_no`
   - PostgreSQL/SQLite เดิม migrate เพิ่มคอลัมน์อัตโนมัติ

3. จำกัดหน้ารวมตามสนาม
   - courtXX เห็นเฉพาะนักกีฬาที่ถูกจัดอยู่สนาม XX
   - ใช้ `display_lane_no` ในรอบ 2 เพื่อรองรับการจัดสนามใหม่ของรอบ 2
   - realtime `/overview-data` ถูกกรองด้วยเช่นกัน จึงไม่ดึงข้อมูลสนามอื่นกลับเข้าหน้า

4. ป้องกันคีย์ข้ามสนามฝั่ง Backend
   - Scorecard GET/POST ตรวจเลขสนาม
   - Autosave API ตรวจเลขสนามและตอบ 403 หากข้ามสนาม
   - เปลี่ยน URL เองก็ไม่สามารถคีย์สนามอื่นได้

5. Audit log หลังผลเซ็นยืนยันแล้ว
   - เพิ่มตาราง `score_edit_log`
   - เก็บค่าเดิม/ค่าใหม่, สถานี, ระยะ, ผู้แก้, เวลา, สนาม, เหตุผล, ลายเซ็นผู้แก้
   - ถ้าผลเซ็นแล้วและมีการเปลี่ยนคะแนน ระบบบังคับลายเซ็นผู้แก้ก่อน Autosave
   - มีหน้า “ประวัติการแก้คะแนน” จาก Scorecard

## การทดสอบก่อน deploy

- Python syntax (`py_compile`) ผ่าน
- Jinja template parse ทุกไฟล์ผ่าน
- Environment ของผู้ช่วยไม่มี Flask package และไม่มี internet จึงยังไม่สามารถรัน Flask integration test จริงใน container นี้ได้

## Checklist ทดสอบบนเครื่อง/Railway

1. login superadmin -> Dashboard -> สร้างสนาม 1 รหัสทดสอบ
2. logout -> login court01
3. เปิด event รอบ 1 -> ต้องเห็นเฉพาะ lane 1
4. เปิด URL Scorecard นักกีฬา lane 2 โดยตรง -> ต้องถูกปฏิเสธ
5. คีย์นักกีฬา lane 1 -> Autosave ต้องทำงาน
6. รอบ 2 -> court01 ต้องเห็นเฉพาะ `display_lane_no = 1`
7. เซ็นครบ 3 ฝ่าย -> ลองแก้คะแนน -> ต้องเด้งให้เซ็นผู้แก้ -> หลังเซ็นจึงบันทึกได้
8. เปิดประวัติการแก้ -> ต้องเห็น old/new + username + ลายเซ็น
