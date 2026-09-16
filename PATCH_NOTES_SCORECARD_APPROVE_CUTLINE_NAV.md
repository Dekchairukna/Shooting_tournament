# Patch: Scorecard Approve + Live Cut Line + Round Navigation

- ย้ายการกด Approve ออกจากหน้า Overview
- Superadmin ต้องเปิด Scorecard ของนักกีฬาแล้วกด Approve จากหน้าคะแนน
- ถ้ามีลายเซ็นครบ 3 ฝ่าย ระบบยัง APPROVED อัตโนมัติตามเดิม
- Cut Line แสดงแบบ LIVE เมื่อมีผู้ตีจบครบอย่างน้อยเท่ากับโควตา
  - Round 1: โควตาผ่านตรง
  - Round 2: โควตาผ่านจาก Round 2 เข้า Bracket
- เมื่อรอบจบครบ Cut Line ใช้ผลสิทธิ์จริงของระบบ
- หัวตารางเปลี่ยนเป็น Qualification Shooting ROUND 1 / ROUND 2
- แยกปุ่ม ROUND 1 / ROUND 2 / Knockout เป็นกลุ่มเด่นจากปุ่มเครื่องมือ
- ปุ่ม Knockout เปลี่ยนชื่อตามค่าที่ตั้งใน Event: Round of 16 / Quarter Final / Semi Final
- หน้า Bracket ใช้หัวข้อให้ตรงกับรอบเริ่มต้น
- เปลี่ยนคำบน Scorecard จาก "กรรมการตัดสิน" เป็น "กรรมการยกคะแนน"
