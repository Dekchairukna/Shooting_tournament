# Results Approved SEA Games Style Fix

สิ่งที่แก้ในชุดนี้

1. เปลี่ยนหน้า Results Approved ให้เป็นรูปแบบ SEA Games Style ตามเอกสารตัวอย่างมากขึ้น
   - Cover แบบ RESULTS / APPROVED
   - หน้า Officials / Umpire
   - หน้า Event Cover พร้อมรายชื่อ COUNTRY/AFFILIATION
   - Name Lists แบบ FAMILY NAME / GIVEN NAME
   - Qualification Round 1 summary
   - Qualification Shooting detail แนวนอนแบบ Atelier 1-5 และ 6M/7M/8M/9M/Tot.
   - Qualification Round 2 summary
   - Qualification Shooting Round 2 detail
   - Semifinal / Final
   - Ranking Result

2. เพิ่มปุ่มตั้งค่า Results Approved
   - route: /events/<event_id>/results-approved/settings
   - ใช้ใส่ชื่อประธาน, Technical Delegate, รายชื่อ Umpire, หัวคอลัมน์ COUNTRY/AFFILIATION/SCHOOL/TEAM, วันที่ และข้อความ APPROVED

3. เพิ่มตารางฐานข้อมูลใหม่
   - ResultsApprovedSetting
   - ผูกกับ Event ทีละรายการ

4. เพิ่มโลโก้จากเอกสารตัวอย่างไว้ที่
   - static/results_approved_assets/

5. อัปเดต DOCX export ให้ใช้ข้อมูล settings และโครงใกล้เอกสารตัวอย่างมากขึ้น

ตรวจสอบแล้ว
- python3 -m py_compile app.py ผ่าน
- ตรวจ syntax Jinja ของ results_approved.html / results_approved_settings.html / overview.html ผ่าน
