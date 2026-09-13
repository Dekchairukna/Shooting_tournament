# Patch: 8 Courts / 48 Athletes Operator Mode

เพิ่มโหมดเจ้าหน้าที่สำหรับการแข่งขัน Shooting ที่มี 8 สนาม / 48 นักกีฬา / เจ้าหน้าที่ 12 คน โดยต่อยอดจากระบบเดิมและไม่เปลี่ยนโครงคะแนน

## สิ่งที่เพิ่ม
- Role `court` สำหรับเจ้าหน้าที่สนาม ใช้ username รูปแบบ `court01` ถึง `court08` เพื่อผูกกับสนามอัตโนมัติ
- Role `checker` สำหรับเจ้าหน้าที่ตรวจผล/ดู Control
- Role `control` สำหรับผู้ควบคุมกลาง
- หน้าคิวสนาม `/events/<event_id>/court` แสดงนักกีฬา 6 คนของสนามนั้นตาม `lane_order`
- หน้า Control `/events/<event_id>/control` แสดงทั้ง 8 สนาม พร้อมสถานะ รอคิว / กำลังยิง / จบแล้ว
- ตาราง `scorecard_lock` ป้องกันผู้ใช้ต่างคนคีย์ Scorecard นักกีฬาคนเดียวกันในรอบเดียวกัน
- Court account เปิด Scorecard ได้เฉพาะนักกีฬาที่ `lane_no` ตรงกับสนามของบัญชี
- Auto-save ตรวจ Lock ทุกครั้งก่อนบันทึก
- Superadmin สามารถสร้าง role ใหม่จากหน้า Users ได้

## การตั้งผู้ใช้แนะนำ
สร้างด้วย superadmin:
- court01 ... court08 (role: court)
- checker01, checker02 (role: checker)
- control (role: control)

ตั้งรหัสผ่านจริงของหน่วยงานเอง ห้ามใช้รหัสตัวอย่างร่วมกันในวันแข่งขัน

## PostgreSQL / Railway
`db.create_all()` จะสร้าง `scorecard_lock` และ `ensure_schema()` จะสร้าง sequence/default สำหรับ PostgreSQL เดิมโดยอัตโนมัติ ไม่ล้างข้อมูลเดิม

## หมายเหตุการทดสอบ
ผ่าน `python -m py_compile app.py` แล้ว แต่ environment ที่ใช้จัดแพตช์ไม่มี dependency Flask และไม่มีอินเทอร์เน็ตให้ติดตั้ง จึงยังไม่ได้รัน Flask integration test ใน container นี้ ควรทดสอบบน local/preview Railway ก่อนใช้วันจริง
