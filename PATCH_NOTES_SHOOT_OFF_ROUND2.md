# Patch: Shoot-off รอบ 2 ไม่บันทึก / บันทึกแล้วไม่หาย

แก้ไขอาการกรอก Shoot-off รอบ 2 แล้วหน้า Overview ยังเหมือนไม่บันทึก หรือผลไม่ถูกนำไปตัดสิน

## จุดที่แก้

1. หน้า `tiebreak.html`
   - เพิ่ม `action` ให้ฟอร์มส่งกลับ URL ที่มี `round=2` และ `ids` ชัดเจน
   - เพิ่ม hidden field `round` และ `ids`
   - กันกรณี browser/redirect ทำ query string หาย แล้วระบบเอาคะแนน Shoot-off รอบ 2 ไปบันทึกผิดเป็นรอบ 1

2. Route `/events/<event_id>/tiebreak`
   - เปลี่ยนให้อ่าน `round` จาก `request.values` ได้ทั้ง GET และ POST
   - ทำให้ POST รอบ 2 บันทึกเป็น `round_no=2` แน่นอน

3. Route `/athletes/<athlete_id>/tiebreak`
   - เปลี่ยนให้อ่าน `round` จาก `request.values` ได้เช่นกัน

4. Logic รอบ 2
   - การตัดสิน Shoot-off รอบ 2 ใช้เฉพาะ `TieBreakEntry.round_no = 2`
   - ไม่เอา Shoot-off รอบ 1 มาบวกกับรอบ 2 แล้ว เพราะทำให้จำนวนเที่ยวพิเศษไม่เท่ากันและระบบยังขึ้น Shoot-off ซ้ำเหมือนไม่บันทึก

5. ปุ่ม Shoot-off
   - ส่งรายชื่อเฉพาะคนในกลุ่มที่ตีจบรอบนั้นแล้ว
   - กันการพ่วงคนเข้ารอบตรง/คนรอคิว/คนไม่เกี่ยวข้องเข้าไปในฟอร์ม Shoot-off รอบ 2

## ตรวจแล้ว

- `python3 -m py_compile app.py` ผ่าน
