# APPROVED row colour – final fix

ตรวจโค้ด Overview ทั้งชุดแล้วพบสาเหตุหลักว่า CSS legacy ของแถว `status-finished` และกลุ่ม qualification ใช้ selector ที่มี specificity สูง (มี `:not(...)` หลายชั้น) และ `!important` ทำให้ rule สีสถานะที่เพิ่มทีหลังไม่สามารถ override ได้ จึงเห็นแถวสีส้ม/สีเดิมทั้งที่ป้ายขึ้น APPROVED.

## การแก้ไข
- APPROVED = `bypass_signed` (ผู้มีสิทธิ์ approve) **หรือ** มีลายเซ็นจริงครบ 3 ฝ่าย
  - ผู้บันทึกคะแนน (`recorder_signature`)
  - นักกีฬา (`athlete_signature`)
  - กรรมการยกคะแนน (`referee_signature`)
- การพิมพ์ชื่ออย่างเดียวไม่ทำให้เป็น APPROVED
- ส่งค่า `approved` ทั้งตอน render หน้าและ API `/overview-data`
- หน้า Overview เพิ่มคลาส `approved-row` และข้อความ `APPROVED`
- Live refresh รักษา `approved-row` ไว้ ไม่ล้างกลับเป็นสีเดิม
- เพิ่ม final CSS override ที่ specificity สูงกว่า legacy selectors

## สีสุดท้าย
- รอตี: เทา `#f1f5f9`
- กำลังตี: เหลือง `#fef3c7`
- ตีแล้วแต่ยังไม่ approve: เขียว `#dcfce7`
- APPROVED: ฟ้า `#dbeafe`, เส้นซ้าย `#2563eb`, ป้าย APPROVED น้ำเงิน
- ช่องคะแนนที่กรอกระหว่างกำลังตี: แดงอ่อน `#fee2e2`

Cut Line ยังคงเป็นเส้นแบ่งและไม่ใช้สีพื้นแถวเพื่อไม่ให้ชนกับสถานะ.
