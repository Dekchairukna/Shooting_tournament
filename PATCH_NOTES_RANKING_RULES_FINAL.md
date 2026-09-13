# Ranking rules final patch

ปรับ logic ตามที่ตกลงหน้างาน:

## Round 1
- Top direct quota (ปกติ 1-4): TOTAL -> count(5) -> count(3) -> Shoot-off
- ถ้ายังเสมอก่อน Shoot-off แสดง Class ซ้ำได้ชั่วคราว เช่น 2,2 และขึ้น Shoot-off
- หลังพ้น direct quota: ใช้ TOTAL สำหรับ Class เท่านั้น
- TOTAL เท่ากัน = Class เท่ากันแบบ competition ranking เช่น 5,5,5,8
- การเลือกไป Round 2 ใช้คะแนน Total ของคนตรง cutoff row และรับทุกคนที่ Total เท่ากับคะแนน cutoff
- ไม่มี Shoot-off เพื่อแยก Class ตั้งแต่อันดับหลัง direct quota

## Round 2
- SUM = Round1 + Round2
- โควตาผ่านจาก Round 2 (ปกติ seed 5-8): SUM -> count(5) รวม -> count(3) รวม -> Shoot-off
- หลังพ้นโควตาผ่าน: SUM เท่ากัน = Class เท่ากันแบบ competition ranking

## Existing features retained
- Event-scoped court IDs / auto court IDs
- Court backend access guard
- Score audit log + editor signature
- Round 2 manual override
