# Overview row color clean reset

- ลบกฎสีพื้นทั้งแถวเดิมที่ผูกกับ waiting / active / qualified / R2 / eliminated / shoot-off
- ไม่เปลี่ยน cursor บอกตำแหน่งที่กำลังตี และไม่เปลี่ยน Cut Line
- ตีครบ/ส่งคะแนนแล้วแต่ยังไม่ APPROVED = เขียวทั้งแถว (`#dcfce7`)
- APPROVED = ฟ้าทั้งแถว (`#dbeafe`) + ป้าย APPROVED สีน้ำเงิน
- รอตีและกำลังตีไม่บังคับสีพื้นทั้งแถว
- APPROVED = admin bypass/approve เดิม หรือมีลายเซ็นจริงครบ 3 ฝ่าย
