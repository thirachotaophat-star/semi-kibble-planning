# semi-kibble-planning

เอกสารนี้สรุป **Business Rules ล่าสุด** สำหรับการวางแผน Semi Kibble โดยยืนยันว่าในรอบนี้
- อัปเดตเฉพาะกติกาและขั้นตอนคำนวณ (Algorithm)
- **ยังไม่เขียน VBA** ตามที่ร้องขอ

## Business Rules (Latest)

1. **MFG/EXP เดียวกัน** สามารถช่วยกันได้เฉพาะการทำยอดรวมให้ครบ **Case 240 PCS**
2. **ห้ามรวมคนละแถว** ที่เป็นคนละ **Shift/MC** เข้า **Outer bag เดียวกัน**
3. **Plan รายแถว** ต้องเป็นจำนวนที่หาร 240 ลงตัวเท่านั้น เช่น 240, 480, 720, 4560
4. ห้ามเกิด Plan รายแถวแบบ 4765 หรือ 8195 แม้ยอดรวมกลุ่มจะหาร 240 ลงตัว
5. เศษ PCS ที่ไม่ครบ 240 ในแต่ละแถว ให้ถือเป็น **Balance / Remainder** และ **ไม่เอาไป Plan**
6. ช่อง **Plan 1 ถึง Plan 9** ต้องแสดงเป็น **PCS** เหมือนเดิม
7. ช่อง **Plan Qty ด้านบน** ยังกรอกเป็น **Case** เหมือนเดิม

---

## Algorithm (อัปเดตตามกฎล่าสุด)

> หน่วยอ้างอิง:
> - 1 Case = 240 PCS
> - Plan รายแถวคำนวณและแสดงผลเป็น PCS

### Input ต่อแถว
- MFG
- EXP
- Shift
- MC
- Qty (PCS)

### ขั้นตอนคำนวณ

1. **จัดกลุ่มข้อมูลเพื่อคุมการช่วยกันของยอด**
   - กลุ่มหลักตาม `MFG + EXP` เพื่ออนุญาตให้ช่วยกันเฉพาะการครบเคส 240
   - แต่การจัดลง Outer bag ต้องรักษาขอบเขตแถวเดิม (ห้ามข้ามแถวที่ Shift/MC ต่างกัน)

2. **คำนวณ Plan รายแถวแบบ Floor to Case**
   - สำหรับแต่ละแถว ให้คำนวณ:
     - `planned_pcs_row = floor(qty_pcs_row / 240) * 240`
     - `remainder_pcs_row = qty_pcs_row - planned_pcs_row`
   - ผลลัพธ์ `planned_pcs_row` ต้องเป็น 0, 240, 480, 720, ... เท่านั้น
   - ค่า remainder เก็บเป็น Balance/Remainder และไม่ถูกส่งเข้า Plan

3. **ห้ามใช้วิธีเฉลี่ย/โยกจนแถวกลายเป็นเลขไม่หาร 240 ลงตัว**
   - แม้ผลรวมในกลุ่ม `MFG+EXP` จะลงตัว 240 แต่ถ้าแถวใดได้ค่า 4765 หรือ 8195 ให้ถือว่า **ผิดกฎ**
   - ระบบต้องไม่สร้างค่า Plan ลักษณะนี้

4. **กติกา Outer bag**
   - ห้ามรวมรายการคนละแถวที่เป็นคนละ Shift/MC ลง Outer bag เดียวกัน
   - การช่วยกันในระดับ MFG/EXP มีไว้เพื่อมองยอดรวมให้ครบเคสเท่านั้น ไม่ใช่เพื่อผสมแถวข้ามเงื่อนไข Shift/MC

5. **การแสดงผล Plan**
   - Plan 1 ถึง Plan 9: แสดงค่าเป็น PCS
   - Plan Qty ด้านบน: รับ/แสดงค่าเป็น Case

### Pseudocode (Reference)

```text
CASE_SIZE = 240

for each row in data_rows:
    planned_pcs_row = floor(row.qty_pcs / CASE_SIZE) * CASE_SIZE
    remainder_pcs_row = row.qty_pcs - planned_pcs_row

    row.plan_pcs = planned_pcs_row
    row.balance_pcs = remainder_pcs_row

    assert row.plan_pcs % CASE_SIZE == 0

# display
# - Plan1..Plan9 => PCS
# - Plan Qty (header/top) => CASE
# - Outer bag must not mix rows with different Shift/MC
```

## หมายเหตุ
- เอกสารนี้เป็น baseline ล่าสุดสำหรับนำไปเขียน VBA ในขั้นถัดไป
- หากต้องการ รอบต่อไปสามารถแตกเป็น Test Cases (Input/Expected Output) ได้ทันที
