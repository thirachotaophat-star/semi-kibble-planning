# Semi Kibble Packing Plan (FIFO) - System Design

เอกสารนี้สรุป **โครงสร้างระบบ (Architecture)** สำหรับทำงานใน Excel + VBA  
ตามเงื่อนไข Semi Kibble Packing Plan โดย **ยังไม่ลงสูตรคำนวณจริง** และยังไม่เขียนโค้ด VBA

---

## 1) Business Rules (เวอร์ชันล่าสุด)

1. 1 Outer bag = 5 PCS  
2. 1 Case = 48 Outer bag = 240 PCS  
3. Plan Qty ด้านบนกรอกเป็น Case  
4. Plan 1 ถึง Plan 9 ด้านล่างแสดงเป็น PCS  
5. ใช้ FIFO ตาม MFG / EXP  
6. MFG/EXP เดียวกันสามารถช่วยกันได้ **เฉพาะการทำยอดรวมให้ครบ Case 240 PCS**  
7. ห้ามรวม Shift/MC คนละแถวเข้า Outer bag เดียวกัน  
8. Plan รายแถวต้องเป็นจำนวนที่หาร 240 ลงตัวเท่านั้น (เช่น 240, 480, 720, 4560)  
9. ห้ามเกิด Plan รายแถวแบบ 4765 หรือ 8195 แม้ยอดรวมกลุ่มจะหาร 240 ลงตัว  
10. เศษ PCS ที่ไม่ครบ 240 ในแต่ละแถวให้ถือเป็น Balance / Remainder (ไม่เอาไป Plan)  
11. ต้องคำนวณ Balance PCS / Case / Outer bag  
12. ต้องรองรับหลาย Customer  
13. ต้องมีปุ่ม Auto Allocate  
14. ใช้กับ Excel ได้จริง

---

## 2) โครงสร้างไฟล์ Excel (Workbook Structure)

เสนอ 6 ชีตหลัก:

1. **`Setup`**  
   - เก็บค่าคงที่ระบบ (เช่น PCS_PER_OUTER = 5, OUTER_PER_CASE = 48, PCS_PER_CASE = 240)
   - เก็บรายการลูกค้า, Shift, MC, Product mapping
   - ใช้เป็นแหล่งข้อมูล Data Validation

2. **`Stock`** (วัตถุดิบ/สินค้าคงคลังสำหรับจัดแผน)  
   - 1 แถว = 1 lot ต่อ Shift/MC/Customer/Product/MFG/EXP
   - ใช้เป็นแหล่งข้อมูลหลักของ FIFO

3. **`Plan_Header`** (แผนด้านบน, หน่วย Case)  
   - ผู้ใช้กรอกเป้าเป็น Case ต่อ Customer / Product / Shift / MC / วันที่
   - ใช้เป็น input ให้ Auto Allocate

4. **`Plan_Detail`** (แผนด้านล่าง, หน่วย PCS: Plan 1 ถึง Plan 9)  
   - ผลจากการ Allocate แบบ lot-by-lot
   - 1 แถวต้องลงตัว 240 PCS
   - แยกแถวตาม Shift/MC (ไม่ merge)

5. **`Balance`**  
   - สรุปยอดคงเหลือหลัง allocate ทั้งหน่วย PCS / Outer / Case
   - แสดงก่อน/ใช้ไป/คงเหลือ

6. **`Log`**  
   - บันทึกการกด Auto Allocate (เวลา, user, run id, จำนวนแถวที่ allocate)
   - ใช้ trace และตรวจสอบย้อนหลัง

---

## 3) โครงสร้างข้อมูลแต่ละชีต (Columns Blueprint)

### 3.1 Sheet: `Stock`

คอลัมน์แนะนำ:
- StockID (unique)
- AsOfDate
- Customer
- ProductCode
- Shift
- MC
- MFG_Date
- EXP_Date
- Qty_PCS
- Qty_Outer (แสดงผล)
- Qty_Case (แสดงผล)
- LotStatus (Available/Hold)

หมายเหตุ:
- **FIFO key:** `MFG_Date` ก่อน แล้ว `EXP_Date` เป็น tie-breaker  
- เพื่อรองรับกติกาใหม่ “MFG/EXP เดียวกันช่วยกันได้เฉพาะครบ 240 PCS”  
  ให้ระบบรวมได้เฉพาะ lot ที่ MFG/EXP เท่ากัน ภายใต้ Customer/Product/Shift/MC เดียวกัน และใช้เพื่อปิดก้อน Case เท่านั้น

### 3.2 Sheet: `Plan_Header` (กรอกเป็น Case)

คอลัมน์แนะนำ:
- PlanID (run-level grouping)
- PlanDate
- Customer
- ProductCode
- Shift
- MC
- PlanQty_Case (input)
- PlanQty_PCS (display = PlanQty_Case × 240)
- Status (Draft/Allocated/Partial/Error)

### 3.3 Sheet: `Plan_Detail` (ผลลัพธ์เป็น PCS)

คอลัมน์แนะนำ:
- PlanLineID
- PlanID
- PlanDate
- Customer
- ProductCode
- Shift
- MC
- SourceStockID
- MFG_Date
- EXP_Date
- Allocated_PCS
- Allocated_Outer
- Allocated_Case
- RuleCheck_240 (Pass/Fail)
- Remark

เงื่อนไขบังคับ:
- `Allocated_PCS` ในแต่ละแถวต้องหาร 240 ลงตัวเท่านั้น
- ห้ามรวม Shift/MC คนละแถวใน Outer เดียวกันหรือ Case เดียวกัน
- PCS ที่เหลือไม่ครบ 240 ให้ไป `Balance` เท่านั้น (ไม่เขียนลง Plan_Detail เป็นแผน)

### 3.4 Sheet: `Balance`

คอลัมน์แนะนำ:
- SnapshotDateTime
- Customer
- ProductCode
- Shift
- MC
- Opening_PCS
- Used_PCS
- Balance_PCS
- Balance_Outer
- Balance_Case

### 3.5 Sheet: `Log`

คอลัมน์แนะนำ:
- RunID
- RunDateTime
- UserName
- PlanRowsInput
- PlanRowsAllocated
- ErrorCount
- Message

---

## 4) กติกา Allocation (Logical Flow)

### Step A: Validate Input
- Plan_Header ต้องมี Customer/Product/Shift/MC/PlanQty_Case
- PlanQty_Case ต้องเป็นจำนวนเต็มบวก
- ตรวจว่ามี stock ตาม key ที่ต้องใช้
- ตรวจว่า stock ที่จะนำมา Plan ไม่ผสมข้าม Shift/MC

### Step B: Convert Unit
- PlanQty_PCS = PlanQty_Case × 240

### Step C: FIFO Allocation
- Filter Stock ตาม Customer + Product + Shift + MC
- Sort ตาม MFG_Date ASC, EXP_Date ASC
- Allocate ทีละ lot โดยคิดเป็นก้อน Case (240 PCS) เท่านั้น
- หากเจอ MFG/EXP เดียวกันหลายแถว ให้หยิบช่วยกันได้เฉพาะเพื่อปิดก้อน 240 PCS
- ห้ามเอา lot ต่าง Shift/MC มาช่วยกันปิดก้อน 240 PCS เดียวกัน
- PCS คงเหลือที่ไม่ถึง 240 หลังจบแต่ละแถว ให้ส่งไป Remainder/Balance

### Step D: Rule Enforcement
- แถวผลลัพธ์ใน Plan_Detail ต้องเป็นจำนวนที่หาร 240 ลงตัวเท่านั้น
- บล็อกค่าเช่น 4765 หรือ 8195 แม้ยอดรวมทั้งกลุ่มจะหาร 240 ลงตัว
- ห้าม merge ข้าม Shift/MC ทั้งในระดับ Outer และระดับ Case
- PCS เศษที่ไม่ครบ 240 ต้องถูกจัดเป็น Balance/Remainder

### Step E: Write Output
- เขียน Plan_Detail เฉพาะรายการที่หาร 240 ลงตัว
- อัปเดต Balance รวมเศษ Remainder ที่ไม่ครบ 240
- บันทึก Log

---

## 5) โครงสร้าง VBA Modules (ยังไม่เขียนโค้ด)

แยกโมดูลให้ดูแลง่าย:

1. **`modConstants`**
   - ค่าคงที่หน่วย: 5, 48, 240
   - ชื่อชีต/ชื่อ table

2. **`modValidation`**
   - ตรวจ input ก่อน run
   - ตรวจเงื่อนไขหาร 240 ลงตัวต่อแถว และบล็อกค่าต้องห้าม (เช่น 4765, 8195)
   - ตรวจไม่ให้ผสม Shift/MC คนละแถวใน Outer/Case เดียวกัน

3. **`modFIFOEngine`**
   - ฟังก์ชันค้นหา stock ตาม key
   - ฟังก์ชัน sort FIFO (MFG/EXP)
   - ฟังก์ชันจัดสรรปริมาณแบบก้อน 240 PCS
   - ฟังก์ชันปิดก้อนด้วย MFG/EXP เดียวกัน (เท่านั้น)

4. **`modPlanWriter`**
   - เขียนผลลง Plan_Detail
   - สรุป Balance

5. **`modUI`**
   - ปุ่ม `Auto Allocate`
   - ปุ่ม `Reset Draft` (ถ้าต้องการ)
   - แสดง message box ผลการทำงาน

6. **`ThisWorkbook` / `Sheet events`**
   - Event สำหรับ lock โครงสร้างและป้องกันกรอกผิด

---

## 6) UI/UX ใน Excel

- ด้านบน: ส่วน `Plan_Header` (กรอก Case)  
- ปุ่มใหญ่: **Auto Allocate**  
- ด้านล่าง: ตาราง `Plan_Detail` (โชว์ PCS พร้อม lot ที่ถูกใช้)  
- ด้านข้าง/อีกชีต: `Balance` สรุปคงเหลือ

ข้อแนะนำเพิ่ม:
- ใช้ Excel Table (ListObject) ทุกชีตหลัก
- ใช้ Data Validation dropdown สำหรับ Customer/Shift/MC/Product
- ใส่ Conditional Formatting สำหรับ Error rows

---

## 7) Error Handling ที่ควรมี

- Stock ไม่พอสำหรับแผน (แจ้งขาดกี่ PCS / กี่ Case)
- พบ output line ที่ไม่หาร 240 ลงตัว (flag และ rollback run)
- พบการผสม Shift/MC ใน Outer/Case เดียวกัน (block ก่อน allocate)
- พบเศษไม่ครบ 240 ให้ลง Balance/Remainder อัตโนมัติ
- ข้อมูลวันที่ MFG/EXP ไม่ถูกต้อง

---

## 8) ขอบเขต Version 1 (MVP)

MVP ควรทำให้ครบ:
- รองรับหลาย Customer
- FIFO ตาม MFG/EXP
- แยก Shift/MC ชัดเจน
- Auto Allocate ปุ่มเดียวจบ
- รายงาน Balance ครบ 3 หน่วย (PCS/Outer/Case)

สิ่งที่ค่อยเพิ่มใน V2:
- รองรับ re-allocation บางส่วน
- รองรับ priority ลูกค้า
- dashboard KPI (Fill Rate, Aging lot)
