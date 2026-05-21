# semi-kibble-planning

เอกสารนี้ออกแบบ **FIFO Allocation Engine** แบบละเอียดตาม Requirement ล่าสุด
- รอบนี้เป็น **Design + Algorithm + Pseudocode** เท่านั้น
- **ยังไม่เขียน VBA**

---

## 1) Requirement Mapping (แปลงความต้องการเป็นกติกาเครื่องยนต์)

### หน่วยมาตรฐาน
- `1 Outer bag = 5 PCS`
- `1 Case = 48 Outer bag = 240 PCS`
- ดังนั้นทุกการ Allocate ที่ลงแผนจริงต้องลงท้ายเป็น **Case เต็มเท่านั้น**

### กติกาหลัก
1. Plan Qty ด้านบนกรอกเป็น **Case**
2. ระบบแปลง Plan Qty เป็น **PCS** อัตโนมัติ
3. Plan1–Plan9 แสดงเป็น **PCS**
4. ห้ามใช้เศษ `1–4 PCS` (เพราะไม่ครบ 1 bag)
5. Plan ต่อแถวต้องลงตัว `240 PCS`
6. การช่วยกันข้ามแถวทำได้เฉพาะกรณี `MFG/EXP เดียวกัน` เพื่อเติมให้ครบ `48 bag (240 PCS)`
7. ห้ามผสม bag เดียวกันข้าม `Shift/MC` เด็ดขาด
8. ถ้าแถวใดมี `46 bag` ต้องหาเพิ่ม `2 bag` จากแถวถัดไปที่อนุญาต
9. ถ้าแถวใดมี `47 bag` ต้องหาเพิ่ม `1 bag` จากแถวถัดไปที่อนุญาต
10. ถ้าเติมไม่ได้จนถึง `48 bag` ให้ **ไม่ allocate** ก้อนนั้น
11. Balance ต้องคำนวณใหม่อัตโนมัติทุกครั้งหลัง allocate
12. ต้องรองรับข้อมูลระดับหลายพันแถว (O(n log n) ได้, หลีกเลี่ยง nested loop แบบ n^2 โดยไม่จำเป็น)

---

## 2) แนวคิดการออกแบบ Allocation Engine

### 2.1 ภาพรวม
Engine ทำงานแบบ 2 ชั้น:
1. **FIFO Layer**: เรียง stock ตามคิวก่อน-หลัง (เช่น MFG ก่อน, แล้ว EXP, แล้วลำดับรับเข้า)
2. **Eligibility Layer**: ตรวจว่าแถวที่จะช่วยเติมกันได้ต้องผ่านกติกา
   - ช่วยกันได้เฉพาะ `MFG/EXP เดียวกัน`
   - แต่ bag ที่ย้ายจากแถวผู้ให้ไปแถวผู้รับ ต้องไม่ทำให้เกิดการผสม Shift/MC ใน bag เดียวกัน

> จุดสำคัญ: การ “ช่วยเติม” คือการโอน **จำนวน bag เต็ม** (ทีละ 1 bag = 5 PCS) จากแถวที่อนุญาต เพื่อให้แถวเป้าหมายครบ 48 bag

### 2.2 คีย์การจัดกลุ่ม
- กลุ่มหลักเพื่อช่วยเติม: `group_key = (MFG, EXP)`
- เอกลักษณ์แถว: `row_id`
- ข้อจำกัดการผสม: `shift_mc_key = (Shift, MC)`

แนวปฏิบัติที่ปลอดภัย:
- อนุญาตการช่วยกันเฉพาะแถวที่มี `shift_mc_key` เดียวกันก่อน (strict mode)
- ถ้าธุรกิจต้องการตีความว่า “ห้ามรวมเฉพาะใน bag เดียวกัน แต่โอนเป็น bag เต็มได้” ให้ยังคงโอนเป็นหน่วย bag เต็มเท่านั้น และเก็บ trace ชัดเจนว่า bag ไหนมาจากแถวไหน

---

## 3) Data Structure (รองรับหลายพันแถว)

## 3.1 โครงสร้างข้อมูลต่อแถว

```text
RowStock {
  row_id: string
  fifo_seq: int                # ลำดับ FIFO
  mfg_date: date
  exp_date: date
  shift: string
  mc: string

  qty_pcs_raw: int             # ปริมาณดิบในแถว
  qty_bag_raw: int             # floor(qty_pcs_raw / 5)
  qty_pcs_reject_1_4: int      # qty_pcs_raw % 5 (ห้ามใช้)

  alloc_bag: int               # bag ที่ถูกเอาไป allocate
  alloc_pcs: int               # alloc_bag * 5

  balance_bag: int             # qty_bag_raw - alloc_bag
  balance_pcs: int             # balance_bag * 5
  balance_reject_pcs: int      # qty_pcs_reject_1_4
}
```

## 3.2 โครงสร้างสำหรับการจัดคิวตามกลุ่ม

```text
GroupBucket {
  group_key: (mfg_date, exp_date)
  rows_fifo: deque<RowStock>   # เรียงตาม fifo_seq
}
```

## 3.3 โครงสร้างแผน Plan1..Plan9

```text
PlanSlot {
  slot_no: 1..9
  target_case: int             # รับจาก Plan Qty ด้านบน หรือ logic แจกจ่ายภายใน
  target_pcs: int              # target_case * 240

  allocated_case: int
  allocated_pcs: int

  alloc_lines: list<AllocationLine>
}

AllocationLine {
  slot_no: int
  group_key: (mfg_date, exp_date)
  receiver_row_id: string
  donor_row_id: string
  move_bag: int                # หน่วย bag เท่านั้น
  move_pcs: int                # move_bag * 5
  fifo_seq_from: int
}
```

---

## 4) Allocation Flow (Step-by-Step)

## Step 0: Pre-validation
- ตรวจว่าค่าที่รับเข้าเป็นเลขไม่ติดลบ
- แปลง Plan Qty (Case) -> PCS
  - `plan_qty_case_total`
  - `plan_qty_pcs_total = plan_qty_case_total * 240`
- ถ้าแถวใด `qty_pcs_raw < 5` จะไม่มี bag ที่ใช้ได้

## Step 1: Normalize หน่วย
สำหรับทุกแถว:
- `qty_bag_raw = floor(qty_pcs_raw / 5)`
- `qty_pcs_reject_1_4 = qty_pcs_raw % 5`  → เก็บเป็นเศษห้ามใช้

## Step 2: Build FIFO Group
- จัดกลุ่มตาม `(MFG, EXP)`
- ภายในกลุ่มเรียง FIFO ตาม `fifo_seq` (หรือเวลารับเข้า)

## Step 3: ทำงานทีละกลุ่มแบบ FIFO
ในแต่ละ `GroupBucket`:
1. ดึงแถว FIFO ตัวแรกเป็นผู้รับ (`receiver`)
2. หา `receiver_available_bag`
3. ถ้า `receiver_available_bag >= 48`:
   - allocate ได้ทีละ `48 bag`
   - จำนวน case จากแถวนี้ = `receiver_available_bag // 48`
4. ถ้า `receiver_available_bag = 46 หรือ 47`:
   - คำนวณ `need_bag = 48 - receiver_available_bag` (2 หรือ 1)
   - พยายามดึงจากแถวถัดไปใน FIFO ที่ **ผ่าน eligibility**
   - ถ้าดึงครบ -> allocate 48 bag ให้ receiver
   - ถ้าดึงไม่ครบ -> ห้าม allocate ก้อนนี้
5. ถ้า `receiver_available_bag` เป็นค่าอื่นที่น้อยกว่า 48 (เช่น 45, 44, ...)
   - ตาม requirement นี้ให้ไม่ทำ partial combine ข้ามหลายแถวแบบเปิดกว้าง
   - จึงถือเป็นคงค้างไว้ก่อน (balance)

## Step 4: Eligibility Check ก่อนเติม bag
ผู้ให้ (`donor`) ต้องผ่านทั้งหมด:
- `donor.group_key == receiver.group_key` (MFG/EXP เดียวกัน)
- ไม่ทำให้ผิดกฎ Shift/MC (แนะนำ strict: `donor.shift==receiver.shift && donor.mc==receiver.mc`)
- donor ต้องมี bag คงเหลือพอ
- การโอนต้องเป็นจำนวนเต็ม bag เท่านั้น

## Step 5: Commit Allocation Transaction
เมื่อเติมครบ 48 bag:
- บันทึก AllocationLine ทุกการโอน
- หัก donor / receiver ตาม bag ที่ถูกใช้
- เพิ่มยอด allocate ของ plan slot
- อัปเดตยอดคงเหลือ (balance) ทันที

## Step 6: Stop Condition
หยุดเมื่ออย่างใดอย่างหนึ่งเกิดขึ้น:
- allocate ครบตาม `plan_qty_pcs_total`
- stock ที่ผ่านกติกาไม่พอ

## Step 7: Recalculate Balance (Auto)
หลังจบงาน:
- `balance_bag = qty_bag_raw - alloc_bag`
- `balance_pcs = balance_bag * 5`
- `balance_reject_pcs = qty_pcs_reject_1_4`

---

## 5) Detailed Pseudocode

```text
CONST PCS_PER_BAG = 5
CONST BAG_PER_CASE = 48
CONST PCS_PER_CASE = 240

function run_allocation(plan_qty_case_total, rows, plan_slots):
    # ------------------------
    # A) PREPARE
    # ------------------------
    plan_qty_pcs_total = plan_qty_case_total * PCS_PER_CASE
    required_case_total = plan_qty_pcs_total / PCS_PER_CASE

    for row in rows:
        row.qty_bag_raw = floor(row.qty_pcs_raw / PCS_PER_BAG)
        row.qty_pcs_reject_1_4 = row.qty_pcs_raw % PCS_PER_BAG
        row.alloc_bag = 0
        row.alloc_pcs = 0

    groups = group rows by (row.mfg_date, row.exp_date)
    for each group in groups:
        sort group.rows by row.fifo_seq asc

    allocated_case_total = 0
    allocation_log = []

    # ------------------------
    # B) FIFO ALLOCATION BY GROUP
    # ------------------------
    for each group in groups in FIFO order:
        i = 0
        while i < len(group.rows) and allocated_case_total < required_case_total:
            receiver = group.rows[i]
            receiver_avail_bag = receiver.qty_bag_raw - receiver.alloc_bag

            # 1) direct full cases
            while receiver_avail_bag >= BAG_PER_CASE and allocated_case_total < required_case_total:
                commit_case_from_single_row(receiver, group, allocation_log)
                allocated_case_total += 1
                receiver_avail_bag -= BAG_PER_CASE

            # 2) top-up only 46/47 bag
            if allocated_case_total < required_case_total and (receiver_avail_bag == 46 or receiver_avail_bag == 47):
                need_bag = BAG_PER_CASE - receiver_avail_bag   # 2 or 1
                success = try_topup_from_next_rows(
                    group_rows = group.rows,
                    receiver_index = i,
                    need_bag = need_bag,
                    receiver = receiver,
                    allocation_log = allocation_log
                )

                if success:
                    # now receiver has exactly +need_bag => full 48 bag
                    commit_case_after_topup(receiver, group, allocation_log)
                    allocated_case_total += 1
                else:
                    # cannot top-up, do not allocate this chunk
                    pass

            i += 1

    # ------------------------
    # C) BUILD PLAN SLOT OUTPUT
    # ------------------------
    # Plan1..Plan9 show PCS
    # distribute allocated cases into slots by defined policy
    distribute_allocation_to_plan_slots(plan_slots, allocation_log)

    # ------------------------
    # D) RECALC BALANCE
    # ------------------------
    for row in rows:
        row.alloc_pcs = row.alloc_bag * PCS_PER_BAG
        row.balance_bag = row.qty_bag_raw - row.alloc_bag
        row.balance_pcs = row.balance_bag * PCS_PER_BAG
        row.balance_reject_pcs = row.qty_pcs_reject_1_4

    return {
      allocated_case_total,
      allocated_pcs_total = allocated_case_total * PCS_PER_CASE,
      rows,
      allocation_log,
      plan_slots
    }


function try_topup_from_next_rows(group_rows, receiver_index, need_bag, receiver, allocation_log):
    remaining = need_bag

    for j in range(receiver_index + 1, len(group_rows)):
        donor = group_rows[j]

        # eligibility
        if donor.mfg_date != receiver.mfg_date or donor.exp_date != receiver.exp_date:
            continue

        # strict policy for Shift/MC
        if donor.shift != receiver.shift or donor.mc != receiver.mc:
            continue

        donor_avail_bag = donor.qty_bag_raw - donor.alloc_bag
        if donor_avail_bag <= 0:
            continue

        take_bag = min(donor_avail_bag, remaining)

        # transfer as whole bag only
        donor.alloc_bag += take_bag
        receiver.alloc_bag += take_bag
        remaining -= take_bag

        allocation_log.append({
          type: "TOPUP_TRANSFER",
          receiver_row_id: receiver.row_id,
          donor_row_id: donor.row_id,
          move_bag: take_bag,
          move_pcs: take_bag * PCS_PER_BAG
        })

        if remaining == 0:
            return true

    # rollback if not enough
    rollback_transfers_for_receiver(receiver.row_id, allocation_log, group_rows)
    return false
```

---

## 6) ตัวอย่างการทำงานตาม Requirement

### กรณี A: 46 bag
- receiver มี 46 bag
- ระบบตั้ง `need_bag = 2`
- ค้นหา donor แถวถัดไป (FIFO) ที่ MFG/EXP เดียวกัน และ Shift/MC ผ่าน
- ถ้าพบครบ 2 bag -> allocate ได้ 48 bag (240 PCS)
- ถ้าพบไม่ครบ -> **ไม่ allocate**

### กรณี B: 47 bag
- receiver มี 47 bag
- ระบบตั้ง `need_bag = 1`
- ถ้าหาได้ 1 bag ตามเงื่อนไข -> allocate ได้
- ถ้าหาไม่ได้ -> **ไม่ allocate**

### กรณี C: donor คนละ Shift/MC
- ถึงแม้ MFG/EXP ตรงกัน แต่ Shift/MC ไม่ตรง
- ระบบต้องข้าม donor แถวนั้นทันที
- ถ้าทำให้เติมไม่ครบ 48 bag -> ไม่ allocate

---

## 7) Complexity และ Performance

- Group + Sort: โดยรวมประมาณ `O(n log n)`
- Allocation scan ภายในกลุ่ม: ใกล้ `O(n)` ถึง `O(n + transfer_count)` ถ้าใช้ pointer/deque ดี
- รองรับหลายพันแถวได้สบายในระดับ worksheet engine
- ควรหลีกเลี่ยงการอ่าน/เขียน cell ทีละช่องซ้ำๆ (เมื่อไปเขียน VBA ให้ใช้ array in-memory)

---

## 8) Output Contract ที่แนะนำ

1. **Plan Header**
   - Input: `PlanQtyCase`
   - Derived: `PlanQtyPCS = PlanQtyCase * 240`

2. **Plan1..Plan9 (PCS)**
   - แสดงเฉพาะ PCS
   - ทุกค่าเป็นหลายเท่าของ 240

3. **Row-level Result**
   - `AllocatedPCS`
   - `BalancePCS`
   - `RejectPCS(1-4)`
   - `TopupIn/TopupOut` (trace ได้ว่าไป-มาจาก row ไหน)

4. **Validation Flags**
   - `FLAG_NO_TOPUP_SOURCE`
   - `FLAG_SHIFT_MC_MISMATCH`
   - `FLAG_INSUFFICIENT_BAG_FOR_CASE`

---

## 9) สรุปสั้น

Engine นี้จะ allocate แบบ FIFO โดยยึดหน่วย bag/case แบบเข้มงวด:
- ใช้ได้เฉพาะ bag เต็ม (5 PCS)
- ปิดเคสได้เฉพาะ 48 bag (240 PCS)
- อนุญาตช่วยเติมเฉพาะ MFG/EXP เดียวกัน และไม่ละเมิดกฎ Shift/MC
- กรณี 46/47 bag จะพยายามเติมให้ครบเท่านั้น; เติมไม่ครบ = ไม่ allocate


## 9) Excel Workbook Test Structure (Semi Kibble FIFO VBA)

ส่วนนี้เป็น template พร้อมใช้งานสำหรับสร้างไฟล์ทดสอบ `SemiKibble_FIFO_Test.xlsx` ให้สอดคล้องกับ macro `AutoAllocate` ใน `SemiKibbleMVP.bas`

### 9.1 ชีทที่ต้องมี
สร้างชีทตามชื่อด้านล่างให้ตรงตัวอักษรทุกตัว:
- `Stock`
- `Plan_Header`
- `Plan_Detail`
- `AllocationLine`
- `Balance`
- `Log`

---

### 9.2 หัวคอลัมน์ทั้งหมด

#### Sheet: `Stock` (แถวหัวตาราง = Row 1)
| Col | Header | ตัวอย่าง | หมายเหตุ |
|---|---|---|---|
| A | RowID | STK001 | unique ต่อแถว |
| B | MFG | 2026-05-01 | วันที่ผลิต |
| C | EXP | 2026-11-01 | วันหมดอายุ |
| D | Shift | A | ใช้ตรวจกติกา Shift/MC |
| E | MC | MC1 | ใช้ตรวจกติกา Shift/MC |
| F | QtyPCS | 240 | จำนวน PCS ในแถว |
| G | FifoSeq | 1 | ลำดับ FIFO จากน้อยไปมาก |

#### Sheet: `Plan_Header`
- Row 1: หัวแผน `Plan1 ... Plan9`
- Row 2: ป้อนแผนที่ต้องการ allocate (แนะนำป้อนเป็น PCS เช่น 240)
- Row 3: ระบบจะเขียนค่า Normalized PCS ให้หลังรัน

ตัวอย่างหัวคอลัมน์:
| Col A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| Plan1 | Plan2 | Plan3 | Plan4 | Plan5 | Plan6 | Plan7 | Plan8 | Plan9 |

#### Sheet: `Plan_Detail`
ระบบจะเขียนให้อัตโนมัติ:
- A1 = `Customer`
- B1..J1 = `Plan1..Plan9`
- A2 = `AllocatedPCS`
- B2..J2 = PCS ที่ allocate ได้จริง

#### Sheet: `AllocationLine`
หัวคอลัมน์:
| Col | Header |
|---|---|
| A | TxnTime |
| B | Action |
| C | PlanNo |
| D | ReceiverRowID |
| E | DonorRowID |
| F | MovePCS |
| G | CaseCount |

#### Sheet: `Balance`
หัวคอลัมน์ (ระบบเขียนอัตโนมัติ):
| Col | Header |
|---|---|
| A | RowID |
| B | QtyPCS |
| C | AllocatedPCS |
| D | BalancePCS |
| E | RejectPCS |
| F | UsablePCS |

#### Sheet: `Log`
หัวคอลัมน์:
| Col | Header |
|---|---|
| A | TxnTime |
| B | Level |
| C | Message |

---

### 9.3 Sample Data (ใช้ทดสอบทีละ Scenario)

> ทุกชุดข้อมูลให้ล้าง `Stock`, `Plan_Header` Row2, `Plan_Detail`, `AllocationLine`, `Balance`, `Log` ก่อนเริ่มรันเคสใหม่

#### Scenario A: 240 PCS allocate success
`Stock`
| RowID | MFG | EXP | Shift | MC | QtyPCS | FifoSeq |
|---|---|---|---|---|---:|---:|
| STK001 | 2026-05-01 | 2026-11-01 | A | MC1 | 240 | 1 |

`Plan_Header` Row2: Plan1=240, Plan2..Plan9=0

#### Scenario B: 239 PCS reject
`Stock`
| RowID | MFG | EXP | Shift | MC | QtyPCS | FifoSeq |
|---|---|---|---|---|---:|---:|
| STK001 | 2026-05-01 | 2026-11-01 | A | MC1 | 239 | 1 |

`Plan_Header` Row2: Plan1=240

#### Scenario C: 46+1+1 top-up success
`Stock`
| RowID | MFG | EXP | Shift | MC | QtyPCS | FifoSeq |
|---|---|---|---|---|---:|---:|
| STK001 | 2026-05-01 | 2026-11-01 | A | MC1 | 230 | 1 |
| STK002 | 2026-05-01 | 2026-11-01 | A | MC1 | 5 | 2 |
| STK003 | 2026-05-01 | 2026-11-01 | A | MC1 | 5 | 3 |

`Plan_Header` Row2: Plan1=240

#### Scenario D: 46+1 fail
`Stock`
| RowID | MFG | EXP | Shift | MC | QtyPCS | FifoSeq |
|---|---|---|---|---|---:|---:|
| STK001 | 2026-05-01 | 2026-11-01 | A | MC1 | 230 | 1 |
| STK002 | 2026-05-01 | 2026-11-01 | A | MC1 | 5 | 2 |

`Plan_Header` Row2: Plan1=240

#### Scenario E: Shift/MC mismatch fail (ตาม Requirement)
`Stock`
| RowID | MFG | EXP | Shift | MC | QtyPCS | FifoSeq |
|---|---|---|---|---|---:|---:|
| STK001 | 2026-05-01 | 2026-11-01 | A | MC1 | 230 | 1 |
| STK002 | 2026-05-01 | 2026-11-01 | B | MC1 | 5 | 2 |
| STK003 | 2026-05-01 | 2026-11-01 | A | MC2 | 5 | 3 |

`Plan_Header` Row2: Plan1=240

#### Scenario F: Stock 480 + Plan1=240 Plan2=240 Plan3=240 => allocate เฉพาะ 2 แผนแรก
`Stock`
| RowID | MFG | EXP | Shift | MC | QtyPCS | FifoSeq |
|---|---|---|---|---|---:|---:|
| STK001 | 2026-05-01 | 2026-11-01 | A | MC1 | 480 | 1 |

`Plan_Header` Row2: Plan1=240, Plan2=240, Plan3=240

---

### 9.4 Expected Result ต่อ Scenario

#### A) 240 PCS allocate success
- `Plan_Detail!B2 (Plan1)` = 240
- `Balance`: STK001 AllocatedPCS=240, BalancePCS=0
- `AllocationLine` มี 1 รายการ Action=PLAN MovePCS=240

#### B) 239 PCS reject
- `Plan_Detail!B2` = 0
- `Balance`: STK001 RejectPCS=4, UsablePCS=235, AllocatedPCS=0
- `AllocationLine` ว่าง

#### C) 46+1+1 top-up success
- `Plan_Detail!B2` = 240
- `AllocationLine` ควรมี Action=TOPUP_46 อย่างน้อย 2 บรรทัด (จาก STK002 และ STK003)
- `Balance`:
  - STK001 ถูกใช้ 230
  - STK002 ถูกใช้ 5
  - STK003 ถูกใช้ 5

#### D) 46+1 fail
- `Plan_Detail!B2` = 0
- `AllocationLine` ว่าง (ไม่มี top-up ครบ 2 bag)
- `Balance` ไม่มีการหัก allocation

#### E) Shift/MC mismatch fail (Expected by business rule)
- หากใช้ strict Shift/MC rule ต้องได้ `Plan_Detail!B2 = 0`
- `AllocationLine` ว่าง
- หมายเหตุ: โค้ดปัจจุบันใน `HandleTopUp4647` ตรวจแค่ MFG/EXP จึงอาจผ่านเคสนี้ ต้องเพิ่มเงื่อนไข Shift/MC เพื่อให้ผลตรง requirement

#### F) 480 / 240 / 240 / 240
- `Plan_Detail`: Plan1=240, Plan2=240, Plan3=0
- `Balance`: STK001 AllocatedPCS=480, BalancePCS=0
- `AllocationLine` รวม 2 case เท่านั้น

---

### 9.5 VBA Setup Instruction
1. เปิด Excel แล้วบันทึกเป็นไฟล์ชนิด `Excel Macro-Enabled Workbook (*.xlsm)`
2. กด `ALT + F11` เปิด VBA Editor
3. `Insert > Module`
4. วางโค้ดทั้งหมดจากไฟล์ `SemiKibbleMVP.bas`
5. ใน VBA Editor ไปที่ `Tools > References` (ปกติไม่ต้องเพิ่ม reference พิเศษ)
6. กลับ Excel แล้วสร้าง 6 ชีทชื่อให้ตรง: `Stock`, `Plan_Header`, `Plan_Detail`, `AllocationLine`, `Balance`, `Log`
7. เติม header ตามหัวข้อ 9.2

---

### 9.6 วิธีรัน `AutoAllocate`
1. ใส่ข้อมูล `Stock`
2. ใส่ค่าแผนใน `Plan_Header` แถว 2
3. กด `ALT + F8`
4. เลือก macro `AutoAllocate`
5. กด Run
6. ตรวจผลใน `Plan_Detail`, `AllocationLine`, `Balance`, `Log`

---

### 9.7 Test Scenario Checklist

ใช้ checklist นี้ทุกครั้ง:
- [ ] ชื่อชีทครบและสะกดตรง
- [ ] Header ตรงตาม spec
- [ ] QtyPCS เป็นตัวเลขจำนวนเต็ม
- [ ] FifoSeq เรียงถูกต้อง
- [ ] Plan_Header แถว 2 กรอกเฉพาะแผนที่ทดสอบ
- [ ] รัน `AutoAllocate` แล้วไม่มี MsgBox error
- [ ] Plan_Detail ได้ค่าตรง expected
- [ ] AllocationLine มี/ไม่มีรายการตาม expected
- [ ] Balance สะท้อน AllocatedPCS / RejectPCS ถูกต้อง
- [ ] Log มีข้อความ `AutoAllocate completed.` เมื่อจบสำเร็จ
