# SLA Dashboard — สัญญายังไม่ส่งมอบในระบบ

Dashboard ติดตามสถานะดำเนินงานสัญญาที่ยังไม่ส่งมอบ ดึงข้อมูลจากไฟล์ Excel (Google Drive) แล้ว build เป็นไฟล์ `index.html` แบบ static ไม่มี backend, ไม่มี database — ข้อมูลทั้งหมด hardcode ไว้ในไฟล์ HTML ตอน build แต่ละครั้ง

**Deploy:** อัปโหลดไฟล์ HTML ทับ `index.html` ใน GitHub repository นี้ แล้ว GitHub Pages จะ deploy ให้อัตโนมัติ

**⚠️ ห้ามอัปโหลดไฟล์ Excel ต้นฉบับขึ้น repo นี้เด็ดขาด** เพราะมีข้อมูลลูกค้า (ชื่อ, รหัสลูกค้า, ยอดหนี้) — ใช้ไฟล์ Excel แค่เป็น input สำหรับ Claude นำไป build ข้อมูลลง `index.html` เท่านั้น แล้วอัปโหลดเฉพาะ `index.html` ที่ build เสร็จแล้วขึ้น repo (เคยมี repo เก่าหลุดไฟล์ Excel ขึ้นไปแล้วครั้งหนึ่ง — ถูกลบทิ้งไปแล้ว ระวังอย่าให้เกิดซ้ำ)

---

## โครงสร้าง Dashboard (4 Tab)

| Tab | id | ที่มาข้อมูล | จำนวนล่าสุด |
|---|---|---|---|
| **📊 ภาพรวม** | `ov` | `AF` — sheet `AF` ทั้งหมด | 1,023 |
| **🔶 ไม่ดำเนินการภาพรวม** | `ovna` | `OVNA = AF.filter(status==='ไม่ดำเนินการ')` (คำนวณใน JS ไม่มีข้อมูลแยก) | 163 |
| **⚠️ ไม่ดำเนินการ(สคล/สคจ ชี้แจง)** | `na` | `NA` — sheet `สำหรับกรอง-ตรวจสอบ` join กับ `AF` | 163 |
| **📅 ไม่ดำเนินการ(ข้อมูลหน่วยงานอื่น)** | `mn` | `MN` — sheet `รายเดือน-ชน` join กับ `AF` | 151 |

**Tab bar pill IDs:** `tb-ov`, `tb-ovna` (คำนวณอัตโนมัติจาก JS ไม่ต้อง inject), `tb-na`, `tb-mn`

**ตำแหน่ง tab:** ภาพรวม → ไม่ดำเนินการภาพรวม → ไม่ดำเนินการ(สคล/สคจ) → ไม่ดำเนินการ(หน่วยงานอื่น)

### เกี่ยวกับ tab "ไม่ดำเนินการภาพรวม" (OVNA) — สำคัญ
- **ไม่มีข้อมูลแยกเก็บ** — คำนวณจาก `const OVNA=AF.filter(x=>x.status==='ไม่ดำเนินการ');` ใน JavaScript โดยตรง
- field `update` ของ record เหล่านี้ **ถูก join กับ สำหรับกรอง-ตรวจสอบ ไว้แล้ว**ตอน build ข้อมูล AF (ดูหัวข้อ "ปรับสรุปการUPDATE" ด้านล่าง) จึงไม่ต้อง build/inject ข้อมูลซ้ำ
- **ผลคือ tab นี้อัปเดตอัตโนมัติทุกครั้งที่ AF อัปเดต** ไม่ต้องแก้ script การ build เพิ่มเติมเลย
- UI เหมือน tab "ไม่ดำเนินการ(สคล/สคจ)" ทุกประการ (KPI, chart, filter ชุดเดียวกัน, ตารางสรุป, ตารางรายละเอียด)

---

## Sheet ที่ใช้ใน Excel ต้นทาง

ไฟล์ Excel: `สัญญายังไม่ส่งมอบในระบบ-SLA.xlsx`

| Sheet | ใช้ทำอะไร |
|---|---|
| `AF` | ข้อมูลหลัก (master) — ทุกสัญญาทั้งหมด รวม status, วันที่แต่ละขั้นตอนซ่อม, หนี้ค้าง, เช็คสถานะซ่อมหลังสุด |
| `สำหรับกรอง-ตรวจสอบ` | รายการ "ไม่ดำเนินการ" ที่เจ้าหน้าที่ชี้แจงเหตุผลไว้ (`ปรับสรุปการUPDATE`) |
| `รายเดือน-ชน` | รายการไม่ดำเนินการจากหน่วยงานอื่น (ข้อมูลย้อนหลัง) พร้อมชี้แจงเหตุผลเช่นกัน |

**หมายเหตุ:** จำนวนแถวใน `สำหรับกรอง-ตรวจสอบ` อาจไม่เท่ากับจำนวน "ไม่ดำเนินการ" ใน AF เสมอไป (บางครั้งเจ้าหน้าที่กรอกไม่ครบ) — ล่าสุดพบว่าเท่ากันพอดี (163=163) แต่อย่าตั้งสมมติฐานว่าเท่ากันเสมอ

---

## Data Fields ทั้งหมด (ต่อ record ใน AF/NA)

| Field | มาจาก Column ใน Excel | หมายเหตุ |
|---|---|---|
| `dept` | `DEPT_SHORT_NAME` | กอง |
| `bp` | `BP_SHORT_NAME` | สาขา |
| `proj` | `PROJ_NAME` | โครงการ |
| `house` | `HOUSE_NO_ACT` | บ้านเลขที่ (**ไม่ใช่** `HOUSE_NO_STD`) |
| `cust` | `CUST_CODE` | รหัสลูกค้า — ต้อง normalize ก่อน join |
| `status` / `next_custcare` | `NEXT_CUSTCARE` | ขั้นตอน**ถัดไป**ที่ต้องทำ (future-looking) — คนละความหมายกับ `update` (ดูด้านล่าง) |
| `update` | ผสมสองแหล่ง (ดูหัวข้อถัดไป) | **สำคัญมาก** อ่านหัวข้อ "ความแตกต่าง NEXT_CUSTCARE vs CUSTCARE_STATUS_NAME" |
| `guarantee` | `ST_GUARANTEE` | การค้ำประกัน |
| `confirm_flag` | `CONFIRM_FLAG` | ซ่อม / ไม่ดำเนินการ |
| `start_con` | `เริ่มสัญญา` | ค.ศ. → ต้อง +543 |
| `due_last` | `Dueล่าสุด` | ค.ศ. → ต้อง +543 |
| `due_amt` | `เงินDue` | |
| `overdue` | `งวดค้าง` | |
| `debt` | `หนี้ค้าง` | ใช้เป็น flag สำหรับ filter "มีหนี้ค้าง/ไม่มีหนี้ค้าง" |
| `days_lag` | คำนวณจาก `MAX_STATUS_DATE` | = (วันที่ update ไฟล์) − (MAX_STATUS_DATE) |
| `ck_custcare_st` | `CK_CUSTCARE_ST` | **ค่าดิบเต็ม** — ใช้แสดงในตารางเท่านั้น |
| `ck_custcare_st_grp` | parse จาก `CK_CUSTCARE_ST` | ตัด date prefix ออกแล้ว — ใช้ทำ filter/dropdown เท่านั้น |
| `d01`–`d11` | `DATE_01`–`DATE_11` | วันที่แต่ละขั้นตอนซ่อม |
| `aging09` | `AGING_W_09` | วันรอจัดหาผู้รับจ้าง |
| `ck_aging09` | `CK_AGING_09` | ช่วงวันรอจัดหาผู้รับจ้าง |

---

## ⚠️ ความแตกต่าง NEXT_CUSTCARE vs CUSTCARE_STATUS_NAME (สำคัญมาก — เพิ่งแก้ไขรอบล่าสุด)

จากการตรวจสอบข้อมูลจริงพบว่า **2 คอลัมน์นี้ต่างกันถึง 669/1023 รายการ (65%)** ไม่ใช่แค่ชื่อต่างกัน แต่ความหมายคนละมุมมอง:

- **`NEXT_CUSTCARE`** = ขั้นตอน**ถัดไป**ที่ต้องทำ (future-looking) → เก็บใน field `next_custcare`/`status`
- **`CUSTCARE_STATUS_NAME`** = สถานะ**ปัจจุบันจริงๆ** ที่อยู่ตอนนี้ (current actual stage)

ตัวอย่าง: `NEXT_CUSTCARE="ประมาณราคา"` (ขั้นต่อไปคือประมาณราคา) แต่ `CUSTCARE_STATUS_NAME="ยืนยันการสำรวจซ่อม"` (ตอนนี้อยู่ที่ขั้นยืนยันสำรวจจริงๆ)

**field `update`** ใน AF ถูกคำนวณผสมสองแหล่งตอน build:
```python
if status=='ไม่ดำเนินการ':
    update = chk_lookup[cust]['update']  # จาก สำหรับกรอง-ตรวจสอบ (คำชี้แจง)
else:
    update = CUSTCARE_STATUS_NAME        # สถานะปัจจุบันจริง
```
ดังนั้น field `update` เป็นได้ 2 แบบขึ้นกับ status — **ต้องระวังเวลานำไปใช้ต่อ** อย่าสับสนว่าเป็นค่าเดียวกันหมด

**Filter "สถานะซ่อม" เดิม (`ov-next`)** ใช้ `NEXT_CUSTCARE` — **ถูกซ่อนไว้แล้ว** (`style="display:none"` ครอบ label+select ไม่ได้ลบโค้ด) เพราะผู้ใช้ต้องการดูสถานะปัจจุบันจริงมากกว่า

**Filter ใหม่ "สถานะซ่อมปัจจุบัน" (`ov-custcare`)** ใช้ `CUSTCARE_STATUS_NAME` (ผ่าน field `update`) — มีอยู่ใน**tab ภาพรวมเท่านั้น** dropdown มี **12 ตัวเลือกคงที่** (hardcode ไว้ใน `initDDs()` ไม่ใช่ dynamic uniq):
```
รับคำร้องแจ้งซ่อม, สำรวจซ่อม, ยืนยันการสำรวจซ่อม, ประมาณราคา, ตรวจสอบและรับรองราคา,
อนุมัติสรุปค่าก่อสร้าง, รอการจัดจ้าง, จัดหาผู้รับจ้าง, ดำเนินการซ่อม, ดำเนินการเสร็จสิ้น,
ขอเบิกเงินค่าจ้าง, ไม่ดำเนินการ
```
**เหตุผลที่ hardcode:** ถ้าใช้ dynamic uniq จาก field `update` ตรงๆ จะปนคำชี้แจง (เช่น "ยังไม่สำรวจ", "ลูกค้ารับตามสภาพ") ของเคส "ไม่ดำเนินการ" เข้ามาด้วย (เจอปัญหานี้มาแล้วรอบหนึ่ง) จึง fix เป็น 11 สถานะซ่อมจริง + "ไม่ดำเนินการ" อีก 1 รวม 12

**Filter logic พิเศษ** เมื่อเลือก "ไม่ดำเนินการ" ต้อง match กับ field `status` แทน `update` (เพราะ `update` ของเคสนี้เป็นคำชี้แจง ไม่ใช่คำว่า "ไม่ดำเนินการ"):
```javascript
(!cs||(cs==='ไม่ดำเนินการ'?x.status==='ไม่ดำเนินการ':x.update===cs))
```

---

## Column Index Mapping (สำหรับอ่านด้วย openpyxl)

### จาก sheet `AF` (header row 3 / index 2, data เริ่ม index 3)
ใช้ `hi = {v:i for i,v in enumerate(header_row) if v}` อ้างชื่อ column ตรงๆ ไม่ hardcode index

Column ที่ต้องมีเสมอ: `DEPT_SHORT_NAME`, `BP_SHORT_NAME`, `PROJ_NAME`, `HOUSE_NO_ACT`, `CUST_CODE`, `CUSTNAME1DISPLAY`, `CHANGE_STATUS`, `NEXT_CUSTCARE`, `CUSTCARE_STATUS_NAME`, `ST_GUARANTEE`, `CONFIRM_FLAG`, `เริ่มสัญญา`, `Dueล่าสุด`, `เงินDue`, `งวดค้าง`, `หนี้ค้าง`, `MAX_STATUS_DATE`, `CK_CUSTCARE_ST`, `DATE_01`–`DATE_11`, `AGING_W_09`, `CK_AGING_09`

### จาก sheet `สำหรับกรอง-ตรวจสอบ` (data เริ่ม row 4 / index 3)
`dept`=col2, `bp`=col3, `proj`=col4, `house`=col5, `cust`=col6 (join key), `cust_name`=col7, `change_status`=col11, `guarantee`=col13, `confirm_flag`=col14, `update`=**col21** (`ปรับสรุปการUPDATE` — **ห้ามใช้** col17 ที่มีค่าว่างปนอยู่)

**Guard เสมอ:** `if len(r) < 22 or not r[2]: continue`

### จาก sheet `รายเดือน-ชน` (header row 61, data เริ่ม row 62)
`dept`=col2, `bp`=col3, `proj`=col4, `house`=col5, `cust`=col6, `update`=**col18**

**Guard เสมอ:** `if len(r) < 19 or not r[2] or not str(r[2]).strip(): continue`

---

## คอลัมน์ "วันค้างสถานะซ่อมหลังสุด" (เพิ่มใหม่ล่าสุด — ต่อท้าย "วันค้างสถานะ" ทุก tab)

คำนวณจาก `ck_custcare_st` (ค่าดิบ) ล้วนๆ ผ่าน JS **ไม่มีการ inject field ใหม่จาก Python** — logic อยู่ในฟังก์ชัน `ckLagValue()`/`ckLagInfo()`/`ckLagTd()`:

| กรณี | เงื่อนไข | ค่าที่แสดง |
|---|---|---|
| ไม่เคยซ่อม | `ck_custcare_st` ว่างเปล่า | `"ไม่เคยซ่อม"` |
| เคยซ่อม (มีวันที่) | มี prefix ตัวเลข 8 หลัก เช่น `20260122 ดำเนินการเสร็จสิ้น` | คำนวณ (`UPD_REF_DATE` − วันที่นั้น) |
| กำลังซ่อมอยู่ (real-time) | มีข้อความอย่างเดียว ไม่มีวันที่ (พบเกือบทุก record ที่ status≠ไม่ดำเนินการ ~860 รายการ) | ใช้ `days_lag` เดิม |

**`UPD_REF_DATE`** เป็นค่าคงที่ hardcode ไว้ใน JS (ต้องแก้ทุกครั้งที่ update ข้อมูลใหม่!):
```javascript
const UPD_REF_DATE=new Date(2026,7,13); // เดือนเริ่มที่ 0 → 7=สิงหาคม
```
⚠️ **อย่าลืมแก้ค่านี้ให้ตรงกับ UPDATE_DATE ทุกครั้งที่ update ข้อมูลใหม่** ไม่งั้นคอลัมน์นี้จะคำนวณผิด (ตอนนี้ต้องแก้ manual ยังไม่ auto-sync กับ timestamp banner)

**Filter คู่กัน "วันค้างสถานะซ่อมหลังสุด"** (`xx-cklag`) มีในทุก tab เช่นกัน ตัวเลือก: ทั้งหมด / ไม่เคยซ่อม (`value="never"`) / 0-30 วัน / 31-60 วัน / 61-90 วัน / เกิน 90 วัน — ใช้ฟังก์ชัน `inCkLagRange(val, range)` ร่วมกับ `ckLagValue()`

---

## ⚠️ บั๊กที่แก้แล้ว: ค่า 0 วันตกหล่นจาก filter ช่วงวัน

**ปัญหาเดิม:** `inLagRange()` เดิม bucket แรกเขียนว่า `lg>=1&&lg<=30` (เริ่มที่ 1) ทำให้ record ที่ `days_lag===0` (วันเดียวกับวันที่ update ข้อมูลพอดี) **ไม่ตกอยู่ใน bucket ไหนเลย** — พบทั้งใน filter "วันค้างสถานะ" เดิม (115 รายการ) และ "วันค้างสถานะซ่อมหลังสุด" ใหม่ (117 รายการ)

**วิธีแก้:** เปลี่ยน bucket แรกเป็น `lg>=0&&lg<=30` (label เปลี่ยนจาก "1-30 วัน" → **"0-30 วัน"**) ทั้ง 2 filter ทุก tab — แก้ที่ `inLagRange()` จุดเดียว มีผลกับทุก tab อัตโนมัติ

**ทดสอบยืนยัน:** ผลรวมทุก bucket ของทั้ง 2 filter = จำนวนรายการทั้งหมดพอดี ไม่ตกหล่นแล้ว

---

## Filter ที่มีในแต่ละ Tab (id จริงในโค้ด ล่าสุด)

| Filter | ov | ovna | na | mn |
|---|---|---|---|---|
| กอง | `ov-dept` | `ovna-dept` | `na-dept` | `mn-dept` |
| สาขา | `ov-bp` | `ovna-bp` | `na-bp` | `mn-bp` |
| โครงการ | `ov-proj` | `ovna-proj` | `na-proj` | `mn-proj` |
| ปรับสรุปการ UPDATE | — | `ovna-upd` | `na-upd` | `mn-upd` |
| **สถานะซ่อม (ซ่อนไว้)** | `ov-next` *(hidden)* | — | — | — |
| **สถานะซ่อมปัจจุบัน (ใหม่)** | `ov-custcare` | — | — | — |
| ช่วงวันรอจัดหาฯ | `ov-ck09` | `ovna-ck09` | `na-ck09` | `mn-ck09` |
| หนี้ค้าง | `ov-debt` | `ovna-debt` | `na-debt` | `mn-debt` |
| วันค้างสถานะ | `ov-lag` | `ovna-lag` | `na-lag` | `mn-lag` |
| **วันค้างสถานะซ่อมหลังสุด (ใหม่)** | `ov-cklag` | `ovna-cklag` | `na-cklag` | `mn-cklag` |
| เช็คสถานะซ่อมหลังสุด | `ov-ck` | `ovna-ck` | `na-ck` | `mn-ck` |
| การค้ำประกัน | `ov-guarantee` | — | — | — |
| การซ่อม | `ov-confirm` | — | — | — |

**Filter "เช็คสถานะซ่อมหลังสุด"** เทียบกับ `x.ck_custcare_st_grp` (ค่าตัดวันที่แล้ว) **ไม่ใช่** `x.ck_custcare_st` (ค่าดิบที่ใช้แสดงในตารางเท่านั้น)

---

## Business Logic สำคัญ

1. **`ปรับสรุปการUPDATE` ไม่ใช่สถานะการซ่อม** เป็นคำชี้แจงเหตุผลว่าทำไมเคสถึงยังค้างอยู่ (เฉพาะเคส "ไม่ดำเนินการ")
2. **ยืนยันว่าซ่อมเริ่มจริงหรือยัง** ดูที่ `DATE_01`–`DATE_11` เท่านั้น
3. **`เช็คสถานะซ่อมหลังสุด` (`CK_CUSTCARE_ST`)** ถ้าว่าง = "ไม่เคยซ่อม" ต้อง default เสมอ
4. **`รอการUpdate`** ปรากฏเมื่อ cust code ใน AF (status="ไม่ดำเนินการ") หาไม่เจอใน `สำหรับกรอง-ตรวจสอบ`
5. **`NEXT_CUSTCARE` ≠ `CUSTCARE_STATUS_NAME`** (ดูหัวข้อด้านบน) — คนละมุมมอง อย่าใช้แทนกัน
6. **ค่า `'X'`** ใน `ปรับสรุปการUPDATE` → normalize เป็น `'รอการUpdate'` เสมอ

---

## UC / MC (สี badge)

`UC`/`UB` ใช้กับ NA + **OVNA** (OVNA ใช้ AF ที่มี update join เดียวกัน จึงใช้ UC ร่วมกันได้) และ `MC`/`MB` ใช้กับ MN

**ทุกครั้งที่ update ข้อมูล ต้องตรวจสอบ 3 จุด:**
```
set(update values ใน NA) − set(UC keys) == empty
set(update values ใน OVNA คือ AF.filter(status='ไม่ดำเนินการ')) − set(UC keys) == empty
set(update values ใน MN) − set(MC keys) == empty
```
ถ้าเจอ key ใหม่ ต้องเพิ่มเข้า UC/UB หรือ MC/MB ก่อน inject ข้อมูล

---

## ขั้นตอนการ Update ข้อมูล (สำหรับ Claude ในแชทใหม่)

1. รับไฟล์ Excel ใหม่ — ทำงานในเครื่อง sandbox เท่านั้น ห้ามอัปโหลดขึ้น GitHub
2. ตรวจสอบ column: `เริ่มสัญญา`, `สิ้นสุดสัญญา`, `Dueล่าสุด`, `เงินDue`, `งวดค้าง`, `หนี้ค้าง`, `MAX_STATUS_DATE`, `CK_CUSTCARE_ST`
3. Build `AF_new`, `NA_data`, `MN_new` — แยก `ck_custcare_st` (ดิบ) กับ `ck_custcare_st_grp` (ตัดวันที่แล้ว) เสมอ
4. หา `UPDATE_DATE` จาก `MAX_STATUS_DATE` สูงสุดใน AF
5. **⚠️ อัปเดต `UPD_REF_DATE` ใน JS ด้วย (ใช้คำนวณคอลัมน์ "วันค้างสถานะซ่อมหลังสุด")** — จุดที่มักลืมแก้
6. Normalize `'X'` → `'รอการUpdate'`
7. ตรวจสอบ UC/MC ครอบคลุมทุกค่า (รวม OVNA ด้วย — ดูหัวข้อด้านบน)
8. Inject `const AF=`, `const NA=`, `const MN=` (regex replace) — **ไม่ต้อง inject OVNA** เพราะคำนวณจาก AF อัตโนมัติใน JS
9. อัปเดต tab pill (`tb-ov`, `tb-na`, `tb-mn`) และ timestamp — `tb-ovna` ไม่ต้องแก้ (auto)
10. รัน `node --check` ก่อนส่งไฟล์เสมอ
11. ตรวจสอบครั้งสุดท้าย: จำนวนแต่ละ tab ตรงกับ pill, ไม่มี `รอการUpdate` ค้าง, debt/days_lag/ck_custcare_st มีข้อมูล, UC/MC ครบ, `UPD_REF_DATE` ตรงกับวันที่ล่าสุด

---

## เวลาเพิ่ม Filter ใหม่ — Checklist

1. [ ] เพิ่ม `<select>` ใน frow ของ tab ที่เกี่ยวข้อง
2. [ ] เพิ่มตัวแปรอ่านค่า (`const xxx=getV('...')`) ใน filter function ของ tab นั้น
3. [ ] เพิ่มเงื่อนไขกรองต่อท้าย **นอกวงเล็บของเงื่อนไขก่อนหน้า**
4. [ ] เพิ่ม `updSel('xxx-yyy', ...)` ใน `initDDs()` (หรือ hardcode fixed list ถ้าต้องการควบคุม option ตายตัว เช่นกรณี `ov-custcare`)
5. [ ] เพิ่ม `setV('xxx-yyy','')` ใน Clear function ของ tab นั้น
6. [ ] รัน `node --check` แล้ว **ทดสอบ runtime จริง** (node -e หรือไฟล์ .js แยก) ด้วยข้อมูลจริง เช็คว่าผลรวมทุก bucket = จำนวนรายการทั้งหมด (ไม่ตกหล่น) — บั๊ก "ค่า 0 ตกหล่น" เจอเพราะไม่ได้เช็คจุดนี้มาก่อน
7. [ ] เช็คว่า chart/card/ตารางในทุก tab ใช้ตัวแปร `f` (ผลลัพธ์จาก filtered function) ไม่ใช่ array ดิบ — สแกนด้วย `grep` หาการใช้ `AF.filter`/`AF.map`/`uniq(AF` ภายใน render function ควรไม่พบเลย

---

## เวลาซ่อน Filter (ไม่ลบ) — Pattern ที่ใช้แล้ว

ถ้าต้องการเปลี่ยน filter โดยไม่ทำลายโครงสร้างเดิม (เผื่อย้อนกลับ) ใช้วิธี:
```html
<span style="display:none;"><label>ชื่อเดิม:</label><select id="เดิม">...</select></span>
<label>ชื่อใหม่:</label><select id="ใหม่">...</select>
```
เพิ่ม filter ใหม่ต่อท้าย ไม่ลบของเดิม — JS logic เดิม (`getV('เดิม')`) ยังทำงานได้ปกติ (คืนค่า `''` เสมอเพราะซ่อนไว้ไม่มีการเปลี่ยนค่า) ไม่กระทบผลลัพธ์ filter รวม

---

## ข้อควรระวังอื่นๆ

- **Browser cache**: หลัง deploy ให้ผู้ใช้ hard refresh (Ctrl+Shift+R / Cmd+Shift+R)
- **cust code format**: normalize ด้วย `nc()` ก่อน join เสมอ (float string → int string)
- **ไม่มีระบบ backup อัตโนมัติ** — พึ่ง Git history บน GitHub เอง
- **GitHub sudo mode**: อาจติดปัญหายืนยันตัวตนซ้ำผ่านอีเมลโดยเฉพาะบน mobile — แนะนำใช้คอมพิวเตอร์หรือ GitHub Desktop App แทน
- **PDF Export**: ใช้เทคนิค Blob URL (`new Blob([html]) → URL.createObjectURL → window.open → window.print()`) ไม่ใช้ library ภายนอก — ใช้ `document.getElementById(id).outerHTML` clone ตารางทั้งก้อน ดังนั้นคอลัมน์ใหม่ที่เพิ่มเข้าตารางจะติด export ไปอัตโนมัติเสมอ ไม่ต้องแก้โค้ด export เพิ่ม
- **Excel Export**: ใช้ `XLSX.utils.table_to_sheet()` clone ตาราง DOM เช่นกัน หลักการเดียวกับ PDF
