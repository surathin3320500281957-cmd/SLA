# SLA Dashboard — สัญญายังไม่ส่งมอบในระบบ

Dashboard ติดตามสถานะดำเนินงานสัญญาที่ยังไม่ส่งมอบ ดึงข้อมูลจากไฟล์ Excel (Google Drive) แล้ว build เป็นไฟล์ `index.html` แบบ static ไม่มี backend, ไม่มี database — ข้อมูลทั้งหมด hardcode ไว้ในไฟล์ HTML ตอน build แต่ละครั้ง

**Deploy:** อัปโหลดไฟล์ HTML ทับ `index.html` ใน GitHub repository นี้ แล้ว GitHub Pages จะ deploy ให้อัตโนมัติ

**⚠️ ห้ามอัปโหลดไฟล์ Excel ต้นฉบับขึ้น repo นี้เด็ดขาด** เพราะมีข้อมูลลูกค้า (ชื่อ, รหัสลูกค้า, ยอดหนี้) — ใช้ไฟล์ Excel แค่เป็น input สำหรับ Claude นำไป build ข้อมูลลง `index.html` เท่านั้น แล้วอัปโหลดเฉพาะ `index.html` ที่ build เสร็จแล้วขึ้น repo (เคยมี repo เก่าหลุดไฟล์ Excel ขึ้นไปแล้วครั้งหนึ่ง — ถูกลบทิ้งไปแล้ว ระวังอย่าให้เกิดซ้ำ)

---

## โครงสร้าง Dashboard (3 Tab)

| Tab | ชื่อ id ตัวแปร JS | ข้อมูลจาก | จำนวนล่าสุด |
|---|---|---|---|
| **ภาพรวม** | `AF` | sheet `AF` ทั้งหมด | ~995 รายการ |
| **ไม่ดำเนินการ(สคล/สคจ ชี้แจง)** | `NA` | sheet `สำหรับกรอง-ตรวจสอบ` join กับ `AF` | ~134 รายการ |
| **ไม่ดำเนินการ(ข้อมูลหน่วยงานอื่น)** | `MN` | sheet `รายเดือน-ชน` join กับ `AF` | ~142 รายการ |

แต่ละ tab มี: KPI cards, chart (donut/bar), ตารางสรุป (pivot by กอง×สาขา×สถานะ), ตารางรายละเอียด พร้อมปุ่ม Export Excel/PDF

---

## Sheet ที่ใช้ใน Excel ต้นทาง

ไฟล์ Excel: `สัญญายังไม่ส่งมอบในระบบ-SLA.xlsx`

| Sheet | ใช้ทำอะไร |
|---|---|
| `AF` | ข้อมูลหลัก (master) — ทุกสัญญาทั้งหมด รวม status, วันที่แต่ละขั้นตอนซ่อม, หนี้ค้าง, เช็คสถานะซ่อมหลังสุด |
| `สำหรับกรอง-ตรวจสอบ` | รายการ "ไม่ดำเนินการ" ที่เจ้าหน้าที่ชี้แจงเหตุผลไว้ (`ปรับสรุปการUPDATE`) |
| `รายเดือน-ชน` | รายการไม่ดำเนินการจากหน่วยงานอื่น (ข้อมูลย้อนหลัง) พร้อมชี้แจงเหตุผลเช่นกัน |

---

## Data Fields ทั้งหมด (ต่อ record)

รวม 32 fields ต่อ record (MN มีน้อยกว่าเล็กน้อย เพราะไม่มี `cust_name`, `change_status`, `confirm_flag`, `next_custcare`, `status`)

| Field | มาจาก Column ใน Excel | หมายเหตุ |
|---|---|---|
| `dept` | `DEPT_SHORT_NAME` | กอง |
| `bp` | `BP_SHORT_NAME` | สาขา |
| `proj` | `PROJ_NAME` | โครงการ |
| `house` | `HOUSE_NO_ACT` | บ้านเลขที่ (**ไม่ใช่** `HOUSE_NO_STD`) |
| `cust` | `CUST_CODE` | รหัสลูกค้า — ต้อง normalize ก่อน join (ดูหัวข้อ cust code format ด้านล่าง) |
| `status` / `next_custcare` | `NEXT_CUSTCARE` | สถานะซ่อม เช่น "ไม่ดำเนินการ", "รออนุมัติงบ" ฯลฯ (เฉพาะ AF/NA) |
| `update` | `CUSTCARE_STATUS_NAME` (non-NA) หรือ join จาก sheet อื่น (NA/MN) | ดูหัวข้อ "ปรับสรุปการUPDATE" ด้านล่าง |
| `guarantee` | `ST_GUARANTEE` | การค้ำประกัน |
| `confirm_flag` | `CONFIRM_FLAG` | ซ่อม / ไม่ดำเนินการ (เฉพาะ AF/NA) |
| `start_con` | `เริ่มสัญญา` | ค.ศ. → ต้อง +543 |
| `due_last` | `Dueล่าสุด` | ค.ศ. → ต้อง +543 |
| `due_amt` | `เงินDue` | |
| `overdue` | `งวดค้าง` | |
| `debt` | `หนี้ค้าง` | ใช้เป็น flag สำหรับ filter "มีหนี้ค้าง/ไม่มีหนี้ค้าง" |
| `days_lag` | คำนวณจาก `MAX_STATUS_DATE` | = (วันที่ update ไฟล์) − (MAX_STATUS_DATE) หน่วยเป็นวัน ใช้ทำ filter ช่วงวัน |
| **`ck_custcare_st`** | `CK_CUSTCARE_ST` | **ค่าดิบเต็ม** (อาจมี date prefix ปน เช่น `20260122 ดำเนินการเสร็จสิ้น`) — **ใช้แสดงในตารางเท่านั้น ห้ามใช้ทำ filter** |
| **`ck_custcare_st_grp`** | parse จาก `CK_CUSTCARE_ST` | **ค่าที่ตัด date prefix ออกแล้ว** (เหลือแค่คำสถานะ, ว่าง→`ไม่เคยซ่อม`) — **ใช้ทำ filter/dropdown เท่านั้น ห้ามใช้แสดงในตาราง** |
| `d01`–`d11` | `DATE_01`–`DATE_11` | วันที่แต่ละขั้นตอนซ่อม (datetime string, +543) |
| `aging09` | `AGING_W_09` | วันรอจัดหาผู้รับจ้าง |
| `ck_aging09` | `CK_AGING_09` | ช่วงวันรอจัดหาผู้รับจ้าง (เกิน 13 วัน / ภายในกำหนด) |

**⚠️ กฎสำคัญเรื่อง `ck_custcare_st` vs `ck_custcare_st_grp`:**
```
ตาราง (td)       → ใช้ ck_custcare_st       (ค่าดิบ ไม่ตัดอะไร)
filter dropdown  → ใช้ ck_custcare_st_grp   (ตัด date prefix แล้ว)
```
ถ้าสลับกันผิด ตารางจะเสียรายละเอียด หรือ filter จะแตกเป็นหลายสิบตัวเลือกซ้ำซ้อนจากวันที่ต่างกัน — เคยเกิดปัญหานี้มาแล้วครั้งหนึ่ง ระวังอย่าทำผิดซ้ำ

**⚠️ ข้อควรระวังเรื่องปี:**
- `เริ่มสัญญา`, `Dueล่าสุด`, `DATE_01-11`, `MAX_STATUS_DATE` → เก็บเป็น **ค.ศ.** ต้อง `+543`
- `สิ้นสุดสัญญา` → column นี้ถูก**ลบออกจาก Dashboard แล้ว** (ผู้ใช้ขอให้เอาออกเพราะข้อมูลสับสน) — field ไม่มีอยู่ใน data structure อีกต่อไป

---

## Column Index Mapping (สำหรับอ่านด้วย openpyxl)

### จาก sheet `AF` (header row 3 / index 2 แบบ 0-based, data เริ่ม index 3)
ใช้ `hi = {v:i for i,v in enumerate(header_row) if v}` แล้วอ้างชื่อ column ตรงๆ (เช่น `hi['CUST_CODE']`) ไม่ต้อง hardcode index เพราะตำแหน่ง column ในไฟล์อาจขยับได้ในแต่ละรอบ

Column ที่ต้องมีเสมอ: `DEPT_SHORT_NAME`, `BP_SHORT_NAME`, `PROJ_NAME`, `HOUSE_NO_ACT`, `CUST_CODE`, `CUSTNAME1DISPLAY`, `CHANGE_STATUS`, `NEXT_CUSTCARE`, `CUSTCARE_STATUS_NAME`, `ST_GUARANTEE`, `CONFIRM_FLAG`, `เริ่มสัญญา`, `Dueล่าสุด`, `เงินDue`, `งวดค้าง`, `หนี้ค้าง`, `MAX_STATUS_DATE`, `CK_CUSTCARE_ST`, `DATE_01`–`DATE_11`, `AGING_W_09`, `CK_AGING_09`

### จาก sheet `สำหรับกรอง-ตรวจสอบ` (data เริ่ม row 4 / index 3)

| Field | Column index (0-based) |
|---|---|
| `dept` | col 2 |
| `bp` | col 3 |
| `proj` | col 4 |
| `house` | col 5 |
| `cust` | col 6 ← key join กับ AF |
| `cust_name` | col 7 |
| `change_status` | col 11 |
| `guarantee` | col 13 |
| `confirm_flag` | col 14 |
| `update` | **col 21** ← `ปรับสรุปการUPDATE` |

**สำคัญ:** ใช้ col 21 เท่านั้น **ห้ามใช้** col 17 (`สรุปการUPDATE`) เพราะมีค่าว่างปนอยู่และไม่ตรงความจริง

**Guard เสมอ:** `if len(r) < 22 or not r[2]: continue` ก่อนอ่าน col 21 — บาง row สั้นกว่านี้ทำให้ `IndexError`

### จาก sheet `รายเดือน-ชน` (header row 61 / index 60, data เริ่ม row 62 / index 61)

| Field | Column index |
|---|---|
| `dept` | col 2 |
| `bp` | col 3 |
| `proj` | col 4 |
| `house` | col 5 |
| `cust` | col 6 ← key join กับ AF |
| `update` | **col 18** ← `ปรับสรุปการUPDATE` |

**Guard เสมอ:** `if len(r) < 19 or not r[2] or not str(r[2]).strip(): continue`

---

## Filter ที่มีในแต่ละ Tab (id ตัวแปรจริงในโค้ด)

| Filter | ov (ภาพรวม) | na (ไม่ดำเนินการ) | mn (หน่วยงานอื่น) |
|---|---|---|---|
| กอง | `ov-dept` | `na-dept` | `mn-dept` |
| สาขา | `ov-bp` | `na-bp` | `mn-bp` |
| โครงการ | `ov-proj` | `na-proj` | `mn-proj` |
| ปรับสรุปการ UPDATE | — | `na-upd` | `mn-upd` |
| ช่วงวันรอจัดหาผู้รับจ้าง | `ov-ck09` | `na-ck09` | `mn-ck09` |
| หนี้ค้าง (มี/ไม่มี) | `ov-debt` | `na-debt` | `mn-debt` |
| วันค้างสถานะ (1-30/31-60/61-90/90+) | `ov-lag` | `na-lag` | `mn-lag` |
| เช็คสถานะซ่อมหลังสุด | `ov-ck` | `na-ck` | `mn-ck` |
| การค้ำประกัน | `ov-guarantee` | — | — |
| การซ่อม | `ov-confirm` | — | — |
| สถานะซ่อม | `ov-next` | — | — |

**Filter "วันค้างสถานะ"** ใช้ helper function กลาง `inLagRange(days_lag, range)` (ประกาศแยกไว้ก่อน 3 filter functions) — เวลาแก้ range หรือเพิ่ม range ใหม่ แก้ที่ฟังก์ชันนี้จุดเดียวพอ ไม่ต้องแก้ 3 ที่

**Filter "เช็คสถานะซ่อมหลังสุด"** เทียบกับ `x.ck_custcare_st_grp` (ค่าตัดแล้ว) ไม่ใช่ `x.ck_custcare_st` (ค่าดิบ) — ดูกฎด้านบน

---

## Business Logic สำคัญ (อย่าลืม!)

1. **`ปรับสรุปการUPDATE` ไม่ใช่สถานะการซ่อม** — เป็นแค่ **คำชี้แจงเหตุผล** ว่าทำไมเคสถึงยังค้างอยู่ (เช่น "ยังไม่สำรวจ", "สำรวจแล้วรอลงระบบ SLA", "อยู่ในค้ำประกันฝ่ายก่อสร้าง") แม้จะมีคำว่า "ซ่อมแล้ว" ปนอยู่ ไม่ได้แปลว่ากระบวนการซ่อมเริ่มแล้วจริง

2. **การยืนยันว่าซ่อมเริ่มจริงหรือยัง** ให้ดูที่ `DATE_01`–`DATE_11` ใน sheet `AF` เท่านั้น — ถ้ามีวันที่ปรากฏ = เริ่มกระบวนการซ่อมแล้วจริง ถ้าว่างทั้งหมด = ยังไม่เริ่มอะไรเลย

3. **`เช็คสถานะซ่อมหลังสุด` (`CK_CUSTCARE_ST`)** เป็นสถานะซ่อมล่าสุดที่ track แยกจาก `NEXT_CUSTCARE`/`ปรับสรุปการUPDATE` — ถ้าว่างหมายถึง **"ไม่เคยซ่อม"** เลย ต้อง default ค่านี้เสมอเวลาไม่มีข้อมูล (ทั้งใน AF join และ empty_extra())

4. **`รอการUpdate` (ค่า default ของระบบ)** จะปรากฏเมื่อ cust code ใน `AF` (status = "ไม่ดำเนินการ") **หาไม่เจอ** ใน sheet `สำหรับกรอง-ตรวจสอบ` — แปลว่าเจ้าหน้าที่ยังไม่ได้กรอกคำชี้แจงลงไฟล์ ไม่ใช่ bug ของ Dashboard

5. **จำนวนรายการในแต่ละ tab จะเปลี่ยนทุกครั้งที่ไฟล์ Excel อัปเดต** เพราะสถานะขยับตลอด (เคสเก่าปิด/ส่งมอบไปแล้วหลุดออก, เคสใหม่เข้ามาเพิ่ม)

6. **บางครั้งพบค่าประหลาด เช่น `'X'`** ใน `ปรับสรุปการUPDATE` (data entry ผิดพลาดจากผู้กรอก) — ให้ normalize เป็น `'รอการUpdate'` ก่อน inject เสมอ

---

## UC / MC (สี badge สำหรับแต่ละสถานะ `update`)

`UC`/`UB` (ใช้กับ NA) และ `MC`/`MB` (ใช้กับ MN) เป็น object เก็บสี/class badge ของแต่ละค่าใน `ปรับสรุปการUPDATE`

**⚠️ ทุกครั้งที่ update ข้อมูล ต้องตรวจสอบว่า:**
```
set(ค่า update ทั้งหมดใน NA/MN) − set(keys ใน UC/MC) == empty set
```
ถ้าเจอ key ใหม่ที่ไม่มีใน UC/MC ต้อง**เพิ่มเข้า UC/UB (สำหรับ NA) หรือ MC/MB (สำหรับ MN)** ก่อน inject ข้อมูล ไม่งั้น column badge จะไม่มีสี/ไม่ถูกนับในตารางสรุป

**หมายเหตุ:** `ck_custcare_st_grp` (filter ใหม่) ไม่มี badge color map แยก — ใช้ text ธรรมดาใน dropdown/table ไม่ต้องเพิ่ม UC/MC สำหรับ field นี้

---

## ขั้นตอนการ Update ข้อมูล (สำหรับ Claude ในแชทใหม่)

1. รับไฟล์ Excel ใหม่จากผู้ใช้ — **ทำงานกับไฟล์ในเครื่อง (sandbox) เท่านั้น อย่าอัปโหลดไฟล์ Excel ขึ้น GitHub เด็ดขาด**
2. **ตรวจสอบก่อนเสมอ** ว่า 6 columns หนี้ค้างมีอยู่ใน sheet `AF` หรือไม่ (`เริ่มสัญญา`, `สิ้นสุดสัญญา`, `Dueล่าสุด`, `เงินDue`, `งวดค้าง`, `หนี้ค้าง`) และ `CK_CUSTCARE_ST` — ถ้าไม่มี ให้แจ้งผู้ใช้ก่อน อย่า build ทับแบบเงียบๆ
3. Build 3 datasets: `AF_new` (ทุก row จาก AF), `NA_data` (จาก สำหรับกรอง-ตรวจสอบ join AF), `MN_new` (จาก รายเดือน-ชน join AF) — **อย่าลืม parse `ck_custcare_st_grp` แยกจาก `ck_custcare_st` ดิบ**
4. หา `UPDATE_DATE` อัตโนมัติจากค่า `MAX_STATUS_DATE` ที่มากที่สุดใน sheet AF (ไม่ต้องถามผู้ใช้)
5. Normalize คำว่า `'X'` หรือค่าประหลาดอื่นๆ ใน `update` field ให้เป็น `'รอการUpdate'`
6. ตรวจสอบ `UC`/`MC` ว่าครอบคลุมทุกค่าที่เจอใน `update` field — ถ้าขาด เพิ่มเข้าไปก่อน inject
7. Inject 3 arrays เข้า `<script>` ด้วย regex replace (`const AF=`, `const NA=`, `const MN=`), อัปเดต tab pill (`tb-ov`, `tb-na`, `tb-mn`) และ timestamp
8. รัน `node --check` ตรวจ syntax ก่อนส่งไฟล์
9. ตรวจสอบครั้งสุดท้าย: จำนวนแต่ละ tab ตรงกับ pill, ไม่มี `รอการUpdate` ค้าง (ถ้ามีให้แจ้งผู้ใช้), debt/days_lag/ck_custcare_st มีข้อมูล, ตาราง+chart อัปเดตตาม filter ใหม่ทุกตัว
10. ส่งไฟล์ `index.html` ที่ build เสร็จแล้วให้ผู้ใช้ **นำไปอัปโหลดเองบน GitHub** — Claude ไม่มีสิทธิ์เข้าถึง GitHub repo โดยตรง

---

## เวลาเพิ่ม Filter ใหม่ — Checklist ที่ต้องแก้ครบทุกจุด

จากประสบการณ์ที่เคยพลาด (วงเล็บผิดตำแหน่งทำให้ filter ใช้ไม่ได้) ให้ไล่ checklist นี้ทุกครั้ง:

1. [ ] เพิ่ม `<select>` ใน frow ของ tab ที่เกี่ยวข้อง (ov/na/mn อาจต้องเพิ่มมากกว่า 1 tab)
2. [ ] เพิ่มตัวแปรอ่านค่า (`const xxx=getV('...')`) ใน `ovFiltered()`/`naFiltered()`/`mnFiltered()`
3. [ ] เพิ่มเงื่อนไขกรองต่อท้าย **นอกวงเล็บของเงื่อนไขก่อนหน้า** (ใช้ `&&(...)` แยกเป็นก้อนใหม่ อย่าแทรกเข้าไปในวงเล็บเดิม)
4. [ ] เพิ่ม `updSel('xxx-yyy', uniq(...).filter(Boolean), '')` ใน `initDDs()`
5. [ ] เพิ่ม `setV('xxx-yyy','')` ใน `ovClear()`/`naClear()`/`mnClear()` (บาง tab ใช้ array `.forEach(id=>setV(id,''))` บาง tab เขียนแยกบรรทัด — ต้องดู pattern เดิมก่อนแก้)
6. [ ] รัน `node --check` และ diff-print ฟังก์ชันที่แก้ออกมาดูด้วยตาก่อน save ทุกครั้ง (อย่าเชื่อแค่ assert ผ่าน เพราะ logic ผิดแต่ syntax ถูกก็เป็นไปได้)
7. [ ] เช็คว่า chart ในทุก tab ยังใช้ตัวแปร `f` (ผลลัพธ์จาก `xxxFiltered()`) ไม่ใช่ `AF`/`NA`/`MN` ดิบ — ถ้าใช้ `f` อยู่แล้ว chart จะขยับตาม filter ใหม่โดยอัตโนมัติไม่ต้องแก้เพิ่ม

---

## ข้อควรระวังอื่นๆ

- **Browser cache**: หลัง deploy ขึ้น GitHub แล้ว ให้ผู้ใช้ hard refresh (Ctrl+Shift+R / Cmd+Shift+R) เพราะ browser มักโชว์เวอร์ชันเก่าค้างไว้
- **cust code format**: ใน Excel บาง column เก็บเป็น float string (`"200295872.0"`) บางที่เป็น int string (`"200295872"`) — ต้อง normalize ด้วยฟังก์ชัน `nc()` ก่อน join เสมอ ไม่งั้นจะ join ไม่เจอและกลายเป็น `รอการUpdate`/ค่าว่างผิดๆ
- **ไม่มีระบบ backup อัตโนมัติ** — ถ้าต้องการย้อนกลับเวอร์ชันเก่า ต้องพึ่ง Git history บน GitHub repository เอง (Claude ไม่ได้เก็บ version history ของไฟล์ที่ build ไว้ข้ามแชท)
- **ไฟล์ใหญ่ (Excel) อัปโหลดขึ้น GitHub ไม่ได้ผ่านหน้าเว็บปกติเมื่อเกิน limit การ preview** และการลบไฟล์/repo อาจติดปัญหา **GitHub sudo mode** (ต้องยืนยันตัวตนซ้ำผ่านอีเมล) โดยเฉพาะบน mobile browser — ถ้าเจอปัญหานี้ แนะนำให้ลองบนคอมพิวเตอร์แทน หรือใช้ GitHub Desktop App
- **PDF Export ใช้เทคนิค Blob URL** — สร้าง HTML string ของตารางที่ต้องการ ห่อด้วย CSS สำหรับพิมพ์ แปลงเป็น `Blob`, สร้าง `URL.createObjectURL()`, เปิดแท็บใหม่ด้วย `window.open()`, ฝัง `<script>setTimeout(()=>window.print(),600)</script>` ให้สั่งพิมพ์อัตโนมัติ ไม่ใช้ library ภายนอกอย่าง jsPDF
