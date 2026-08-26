"""
update_data.py — สคริปต์แปลงข้อมูล Excel → JSON สำหรับ SLA Dashboard (index.html)

วิธีใช้:
    1. แก้ตัวแปร EXCEL_FILE, REF_DATE_STR ด้านล่างให้ตรงกับไฟล์/วันที่ที่จะอัปเดต
    2. รัน: python3 update_data.py
    3. สคริปต์จะสร้างไฟล์ AF_new.json / NA_new.json / MN_new.json / ZM_new.json / RJ_new.json
       พร้อมพิมพ์รายงานสรุปมาตรฐาน (ตาม README หัวข้อ 9.1) ออกทางหน้าจอ
    4. ตรวจสอบตัวเลขในรายงานให้ตรงกับที่คาดไว้ก่อน แล้วค่อยเขียนทับ AF/NA/MN/ZM/RJ ใน index.html
       (การ integrate เข้า index.html แนะนำให้ทำแยกเป็นอีกขั้นตอน ไม่ auto-write ในสคริปต์นี้
        เพื่อให้มีจังหวะตรวจสอบก่อนเขียนทับไฟล์จริงเสมอ ตาม README 9.1)

⚠️ อ่าน README.md หัวข้อ 7, 7.1, 8, 8.1-8.5, 9 ก่อนใช้สคริปต์นี้เสมอ — มีบั๊ก/ข้อควรระวังที่เจอมาแล้ว
   หลายจุดที่สคริปต์นี้แก้ไว้ให้แล้ว (ดู comment แต่ละจุด) แต่โครงสร้าง sheet ต้นทางอาจเปลี่ยนได้อีก
   ต้องตรวจสอบ dtype/ชื่อคอลัมน์ทุกรอบตามที่ README เตือนไว้

ประวัติ: สคริปต์นี้รวบรวม logic จากการอัปเดตข้อมูลหลายรอบ (18/21/25-08-69) รวมบั๊กที่เจอและแก้แล้ว:
  - ปี พ.ศ./ค.ศ. ปนกันในคอลัมน์ datetime เดียวกัน (CONTRACT_DATE vs DATE_01-11)
  - update field ต้อง coalesce 3 ชั้น
  - NA/MN ต้องยึดฐานจาก sheet ต้นทาง ไม่ใช่ filter จาก AF
  - days_lag/ck_custcare_st ล้ำหน้า 1 สถานะ (แก้ฝั่ง JS ใน index.html ไม่ใช่ตรงนี้)
  - safe_start_con(): วันทำสัญญาของ ZM/RJ มีทั้ง serial number และ datetime แล้วแต่ sheet/รอบข้อมูล
  - safe_yyyymmdd_or_text(): due_first/change_status ของ ZM/RJ สลับบทบาทกันได้ระหว่างรอบข้อมูล
  - sheet_status (RJ) ห้าม hardcode ต้องอ่านจาก CUSTCARE_STATUS_NAME จริงทุกครั้ง
  - guarantee ของ ZM/RJ เปลี่ยนชื่อคอลัมน์จาก CONSTRUCTION_AGENCY.1 เป็น ST_GUARANTEE
  - บ้านเลขที่เพี้ยนเป็นวันที่ (Google Sheets auto-convert) ต้องเช็ค+แก้ทุกรอบ
"""

import pandas as pd
import datetime
import json
import re

# ═══════════════════════════════════════════════════════════
# ตั้งค่าก่อนรันทุกครั้ง
# ═══════════════════════════════════════════════════════════
EXCEL_FILE = '/mnt/user-data/uploads/ส_ญญาย_งไม_ส_งมอบในระบบ-SLA.xlsx'  # ⚠️ แก้ path ให้ตรงกับไฟล์ที่อัปโหลดใหม่ทุกครั้ง
REF_DATE_STR = '2026-08-25'   # ⚠️ ต้องตรงกับ MAX_STATUS_DATE ล่าสุดใน AF sheet (เช็คก่อนทุกครั้ง ไม่ใช่วันที่ upload ไฟล์)
OLD_INDEX_HTML = '/home/claude/index.html'  # ไฟล์ dashboard เดิม (เอาไว้เทียบผลลัพธ์)

# ชื่อ sheet ปัจจุบัน — เคยเปลี่ยนมาแล้ว 1 ครั้ง (รอจัดจ้าง → ขอลดค่าซ่อม) เช็คชื่อ sheet จริงก่อนรันเสมอ
SHEET_AF = 'AF'
SHEET_CK = 'สำหรับกรอง-ตรวจสอบ'   # → NA, และ note ของ AF
SHEET_MN = 'รายเดือน-ชน'          # → MN
SHEET_ZM = 'ซ่อมเสร็จ'            # → ZM (ดำเนินการเสร็จสิ้น)
SHEET_RJ = 'ขอลดค่าซ่อม'          # → RJ (ประสงค์ขอส่วนลด) — เดิมชื่อ "รอจัดจ้าง"

REF = datetime.date.fromisoformat(REF_DATE_STR)

STAGES11 = ['รับคำร้องแจ้งซ่อม','สำรวจซ่อม','ยืนยันการสำรวจซ่อม','ประมาณราคา',
            'ตรวจสอบและรับรองราคา','อนุมัติสรุปค่าก่อสร้าง','รอการจัดจ้าง','จัดหาผู้รับจ้าง',
            'ดำเนินการซ่อม','ดำเนินการเสร็จสิ้น','ขอเบิกเงินค่าจ้าง']


# ═══════════════════════════════════════════════════════════
# Helper functions (แปลงชนิดข้อมูล / วันที่)
# ═══════════════════════════════════════════════════════════
def s(v):
    if pd.isna(v): return ''
    return str(v).strip()

def num_int(v):
    if pd.isna(v): return ''
    return str(int(v))

def num_comma(v, dec=0):
    if pd.isna(v): return ''
    return f'{v:,.0f}' if dec == 0 else f'{v:,.2f}'

def dt_dmy_be(dt):
    """datetime -> 'DD/MM/YY' (พ.ศ.) — auto-detect ว่าปีเก็บเป็น ค.ศ. หรือ พ.ศ.ตรงๆ"""
    if pd.isna(dt): return ''
    dt = pd.Timestamp(dt)
    y = dt.year
    if y < 2100:
        y += 543
    return f'{dt.day:02d}/{dt.month:02d}/{y % 100:02d}'

def yyyymmdd_dmy_be(v):
    """ตัวเลข YYYYMMDD (ค.ศ.) -> 'DD/MM/YY' (พ.ศ.)"""
    if pd.isna(v): return ''
    sv = str(int(v))
    if len(sv) != 8: return ''
    y, mo, d = int(sv[:4]), int(sv[4:6]), int(sv[6:8])
    return f'{d:02d}/{mo:02d}/{(y + 543) % 100:02d}'

def yyyymmdd_to_date(v):
    if pd.isna(v): return None
    sv = str(int(v))
    if len(sv) != 8: return None
    return datetime.date(int(sv[:4]), int(sv[4:6]), int(sv[6:8]))

def parse_dmy(s_):
    """'DD/MM/YY' (พ.ศ.) -> datetime.date"""
    if not s_: return None
    p = s_.split('/')
    if len(p) != 3: return None
    d, mo, yy = int(p[0]), int(p[1]), int(p[2])
    return datetime.date(1957 + yy, mo, d)

def safe_yyyymmdd_or_text(v):
    """รองรับทั้งกรณีเป็นตัวเลข YYYYMMDD และข้อความ (ZM/RJ due_first/change_status สลับบทบาทกันได้)"""
    if pd.isna(v): return ''
    if isinstance(v, str): return v.strip()
    try:
        return yyyymmdd_dmy_be(v)
    except Exception:
        return s(v)

def safe_start_con(v):
    """วันทำสัญญาของ ZM/RJ — รองรับทั้ง serial number ดิบ (int) และ datetime object ที่ parse มาแล้ว
    ⚠️ ห้ามใช้ pd.Timedelta(days=...) บวกตรงๆ จะ overflow เงียบๆ กับตัวเลข serial ใหญ่ (~244000+)
       ต้องใช้ pd.to_datetime(..., origin='1899-12-30', unit='D') เท่านั้น"""
    if pd.isna(v): return ''
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        return dt_dmy_be(v)
    try:
        dt = pd.to_datetime(v, origin='1899-12-30', unit='D')
        return dt_dmy_be(dt)
    except Exception:
        return ''

def status_lag(af_rec):
    """วันค้างสถานะปัจจุบันจริง — ใช้ d0N ของสถานะที่ update อยู่จริง ไม่ใช่ raw days_lag ตรงๆ"""
    if af_rec is None: return None
    upd = af_rec.get('update', '')
    if upd in STAGES11:
        idx = STAGES11.index(upd) + 1
        dt = parse_dmy(af_rec.get(f'd{idx:02d}', ''))
        if dt:
            return (REF - dt).days
    dl = af_rec.get('days_lag', '')
    return int(dl) if str(dl).strip().isdigit() else None

def house_corrupted(house_val):
    """เช็คบ้านเลขที่เพี้ยนเป็นวันที่ (Google Sheets auto-convert เลข 2 ตัวคั่น / เช่น 8/2)"""
    hv = str(house_val or '')
    return bool(re.match(r'^\d{4}-\d{2}-\d{2}', hv)) or '00:00:00' in hv

# รายการบ้านเลขที่เพี้ยนที่เคยเจอและยืนยันค่าที่ถูกต้องแล้ว (แก้อัตโนมัติทุกรอบที่เจอซ้ำ)
KNOWN_HOUSE_FIXES = {
    '200199759': '2/8',   # ยืนยันจาก number_format "d/m" ในไฟล์ xlsx ต้นฉบับ + serial 46236
}


# ═══════════════════════════════════════════════════════════
# 1) AF
# ═══════════════════════════════════════════════════════════
def build_af():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_AF, header=2)
    df_ck = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CK, header=2)

    ck_lookup = {}       # cust -> ปรับสรุปการUPDATE (fallback สำหรับ update field)
    note_lookup = {}     # cust -> หมายเหตุ (join เข้า AF โดยตรง ให้ OVNA ได้ด้วยอัตโนมัติ)
    for _, r in df_ck.iterrows():
        c = num_int(r['รหัสลูกค้า'])
        if c:
            ck_lookup[c] = s(r.get('ปรับสรุปการUPDATE'))
            note_lookup[c] = s(r.get('หมายเหตุ'))

    records = []
    fallback_count = 0
    for _, r in df.iterrows():
        cust = num_int(r['CUST_CODE'])
        if not cust:
            continue

        ck_raw = s(r.get('CK_CUSTCARE_ST'))
        if ck_raw:
            parts = ck_raw.split(' ', 1)
            ck_grp = parts[1] if (len(parts) == 2 and parts[0].isdigit() and len(parts[0]) == 8) else ck_raw
        else:
            ck_grp = 'ไม่เคยซ่อม'

        maxd = yyyymmdd_to_date(r.get('MAX_STATUS_DATE'))
        days_lag = str((REF - maxd).days) if maxd else ''

        # update field: coalesce 3 ชั้น (สำคัญมาก — ห้ามใช้แค่ CUSTCARE_STATUS_NAME อย่างเดียว)
        custcare_status = s(r.get('CUSTCARE_STATUS_NAME'))
        if custcare_status:
            upd = custcare_status
        elif cust in ck_lookup and ck_lookup[cust]:
            upd = ck_lookup[cust]
        else:
            upd = s(r.get('NEXT_CUSTCARE'))
            fallback_count += 1

        house = s(r.get('HOUSE_NO_ACT'))
        if cust in KNOWN_HOUSE_FIXES and house_corrupted(house):
            house = KNOWN_HOUSE_FIXES[cust]

        rec = {
            'dept': s(r.get('DEPT_SHORT_NAME')), 'bp': s(r.get('BP_SHORT_NAME')), 'proj': s(r.get('PROJ_NAME')),
            'house': house, 'cust': cust, 'cust_name': s(r.get('CUSTNAME1DISPLAY')),
            'change_status': s(r.get('CHANGE_STATUS')), 'status': s(r.get('NEXT_CUSTCARE')), 'update': upd,
            'guarantee': s(r.get('ST_GUARANTEE')), 'confirm_flag': s(r.get('CONFIRM_FLAG')),
            'next_custcare': s(r.get('NEXT_CUSTCARE')),
            'start_con': dt_dmy_be(r.get('CONTRACT_DATE')), 'due_last': yyyymmdd_dmy_be(r.get('Dueล่าสุด')),
            'due_amt': num_comma(r.get('เงินDue'), 2), 'overdue': num_int(r.get('งวดค้าง')),
            'debt': num_comma(r.get('หนี้ค้าง'), 2), 'days_lag': days_lag,
            'ck_custcare_st': ck_raw, 'ck_custcare_st_grp': ck_grp,
            'd01': dt_dmy_be(r.get('DATE_01')), 'd02': dt_dmy_be(r.get('DATE_02')), 'd03': dt_dmy_be(r.get('DATE_03')),
            'd04': dt_dmy_be(r.get('DATE_04')), 'd05': dt_dmy_be(r.get('DATE_05')), 'd06': dt_dmy_be(r.get('DATE_06')),
            'd07': dt_dmy_be(r.get('DATE_07')), 'd08': dt_dmy_be(r.get('DATE_08')), 'd09': dt_dmy_be(r.get('DATE_09')),
            'd10': dt_dmy_be(r.get('DATE_10')), 'd11': dt_dmy_be(r.get('DATE_11')),
            'aging09': num_int(r.get('AGING_09')), 'ck_aging09': s(r.get('CK_AGING_09')),
            'contract_type': s(r.get('CONTRACT_TYPE1')),
            'note': note_lookup.get(cust, ''),   # ← join จาก สำหรับกรอง-ตรวจสอบ โดยตรง (ไม่ใช่แค่ NA)
        }
        records.append(rec)

    return records, fallback_count


# ═══════════════════════════════════════════════════════════
# 2) NA (ไม่ดำเนินการ สคล/สคจ ชี้แจง) — ยึดฐานจาก sheet สำหรับกรอง-ตรวจสอบ ไม่ใช่ filter จาก AF
# ═══════════════════════════════════════════════════════════
def build_na(af_by_cust, na_fields):
    df_ck = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_CK, header=2)
    na_new = []
    found_cnt = not_found_cnt = 0
    for _, r in df_ck.iterrows():
        cust = num_int(r['รหัสลูกค้า'])
        if not cust:
            continue
        note = s(r.get('หมายเหตุ'))
        if cust in af_by_cust:
            rec = dict(af_by_cust[cust])
            found_cnt += 1
        else:
            upd = s(r.get('ปรับสรุปการUPDATE'))
            rec = {k: '' for k in na_fields}
            rec.update({
                'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
                'house': s(r.get('บ้านเลขที่')), 'cust': cust, 'cust_name': s(r.get('ชื่อลูกค้า')),
                'change_status': s(r.get('CHANGE_STATUS')), 'status': upd, 'update': upd,
                'confirm_flag': s(r.get('CONFIRM_FLAG')), 'next_custcare': upd,
                'ck_custcare_st_grp': 'ไม่เคยซ่อม',
            })
            not_found_cnt += 1
        rec['note'] = note
        na_new.append(rec)
    return na_new, found_cnt, not_found_cnt


# ═══════════════════════════════════════════════════════════
# 3) MN (ข้อมูลหน่วยงานอื่น) — ยึดฐานจาก sheet รายเดือน-ชน
# ═══════════════════════════════════════════════════════════
def build_mn(af_by_cust):
    df_mn = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_MN, header=60)  # ⚠️ header row เคยขยับ เช็คทุกรอบ
    mn_fields = ['dept','bp','proj','house','cust','update','guarantee','start_con','due_last','due_amt',
                 'overdue','debt','days_lag','ck_custcare_st','ck_custcare_st_grp','d01','d02','d03','d04',
                 'd05','d06','d07','d08','d09','d10','d11','aging09','ck_aging09','contract_type']
    mn_new = []
    found_in_af = not_found = 0
    for _, r in df_mn.iterrows():
        cust = num_int(r['รหัสลูกค้า'])
        if not cust:
            continue
        if cust in af_by_cust:
            af_r = af_by_cust[cust]
            rec = {k: af_r[k] for k in mn_fields}
            found_in_af += 1
        else:
            house = s(r.get('บ้านเลขที่'))
            if cust in KNOWN_HOUSE_FIXES and house_corrupted(house):
                house = KNOWN_HOUSE_FIXES[cust]
            rec = {
                'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
                'house': house, 'cust': cust, 'update': s(r.get('ปรับสรุปการUPDATE')),
                'guarantee': 'หลังค้ำประกัน', 'start_con': '', 'due_last': '', 'due_amt': '', 'overdue': '',
                'debt': '', 'days_lag': '', 'ck_custcare_st': '', 'ck_custcare_st_grp': 'ไม่เคยซ่อม',
                'd01': '', 'd02': '', 'd03': '', 'd04': '', 'd05': '', 'd06': '', 'd07': '', 'd08': '',
                'd09': '', 'd10': '', 'd11': '', 'aging09': '', 'ck_aging09': '', 'contract_type': '',
            }
            not_found += 1
        mn_new.append(rec)
    return mn_new, found_in_af, not_found


# ═══════════════════════════════════════════════════════════
# 4) ZM (ดำเนินการเสร็จสิ้น) และ RJ (ประสงค์ขอส่วนลด)
# ═══════════════════════════════════════════════════════════
def build_zm_rj_base(sheet, p_col, af_by_cust):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=3)
    recs = []
    for _, r in df.iterrows():
        cust = s(int(r['รหัสลูกค้า'])) if pd.notna(r['รหัสลูกค้า']) else ''
        if not cust:
            continue
        af_r = af_by_cust.get(cust)
        house = s(r.get('บ้านเลขที่'))
        if cust in KNOWN_HOUSE_FIXES and house_corrupted(house):
            house = KNOWN_HOUSE_FIXES[cust]
        rec = {
            'dept': s(r.get('กอง')), 'bp': s(r.get('สำนักงาน')), 'proj': s(r.get('โครงการ')),
            'house': house, 'cust': cust, 'cust_name': s(r.get('ชื่อลูกค้า')),
            'start_con': safe_start_con(r.get('วันที่ทำสัญญา')),
            'due_first': safe_yyyymmdd_or_text(r.get('วันที่นัด\nส่งมอบครั้งแรก')),
            'due_new': yyyymmdd_dmy_be(r.get('วันที่นัด\nส่งมอบครั้งใหม่')),
            'change_status': safe_yyyymmdd_or_text(r.get('CHANGE_STATUS')),
            'agency': s(r.get('CONSTRUCTION_AGENCY')),
            'guarantee': s(r.get('ST_GUARANTEE')),   # ⚠️ ชื่อคอลัมน์เปลี่ยนจาก CONSTRUCTION_AGENCY.1 แล้ว
            'o_letter': s(r.get('ส่งจดหมาย')), 'p_extra': s(r.get(p_col)), 'q_appt': s(r.get('นัดรับมอบ')),
            'note': s(r.get('หมายเหตุ')),
            'sheet_custcare_status': s(r.get('CUSTCARE_STATUS_NAME')),  # ห้าม hardcode! อ่านค่าจริงเสมอ
            'contract_type': af_r.get('contract_type', '') if af_r else '',
            'delivery_status': 'พบใน CT_NotSend' if af_r else 'ไม่พบใน CT_NotSend',
        }
        recs.append((rec, af_r))
    return recs

def build_zm(af_by_cust):
    pairs = build_zm_rj_base(SHEET_ZM, 'แจ้งซ่อมเพิ่มเติม', af_by_cust)
    ZM = []
    for rec, af_r in pairs:
        sl = status_lag(af_r)
        rec['status_lag'] = str(sl) if sl is not None else ''
        del rec['sheet_custcare_status']  # ZM ไม่ใช้ฟิลด์นี้
        ZM.append(rec)
    return ZM

def build_rj(af_by_cust):
    pairs = build_zm_rj_base(SHEET_RJ, 'ใบตอบรับ', af_by_cust)
    RJ = []
    for rec, af_r in pairs:
        rec['sheet_status'] = rec.pop('sheet_custcare_status')  # ← อ่านจาก sheet จริง ไม่ hardcode
        if af_r:
            rec['af_status'] = af_r.get('update', '')
            sl = status_lag(af_r)
            rec['af_status_lag'] = str(sl) if sl is not None else ''
        else:
            rec['af_status'] = ''
            rec['af_status_lag'] = ''
        RJ.append(rec)
    return RJ


# ═══════════════════════════════════════════════════════════
# Main — build ทั้งหมด + รายงานสรุปมาตรฐาน (README หัวข้อ 9.1)
# ═══════════════════════════════════════════════════════════
def main():
    print(f'=== อัปเดตข้อมูลจากไฟล์: {EXCEL_FILE} ===')
    print(f'=== วันที่อ้างอิง (REF_DATE): {REF_DATE_STR} — ต้องตรงกับ MAX_STATUS_DATE ล่าสุดใน AF ===\n')

    AF, af_fallback = build_af()
    af_by_cust = {r['cust']: r for r in AF}
    print(f'AF: {len(AF)} รายการ | update fallback (ไม่มีทั้ง CUSTCARE_STATUS_NAME และ ปรับสรุปการUPDATE): {af_fallback}')

    NA, na_found, na_not_found = build_na(af_by_cust, list(AF[0].keys()))
    print(f'NA: {len(NA)} รายการ | พบใน AF: {na_found} | ไม่พบ: {na_not_found}')

    MN, mn_found, mn_not_found = build_mn(af_by_cust)
    print(f'MN: {len(MN)} รายการ | พบใน AF: {mn_found} | ไม่พบ: {mn_not_found}')

    ZM = build_zm(af_by_cust)
    zm_delivered = sum(1 for r in ZM if r['delivery_status'] == 'ไม่พบใน CT_NotSend')
    print(f'ZM: {len(ZM)} รายการ | ไม่พบใน CT_NotSend(หลุดจาก AF): {zm_delivered}')

    RJ = build_rj(af_by_cust)
    rj_changed = sum(1 for r in RJ if r['af_status'] and r['af_status'] != r['sheet_status'])
    print(f'RJ: {len(RJ)} รายการ | สถานะเปลี่ยนจาก sheet_status แล้ว: {rj_changed}')

    # เช็คบ้านเลขที่เพี้ยน
    print('\n--- เช็คบ้านเลขที่เพี้ยนเป็นวันที่ ---')
    for name, arr in [('AF', AF), ('NA', NA), ('MN', MN), ('ZM', ZM), ('RJ', RJ)]:
        bad = [(r['cust'], r['house']) for r in arr if house_corrupted(r.get('house', ''))]
        print(f'{name}: {len(bad)} จุด' + (f' -> {bad}' if bad else ''))
        if bad:
            print('  ⚠️ เจอ cust ใหม่ที่ไม่อยู่ใน KNOWN_HOUSE_FIXES — ต้องตรวจสอบค่าที่ถูกต้องเองก่อนแก้ (ดู README หัวข้อ 7 ข้อ 5)')

    # no-op / สถานะจริงปนอยู่ (README หัวข้อ 4.3)
    print('\n--- เช็ค record ที่มีสถานะจริงปนอยู่ (README 4.3) ---')
    ovna = [r for r in AF if r['status'] == 'ไม่ดำเนินการ']
    print(f'OVNA (derive จาก AF.filter): {len(ovna)} | ในสถานะจริง (ต้องเป็น 0 เสมอ): {sum(1 for r in ovna if r["update"] in STAGES11)}')
    print(f'NA: {len(NA)} | ในสถานะจริงปนอยู่: {sum(1 for r in NA if r["update"] in STAGES11)}')
    print(f'MN: {len(MN)} | ในสถานะจริงปนอยู่: {sum(1 for r in MN if r["update"] in STAGES11)}')

    # เทียบกับข้อมูลเดิมในไฟล์ (ถ้ามี)
    try:
        with open(OLD_INDEX_HTML, encoding='utf-8') as f:
            old_content = f.read()
        old_af = json.loads(re.search(r'const AF=(\[.*?\]);', old_content).group(1))
        old_by_cust = {str(r['cust']): r for r in old_af}
        common = [c for c in af_by_cust if c in old_by_cust]
        print(f'\n--- เทียบกับ AF เดิมในไฟล์ ---')
        print(f'ตรงกัน: {len(common)}/{len(AF)} (เดิมมี {len(old_af)})')
        for f in ['dept', 'update', 'contract_type', 'd01', 'debt', 'guarantee']:
            match = sum(1 for c in common if str(old_by_cust[c][f]) == str(af_by_cust[c][f]))
            print(f'  {f}: {match}/{len(common)} ({match/len(common)*100:.1f}%)')
    except Exception as e:
        print(f'\n(ข้ามการเทียบกับไฟล์เดิม: {e})')

    # เขียนไฟล์ผลลัพธ์
    for name, data in [('AF', AF), ('NA', NA), ('MN', MN), ('ZM', ZM), ('RJ', RJ)]:
        with open(f'{name}_new.json', 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False)
    print('\n=== เขียนไฟล์ AF_new.json / NA_new.json / MN_new.json / ZM_new.json / RJ_new.json เรียบร้อย ===')
    print('=== ตรวจสอบตัวเลขข้างบนให้ครบตามรายงานสรุปมาตรฐาน (README 9.1) ก่อน integrate เข้า index.html ===')


if __name__ == '__main__':
    main()
