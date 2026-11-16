"""อ่าน Excel 2 ไฟล์ แล้วเขียน data.json
ใช้:  python scripts/build_data.py data site/data.json
  data/budget.xlsx       ชีท "ใช้งาน"  (ข้อมูลงบประมาณ)
  data/KPI_summaryv2.xlsx  ชีท "สรุปรายหน่วยงาน" และ "รายตัวชี้วัด"  (ผลการดำเนินงาน)
"""
import sys, os, json, itertools, collections, re
import openpyxl

def die(msg):
    print('ERROR:', msg); sys.exit(1)

data_dir = sys.argv[1] if len(sys.argv) > 1 else 'data'
out_path = sys.argv[2] if len(sys.argv) > 2 else 'data.json'

def num(x):
    try: return round(float(x or 0), 4)      # ค่าว่าง, '-', #N/A -> 0
    except (ValueError, TypeError): return 0.0

# ---------- 1) งบประมาณ ----------
wb = openpyxl.load_workbook(os.path.join(data_dir, 'budget.xlsx'), data_only=True)
if 'ใช้งาน' not in wb.sheetnames: die('budget.xlsx ไม่มีชีท "ใช้งาน"')
ws = wb['ใช้งาน']
head = [str(c or '').strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]
expect = {0:'รหัสกรม', 1:'ชื่อกรม', 2:'วงเงิน พ.ร.บ.', 3:'วงเงินสุทธิหลังโอน', 4:'พรบ โอน',
          5:'วงเงิน ภายหลังพ.ร.บ.', 8:'เบิกจ่าย', 17:'PO', 18:'สำรองเงิน', 19:'การใช้จ่าย',
          28:'เงินกัน68', 29:'เบิกจ่ายเงินกัน68', 31:'โอน68'}
for i, name in expect.items():
    if i >= len(head) or head[i] != name:
        die(f'คอลัมน์ที่ {i+1} ควรชื่อ "{name}" แต่พบ "{head[i] if i < len(head) else None}" (โครงสร้างชีทเปลี่ยน)')
rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
if not rows: die('ชีท "ใช้งาน" ไม่มีข้อมูล')

def dept(r):
    return [str(r[0]), str(r[1]).strip(), num(r[2]), num(r[3]), num(r[8]), num(r[17]), num(r[18]),
            num(r[19]), num(r[27]), num(r[28]), num(r[29]), num(r[30]), num(r[31]), num(r[4]), num(r[5])]

mins = []
for k, g in itertools.groupby(rows, key=lambda r: str(r[0])[:2]):
    g = [dept(r) for r in g]
    if k in ('50', '80') and mins:            # รัฐวิสาหกิจ/กองทุน รวมเข้ากับกระทรวงก่อนหน้า
        mins[-1]['depts'] += g; continue
    n = g[0][1]
    for p in ('สำนักงานปลัด', 'สํานักงานปลัด'):
        if n.startswith(p): n = n[len(p):]; break
    n = {'70': 'จังหวัด', '75': 'กรุงเทพมหานครและองค์กรปกครองส่วนท้องถิ่น'}.get(k, n)
    mins.append({'code': k, 'name': n, 'depts': g})

# ---------- 2) ผลการดำเนินงาน (KPI) ----------
# กติกานับตัวชี้วัด (สำคัญ):
#   - นับเฉพาะตัวที่ "มีผลงาน" (คอลัมน์ ผลงาน เป็นตัวเลข รวมค่า 0 ซึ่งถือว่าได้ 0 จริง = ไม่บรรลุ)
#   - ผลงานที่เป็นค่าว่าง / - / N/A / ข้อความ (เช่น "รายงานผลไตรมาส 4", "รอรายงาน") = ตัวชี้วัดไม่สมบูรณ์ ไม่นับ
#   - ผลงานแบบอัตราส่วน/หลายค่า เช่น 18/1, 563,060/50,000 ถือว่ามีผลงาน (นับ)
#   - ดังนั้น จำนวนทั้งหมด = บรรลุ + ไม่บรรลุ เสมอ
#   - ร้อยละที่ได้ = ค่าเฉลี่ยของ "ร้อยละที่ใช้คำนวณ (สูงสุด 100)" เฉพาะตัวที่นับ
kw = openpyxl.load_workbook(os.path.join(data_dir, 'KPI_summaryv2.xlsx'), data_only=True)
if 'รายตัวชี้วัด' not in kw.sheetnames: die('KPI_summaryv2.xlsx ไม่มีชีท "รายตัวชี้วัด"')
TYPE = {
    'เชิงปริมาณ': 'quant', 'ปริมาณ': 'quant',
    'เชิงคุณภาพ': 'qual',  'คุณภาพ': 'qual',
    'เชิงเวลา': 'time',    'เวลา': 'time',
    'เชิงต้นทุน': 'cost',  'ต้นทุน': 'cost',
}
NUMLIKE = re.compile(r'^[\d.,/\-\s%]+$')          # ตัวเลข หรืออัตราส่วน เช่น 18/1  563,060/50,000

def has_result(v):
    """ผลงานนี้ถือว่ามีผลหรือไม่ (0 ถือว่ามี)"""
    if v is None or isinstance(v, bool): return False
    if isinstance(v, (int, float)): return True
    s = str(v).strip()
    if not s or s.upper() in ('-', 'N/A', 'NA', '#N/A', 'NONE'): return False
    return bool(NUMLIKE.match(s)) and any(ch.isdigit() for ch in s)

def code5(v):
    s = str(v).split()[0].strip()
    return s.zfill(5) if s.isdigit() else s

acc = collections.defaultdict(lambda: collections.defaultdict(list))   # code -> type -> [(achieved, pct)]
skipped = collections.Counter()          # code -> จำนวนที่ไม่นับ
skipped_k = collections.Counter()        # (code, type) -> จำนวนที่ไม่นับ
for r in kw['รายตัวชี้วัด'].iter_rows(min_row=2, values_only=True):
    if not r[0] or r[5] not in TYPE: continue
    code, k = code5(r[0]), TYPE[r[5]]
    if not has_result(r[7]):
        skipped[code] += 1; skipped_k[(code, k)] += 1; continue
    pc = r[9] if isinstance(r[9], (int, float)) and not isinstance(r[9], bool) else 0.0
    acc[code][k].append((r[10] == 'บรรลุ', max(0.0, min(100.0, float(pc)))))

PERF = {}
for code, kinds in acc.items():
    p, allpc = {}, []
    for k, items in kinds.items():
        ach = sum(1 for a, _ in items if a)
        pcs = [x for _, x in items]
        # [จำนวนทั้งหมด(=บรรลุ+ไม่บรรลุ), บรรลุ, ไม่บรรลุ, ไม่นับ(ไม่มีผลงาน), ร้อยละเฉลี่ย]
        p[k] = [len(items), ach, len(items) - ach, skipped_k.get((code, k), 0), round(sum(pcs) / len(pcs), 4)]
        allpc += pcs
    p['pct'] = round(sum(allpc) / len(allpc), 4) if allpc else 0.0
    PERF[code] = p
for code in skipped:                       # หน่วยที่ทุกตัวไม่มีผลงาน
    PERF.setdefault(code, {'pct': 0.0})
known = {d[0] for m in mins for d in m['depts']}
missing = [c for c in PERF if c not in known]
if missing: print('WARNING: รหัสใน KPI ที่ไม่พบในไฟล์งบ:', missing)
print(f'  KPI: ตัวชี้วัดที่ไม่นับ (ไม่มีผลงาน) {sum(skipped.values())} ตัว ใน {len(skipped)} หน่วยงาน')

# ---------- 3) ชื่อกระทรวงที่ถูกต้อง (ตามรหัส 2 หลักแรก) ----------
# แก้เฉพาะชื่อกระทรวงที่แสดงในดรอปดาวน์ ลำดับ/การจัดกลุ่ม/ข้อมูลอื่นไม่เปลี่ยน
MIN_NAME = {
    '01': 'สำนักนายกรัฐมนตรี',
    '02': 'กระทรวงกลาโหม',
    '03': 'กระทรวงการคลัง',
    '04': 'กระทรวงการต่างประเทศ',
    '05': 'กระทรวงการท่องเที่ยวและกีฬา',
    '06': 'กระทรวงการพัฒนาสังคมและความมั่นคงของมนุษย์',
    '07': 'กระทรวงเกษตรและสหกรณ์',
    '08': 'กระทรวงคมนาคม',
    '09': 'กระทรวงทรัพยากรธรรมชาติ และสิ่งแวดล้อม',
    '11': 'กระทรวงดิจิทัลเพื่อเศรษฐกิจและสังคม',
    '12': 'กระทรวงพลังงาน',
    '13': 'กระทรวงพาณิชย์',
    '15': 'กระทรวงมหาดไทย',
    '16': 'กระทรวงยุติธรรม',
    '17': 'กระทรวงแรงงาน',
    '18': 'กระทรวงวัฒนธรรม',
    '20': 'กระทรวงศึกษาธิการ',
    '21': 'กระทรวงสาธารณสุข',
    '22': 'กระทรวงอุตสาหกรรม',
    '23': 'กระทรวงการอุดมศึกษา วิทยาศาสตร์ วิจัยและนวัตกรรม',
    '25': 'ส่วนราชการไม่สังกัดสำนักนายกรัฐมนตรี กระทรวง หรือทบวง และหน่วยงานภายใต้การควบคุมดูแลของนายกรัฐมนตรี',
    '27': 'หน่วยงานของรัฐสภา',
    '28': 'หน่วยงานของศาล',
    '29': 'หน่วยงานขององค์กรอิสระและองค์กรอัยการ',
    '56': 'ส่วนราชการในพระองค์',
    '60': 'สภากาชาดไทย',
    '65': 'หน่วยงานอื่นของรัฐ',
    '70': 'จังหวัดและกลุ่มจังหวัด',
    '75': 'องค์กรปกครองส่วนท้องถิ่น',
    '90': 'งบกลาง',
    '95': 'รายจ่ายเพื่อชดใช้เงินคงคลัง',
}
for m in mins:
    new = MIN_NAME.get(m['code'])
    if not new:
        print(f'WARNING: ไม่มีชื่อกระทรวงรหัส {m["code"]} ในตาราง MIN_NAME คงชื่อเดิม "{m["name"]}"'); continue
    if new != m['name']: print(f'  เปลี่ยนชื่อกระทรวง {m["code"]}: "{m["name"]}" -> "{new}"')
    m['name'] = new

os.makedirs(os.path.dirname(out_path) or '.', exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump({'data': mins, 'perf': PERF}, f, ensure_ascii=False, separators=(',', ':'))
print(f'OK: {len(mins)} กลุ่ม, {sum(len(m["depts"]) for m in mins)} หน่วยงาน, KPI {len(PERF)} หน่วย -> {out_path}')
