"""อ่าน Excel 2 ไฟล์ แล้วเขียน data.json
ใช้:  python scripts/build_data.py data site/data.json
  data/budget.xlsx       ชีท "ใช้งาน"  (ข้อมูลงบประมาณ)
  data/KPI_summaryv2.xlsx  ชีท "สรุปรายหน่วยงาน" และ "รายตัวชี้วัด"  (ผลการดำเนินงาน)
                           ชีท "Sheet1" (ชื่อกระทรวงที่ถูกต้อง ใช้แทนชื่อกระทรวงในดรอปดาวน์)
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
kw = openpyxl.load_workbook(os.path.join(data_dir, 'KPI_summaryv2.xlsx'), data_only=True)
for s in ('สรุปรายหน่วยงาน', 'รายตัวชี้วัด'):
    if s not in kw.sheetnames: die(f'KPI_summaryv2.xlsx ไม่มีชีท "{s}"')
det = collections.defaultdict(lambda: collections.defaultdict(list))
# TYPE = {'เชิงปริมาณ': 'quant', 'เชิงคุณภาพ': 'qual'}
TYPE = {
    'เชิงปริมาณ': 'quant',
    'ปริมาณ': 'quant',
    'เชิงคุณภาพ': 'qual',
    'คุณภาพ': 'qual'
}
for r in kw['รายตัวชี้วัด'].iter_rows(min_row=2, values_only=True):
    if r[0] and r[5] in TYPE: det[str(r[0]).split()[0]][TYPE[r[5]]].append(r)
PERF = {}
for s in kw['สรุปรายหน่วยงาน'].iter_rows(min_row=2, values_only=True):
    if not s[0] or not isinstance(s[8], (int, float)): continue   # ข้ามแถวรวม/หมายเหตุ
    code = str(s[0]).split()[0]
    p = {'pct': round(float(s[8]), 4)}
    for k, rs in det[code].items():
        pc = [x[9] for x in rs if isinstance(x[9], (int, float))]
        p[k] = [len(rs), sum(x[10] == 'บรรลุ' for x in rs), sum(x[10] == 'ไม่บรรลุ' for x in rs),
                sum(x[10] == 'ไม่มีผล' for x in rs), round(sum(pc) / len(pc), 4) if pc else 0]
    PERF[code] = p
known = {d[0] for m in mins for d in m['depts']}
missing = [c for c in PERF if c not in known]
if missing: print('WARNING: รหัสใน KPI ที่ไม่พบในไฟล์งบ:', missing)

# ---------- 3) ชื่อกระทรวงที่ถูกต้องจาก KPI_summaryv2.xlsx ชีท "Sheet1" ----------
# แก้เฉพาะชื่อกระทรวงที่แสดงในดรอปดาวน์ (รหัสกลุ่ม ลำดับ และข้อมูลอื่นไม่เปลี่ยน)
def _txt(v): return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''
def _code(v):
    if isinstance(v, (int, float)) and float(v) == int(v): return str(int(v)).zfill(5)
    t = _txt(v).split(' ')[0] if _txt(v) else ''
    return t if re.fullmatch(r'\d{5}', t) else None

def ministry_map(ws):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    ncol = max((len(r) for r in rows), default=0)
    rows = [r + [None] * (ncol - len(r)) for r in rows]
    # คอลัมน์รหัสหน่วยงาน = คอลัมน์ที่มีรหัส 5 หลักมากที่สุด
    cnt = [sum(1 for r in rows if _code(r[c])) for c in range(ncol)]
    if not cnt or max(cnt) == 0: die('Sheet1 ไม่พบคอลัมน์รหัสหน่วยงาน 5 หลัก')
    ccode = cnt.index(max(cnt))
    # คอลัมน์กระทรวง = หัวคอลัมน์ที่มีคำว่า "กระทรวง" (สั้นๆ) ถ้าไม่มี ใช้คอลัมน์ที่ขึ้นต้นด้วย "กระทรวง" มากที่สุด
    cmin = None
    for r in rows[:15]:
        for c, v in enumerate(r):
            t = _txt(v)
            if c != ccode and t and len(t) <= 25 and 'กระทรวง' in t and not _code(v): cmin = c; break
        if cmin is not None: break
    if cmin is None:
        sc = [(sum(1 for r in rows if _txt(r[c]).startswith('กระทรวง')), c) for c in range(ncol) if c != ccode]
        if not sc or max(sc)[0] == 0: die('Sheet1 ไม่พบคอลัมน์ชื่อกระทรวง')
        cmin = max(sc)[1]
    print(f'Sheet1: คอลัมน์รหัส = {ccode+1}, คอลัมน์กระทรวง = {cmin+1}')
    out, last = {}, ''
    for r in rows:
        c = _code(r[ccode])
        if not c: continue
        m = _txt(r[cmin]) or last          # เซลล์ที่ merge/ว่าง ใช้ค่าก่อนหน้า
        if m: out[c] = last = m
    return out

if 'Sheet1' not in kw.sheetnames: die('KPI_summaryv2.xlsx ไม่มีชีท "Sheet1" (ชีทที่มี: %s)' % ', '.join(kw.sheetnames))
MIN = ministry_map(kw['Sheet1'])
for m in mins:
    votes = collections.Counter(MIN[d[0]] for d in m['depts'] if d[0][:2] == m['code'] and d[0] in MIN)
    if not votes:
        print(f'WARNING: กระทรวง {m["code"]} "{m["name"]}" ไม่พบใน Sheet1 คงชื่อเดิม'); continue
    new = votes.most_common(1)[0][0]
    if new != m['name']: print(f'  เปลี่ยนชื่อกระทรวง {m["code"]}: "{m["name"]}" -> "{new}"')
    m['name'] = new

os.makedirs(os.path.dirname(out_path) or '.', exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump({'data': mins, 'perf': PERF}, f, ensure_ascii=False, separators=(',', ':'))
print(f'OK: {len(mins)} กลุ่ม, {sum(len(m["depts"]) for m in mins)} หน่วยงาน, KPI {len(PERF)} หน่วย -> {out_path}')
