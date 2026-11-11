"""อ่าน Excel 2 ไฟล์ แล้วเขียน data.json
ใช้:  python scripts/build_data.py data site/data.json
  data/budget.xlsx       ชีท "ใช้งาน"  (ข้อมูลงบประมาณ)
  data/KPI_summaryv2.xlsx  ชีท "สรุปรายหน่วยงาน" และ "รายตัวชี้วัด"  (ผลการดำเนินงาน)
"""
import sys, os, json, itertools, collections
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
TYPE = {'เชิงปริมาณ': 'quant', 'เชิงคุณภาพ': 'qual'}
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

os.makedirs(os.path.dirname(out_path) or '.', exist_ok=True)
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump({'data': mins, 'perf': PERF}, f, ensure_ascii=False, separators=(',', ':'))
print(f'OK: {len(mins)} กลุ่ม, {sum(len(m["depts"]) for m in mins)} หน่วยงาน, KPI {len(PERF)} หน่วย -> {out_path}')
