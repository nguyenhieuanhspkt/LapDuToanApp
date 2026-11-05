import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb_path = 'd:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
wb = openpyxl.load_workbook(wb_path, data_only=True)
ws = wb['hao phi VT']

print(f"Total rows in 'hao phi VT': {ws.max_row}")

nc_items = {}
may_items = {}
vl_items = {}

curr_task_stt = None
curr_task_name = None
curr_task_q = None

for r in range(7, ws.max_row + 1):
    stt = ws.cell(r, 1).value
    code = ws.cell(r, 2).value
    loai = ws.cell(r, 3).value
    khu_vuc = ws.cell(r, 4).value
    name = ws.cell(r, 5).value
    dvt = ws.cell(r, 6).value
    q = ws.cell(r, 7).value
    hs = ws.cell(r, 12).value
    
    if loai == 'CÔNG TÁC':
        curr_task_stt = stt
        curr_task_name = name
        curr_task_q = q
    elif loai == 'NHÂN CÔNG':
        nc_items.setdefault(name, []).append({
            'row': r,
            'stt': curr_task_stt,
            'task_name': curr_task_name,
            'code': code,
            'dvt': dvt,
            'hs': hs
        })
    elif loai in ['MÁY THI CÔNG', 'DỤNG CỤ', 'THIẾT BỊ']:
        may_items.setdefault(name, []).append({
            'row': r,
            'stt': curr_task_stt,
            'task_name': curr_task_name,
            'code': code,
            'dvt': dvt,
            'hs': hs
        })
    elif loai == 'VẬT LIỆU':
        vl_items.setdefault(name, []).append({
            'row': r,
            'stt': curr_task_stt,
            'task_name': curr_task_name,
            'code': code,
            'dvt': dvt,
            'hs': hs
        })

print(f"\n--- NHÂN CÔNG ({len(nc_items)} unique types) ---")
for k, v in nc_items.items():
    applied_count = sum(1 for item in v if item['hs'] and float(item['hs']) > 0)
    print(f"  * {k[:50]}: {len(v)} occurrences (applied: {applied_count})")

print(f"\n--- MÁY THI CÔNG ({len(may_items)} unique types) ---")
for k, v in may_items.items():
    applied_count = sum(1 for item in v if item['hs'] and float(item['hs']) > 0)
    print(f"  * {k[:50]}: {len(v)} occurrences (applied: {applied_count})")

print(f"\n--- VẬT LIỆU ({len(vl_items)} unique types) ---")
for k, v in vl_items.items():
    applied_count = sum(1 for item in v if item['hs'] and float(item['hs']) > 0)
    print(f"  * {k[:50]}: {len(v)} occurrences (applied: {applied_count})")
