import openpyxl
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb_path = 'd:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
wb = openpyxl.load_workbook(wb_path, data_only=True)
ws_hp = wb['hao phi VT']

# Scan hao phi VT rows
nc_groups = {}
may_groups = {}
vl_groups = {}

curr_task_stt = None
curr_task_name = None
curr_task_khuvuc = None

for r in range(7, ws_hp.max_row + 1):
    stt = ws_hp.cell(r, 1).value
    code = ws_hp.cell(r, 2).value
    loai = ws_hp.cell(r, 3).value
    khu_vuc = ws_hp.cell(r, 4).value
    name = ws_hp.cell(r, 5).value
    dvt = ws_hp.cell(r, 6).value
    hs = ws_hp.cell(r, 12).value
    qty_nc = ws_hp.cell(r, 14).value
    qty_may = ws_hp.cell(r, 16).value
    qty_vl = ws_hp.cell(r, 13).value
    qty_dc = ws_hp.cell(r, 15).value

    if loai == 'CÔNG TÁC':
        curr_task_stt = stt
        curr_task_name = name
        curr_task_khuvuc = khu_vuc
    elif loai == 'NHÂN CÔNG':
        # Only include if applied (hs > 0)
        if hs and float(hs) > 0:
            clean_name = name.lstrip('- ').strip()
            nc_groups.setdefault(clean_name, []).append({
                'hp_row': r,
                'stt': curr_task_stt,
                'task_name': curr_task_name,
                'khu_vuc': curr_task_khuvuc,
                'code': code,
                'dvt': dvt,
                'val': qty_nc
            })
    elif loai in ['MÁY THI CÔNG', 'DỤNG CỤ', 'THIẾT BỊ']:
        if hs and float(hs) > 0:
            clean_name = name.lstrip('- ').strip()
            col_letter = 'O' if loai == 'DỤNG CỤ' else 'P'
            val = qty_dc if loai == 'DỤNG CỤ' else qty_may
            may_groups.setdefault(clean_name, []).append({
                'hp_row': r,
                'col_letter': col_letter,
                'stt': curr_task_stt,
                'task_name': curr_task_name,
                'khu_vuc': curr_task_khuvuc,
                'code': code,
                'dvt': dvt,
                'val': val
            })
    elif loai == 'VẬT LIỆU':
        if hs and float(hs) > 0:
            clean_name = name.lstrip('- ').strip()
            vl_groups.setdefault(clean_name, []).append({
                'hp_row': r,
                'col_letter': 'M',
                'stt': curr_task_stt,
                'task_name': curr_task_name,
                'khu_vuc': curr_task_khuvuc,
                'code': code,
                'dvt': dvt,
                'val': qty_vl
            })

print(f"Total NC Master Groups: {len(nc_groups)}")
total_nc_details = sum(len(v) for v in nc_groups.values())
print(f"Total NC Detail Rows: {total_nc_details}")

print(f"\nTotal Máy Master Groups: {len(may_groups)}")
total_may_details = sum(len(v) for v in may_groups.values())
print(f"Total Máy Detail Rows: {total_may_details}")

print(f"\nTotal VL Master Groups: {len(vl_groups)}")
total_vl_details = sum(len(v) for v in vl_groups.values())
print(f"Total VL Detail Rows: {total_vl_details}")
