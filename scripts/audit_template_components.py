import openpyxl
import sys
sys.path.append('.')

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_dm = wb['Định mức áp dụng']

print("Auditing all potential references in Định mức áp dụng...")

# Let's inspect all rows in ws_dm from row 45 to 250
dm_rows = {}
for r in range(45, ws_dm.max_row + 1):
    desc = str(ws_dm.cell(r, 2).value or '').strip()
    unit = str(ws_dm.cell(r, 3).value or '').strip()
    val = ws_dm.cell(r, 4).value
    stt = str(ws_dm.cell(r, 1).value or '').strip()
    dm_rows[r] = {'stt': stt, 'desc': desc, 'unit': unit, 'val': val}

# Now let's list every single template entry in get_components_for_code
# and verify if its link points to the exact right row, right unit, and right resource!

from scripts.build_hao_phi_vt_856 import get_components_for_code

# Test with all unique codes from 02_Chi_Tiet_856_Dich_Vu
ws_02 = wb['02_Chi_Tiet_856_Dich_Vu']
test_codes = set()
for r in range(5, ws_02.max_row + 1):
    c = ws_02.cell(r, 9).value
    if c:
        test_codes.add(str(c).strip())

print(f"Testing {len(test_codes)} unique codes...")

mismatches = []
verified = []

for code in sorted(test_codes):
    components = get_components_for_code(code)
    for comp in components:
        comp_code, comp_type, stage, comp_name, unit, val_col, norm_val, coeff, basis = comp
        if isinstance(norm_val, str) and 'Định mức áp dụng' in norm_val:
            ref = norm_val.replace('=', '').replace("'", '').replace('Định mức áp dụng!', '').strip()
            # ref like D61
            row_idx = int(ref.replace('D', ''))
            target = dm_rows.get(row_idx, {})
            t_desc = target.get('desc', '')
            t_unit = target.get('unit', '')
            t_val = target.get('val', '')
            
            # Validation rules:
            # 1. Unit must match
            # 2. Type must match
            err = None
            if comp_type == 'NHÂN CÔNG':
                if t_unit != 'công':
                    err = f"Unit mismatch: comp={unit} vs target={t_unit}"
                elif not any(x in t_desc.lower() for x in ['nhân công', 'thợ', 'kỹ sư', 'bậc', 'bước', 'công']):
                    err = f"Description not Nhân công: {t_desc}"
            elif comp_type == 'MÁY THI CÔNG':
                if t_unit != 'ca':
                    err = f"Unit mismatch: comp={unit} vs target={t_unit}"
                elif any(x in t_desc.lower() for x in ['nhân công kỹ thuật', 'thợ', 'kỹ sư công nghệ']):
                    err = f"Description is Nhân công, NOT Máy thi công! Target: {t_desc}"
            elif comp_type == 'VẬT LIỆU':
                if t_unit not in ['kg', 'm2', 'bộ', 'lít']:
                    err = f"Unit mismatch: comp={unit} vs target={t_unit}"
                elif any(x in t_desc.lower() for x in ['nhân công', 'máy thi công']):
                    err = f"Description not Vật liệu: {t_desc}"
            elif comp_type == 'DỤNG CỤ':
                if t_unit not in ['ca', 'bộ']:
                    err = f"Unit mismatch: comp={unit} vs target={t_unit}"
                    
            if err:
                mismatches.append({
                    'code': code,
                    'comp_code': comp_code,
                    'comp_type': comp_type,
                    'comp_name': comp_name,
                    'unit': unit,
                    'val_col': val_col,
                    'ref': ref,
                    't_desc': t_desc,
                    't_unit': t_unit,
                    't_val': t_val,
                    'error': err
                })
            else:
                verified.append({
                    'code': code,
                    'comp_code': comp_code,
                    'comp_type': comp_type,
                    'comp_name': comp_name,
                    'ref': ref,
                    't_desc': t_desc,
                    't_val': t_val
                })

print(f"\n--- VERIFICATION RESULT ---")
print(f"Total verified links: {len(verified)}")
print(f"Total mismatched links: {len(mismatches)}")

if mismatches:
    print("\n!!! MISMATCHED ITEMS FOUND !!!")
    for m in mismatches:
        print(f"Code: {m['code']} | Comp: {m['comp_name']} ({m['comp_type']})")
        print(f"  Col: {m['val_col']} | Ref: {m['ref']} | Target Unit: {m['t_unit']} | Target Val: {m['t_val']}")
        print(f"  Target Desc: {m['t_desc']}")
        print(f"  ERROR: {m['error']}\n")
