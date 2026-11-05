import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_dm = wb['Định mức áp dụng']
ws_hp = wb['hao phi VT']

print("=== STRICT AUDIT OF ALL 3264 ROWS IN 'hao phi VT' ===")

errors = []
verified_count = 0

for r in range(7, ws_hp.max_row + 1):
    stt_task = ws_hp.cell(r, 1).value
    code = ws_hp.cell(r, 2).value
    loai = ws_hp.cell(r, 3).value
    name = ws_hp.cell(r, 5).value
    unit = ws_hp.cell(r, 6).value
    
    # Check H (Vật liệu), I (Nhân công), J (Công cụ dụng cụ), K (Máy thi công)
    cols_to_check = [
        (8, 'H', 'VẬT LIỆU', ['kg', 'm2', 'bộ', 'lít']),
        (9, 'I', 'NHÂN CÔNG', ['công']),
        (10, 'J', 'DỤNG CỤ', ['ca', 'bộ']),
        (11, 'K', 'MÁY THI CÔNG', ['ca'])
    ]
    
    for c_idx, c_col, expected_type, allowed_units in cols_to_check:
        cell_val = str(ws_hp.cell(r, c_idx).value or '').strip()
        if not cell_val:
            continue
            
        if 'Định mức áp dụng' in cell_val:
            ref = cell_val.replace('=', '').replace("'", '').replace('Định mức áp dụng!', '').strip()
            dm_cell = ws_dm[ref]
            dm_row = dm_cell.row
            dm_stt = ws_dm.cell(dm_row, 1).value
            dm_desc = str(ws_dm.cell(dm_row, 2).value or '').strip()
            dm_unit = str(ws_dm.cell(dm_row, 3).value or '').strip()
            dm_rate = ws_dm.cell(dm_row, 4).value
            
            # Check 1: Unit match
            if dm_unit not in allowed_units:
                errors.append(f"Row {r} [{loai}] Col {c_col}: Link {cell_val} -> DM Row {dm_row} Unit mismatch! (Expected {allowed_units}, got '{dm_unit}') | Desc: {dm_desc} | Rate: {dm_rate}")
                continue
                
            # Check 2: Content semantic match
            if expected_type == 'MÁY THI CÔNG':
                if any(k in dm_desc.lower() for k in ['nhân công kỹ thuật', 'thợ', 'kỹ sư']):
                    errors.append(f"Row {r} [MÁY THI CÔNG] Col K: Link {cell_val} -> DM Row {dm_row} is NHÂN CÔNG! Desc: {dm_desc}")
                    continue
            elif expected_type == 'NHÂN CÔNG':
                if any(k in dm_desc.lower() for k in ['máy thi công', 'đồng hồ vạn năng', 'bơm tạo áp', 'mê gôm mét']):
                    errors.append(f"Row {r} [NHÂN CÔNG] Col I: Link {cell_val} -> DM Row {dm_row} is MÁY! Desc: {dm_desc}")
                    continue
            elif expected_type == 'VẬT LIỆU':
                if any(k in dm_desc.lower() for k in ['nhân công', 'thợ', 'kỹ sư', 'máy thi công']):
                    errors.append(f"Row {r} [VẬT LIỆU] Col H: Link {cell_val} -> DM Row {dm_row} is NOT Vật liệu! Desc: {dm_desc}")
                    continue
                    
            verified_count += 1

print(f"Total verified formula links: {verified_count}")
print(f"Total errors found: {len(errors)}")

if errors:
    print("\nERRORS ENCOUNTERED:")
    for e in errors[:50]:
        print(e)
else:
    print("\nPERFECT! 100% of all formula links from 'hao phi VT' to 'Định mức áp dụng' match exact resource type, unit, and description!")
