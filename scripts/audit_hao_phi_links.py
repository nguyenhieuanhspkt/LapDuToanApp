import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_dm = wb['Định mức áp dụng']
ws_hp = wb['hao phi VT']

print('Checking all links in hao phi VT to Định mức áp dụng...')
all_links = []
errors = []

for r in range(7, ws_hp.max_row + 1):
    stt_task = ws_hp.cell(r, 1).value
    code = ws_hp.cell(r, 2).value
    loai = ws_hp.cell(r, 3).value # CÔNG TÁC, NHÂN CÔNG, MÁY THI CÔNG, VẬT LIỆU, DỤNG CỤ
    name = ws_hp.cell(r, 5).value
    
    # check columns H(8), I(9), J(10), K(11)
    for col_idx, col_name, expected_types in [
        (8, 'H (Vật liệu)', ['VẬT LIỆU']),
        (9, 'I (Nhân công)', ['NHÂN CÔNG']),
        (10, 'J (Công cụ dụng cụ)', ['DỤNG CỤ']),
        (11, 'K (Máy thi công)', ['MÁY THI CÔNG', 'THIẾT BỊ'])
    ]:
        cell_val = str(ws_hp.cell(r, col_idx).value or '')
        if 'Định mức áp dụng' in cell_val:
            ref = cell_val.replace('=', '').replace("'", '').replace('Định mức áp dụng!', '').strip()
            target_cell = ws_dm[ref]
            target_row = target_cell.row
            target_stt = ws_dm.cell(target_row, 1).value
            target_desc = str(ws_dm.cell(target_row, 2).value or '')
            target_unit = str(ws_dm.cell(target_row, 3).value or '')
            target_val = ws_dm.cell(target_row, 4).value
            
            all_links.append((r, loai, name, col_name, cell_val, target_row, target_desc, target_unit, target_val))
            
            # Check mismatch:
            # If col_name is K (Máy thi công), target must be Máy thi công / Thiết bị, NOT Nhân công!
            td_lower = target_desc.lower()
            if 'máy' in col_name.lower():
                # If target describes nhân công or công or thợ, or doesn't mention máy / ca / thiết bị
                if any(x in td_lower for x in ['nhân công', 'kỹ sư', 'thợ', 'bậc']) or target_unit.lower() == 'công':
                    errors.append((r, loai, name, col_name, cell_val, target_row, target_desc, target_unit, target_val))
            elif 'nhân công' in col_name.lower():
                if any(x in td_lower for x in ['máy thi công', 'đồng hồ', 'thiết bị', 'ca']) or target_unit.lower() == 'ca':
                    errors.append((r, loai, name, col_name, cell_val, target_row, target_desc, target_unit, target_val))
            elif 'vật liệu' in col_name.lower():
                if any(x in td_lower for x in ['nhân công', 'thợ', 'kỹ sư', 'máy thi công']) or target_unit.lower() in ['công', 'ca']:
                    errors.append((r, loai, name, col_name, cell_val, target_row, target_desc, target_unit, target_val))

print(f'Total links checked: {len(all_links)}')
print(f'Total mismatches found: {len(errors)}')

for e in errors[:50]:
    r, loai, name, col_name, cell_val, target_row, target_desc, target_unit, target_val = e
    print(f'Row {r} [{loai}] ({name[:35]}...) | Col {col_name}: {cell_val} -> DinhMuc Row {target_row}: {target_desc} ({target_unit}) = {target_val}')
