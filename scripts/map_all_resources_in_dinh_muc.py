import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_dm = wb['Định mức áp dụng']

print("=== MAPPING ALL RESOURCES IN 'Định mức áp dụng' ===")

# Let's inspect rows 45 to 246
current_norm = None
resources = []

for r in range(45, 247):
    a = str(ws_dm.cell(r, 1).value or '').strip()
    b = str(ws_dm.cell(r, 2).value or '').strip()
    c = str(ws_dm.cell(r, 3).value or '').strip()
    d = ws_dm.cell(r, 4).value
    e = str(ws_dm.cell(r, 5).value or '').strip()
    
    # Check if this row is a Norm Header (e.g. II.1, II.2, III.1, etc.)
    if a.startswith('II.') or a.startswith('III.'):
        current_norm = (r, a, b)
        print(f"\n--- NORM HEADER: Row {r} | STT: {a} | Title: {b[:60]}")
    elif b.startswith('•') or b.startswith('-') or b.startswith('   •') or b.startswith('   -') or 'Nhân công' in b or 'Máy thi công' in b or 'Vật liệu' in b:
        # Resource line
        res_type = 'UNKNOWN'
        if 'nhân công' in b.lower() or 'thợ' in b.lower() or 'kỹ sư' in b.lower() or c == 'công':
            res_type = 'NHÂN CÔNG'
        elif 'máy' in b.lower() or 'thiết bị' in b.lower() or 'đồng hồ' in b.lower() or 'bơm' in b.lower() or 'bàn' in b.lower() or 'terrameter' in b.lower() or c == 'ca':
            res_type = 'MÁY THI CÔNG'
        elif 'sơn' in b.lower() or 'vật liệu' in b.lower() or 'nước cất' in b.lower() or c in ['kg', 'm2']:
            res_type = 'VẬT LIỆU'
        elif 'dụng cụ' in b.lower():
            res_type = 'DỤNG CỤ'
            
        print(f"  Res: Row {r} | Type: {res_type} | Desc: {b} | Unit: {c} | Value: {d}")
    elif a.startswith('Bước'):
        # Step row in 11 steps
        pass
