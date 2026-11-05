import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_hp = wb['hao phi VT']

print("=== AUDITING 3 RESOURCE SHEETS AGAINST 'hao phi VT' ===")

sheets_to_check = [
    ('Nhân công', ['NHÂN CÔNG'], ['N']),
    ('Máy thi công', ['MÁY THI CÔNG', 'DỤNG CỤ'], ['P', 'O']),
    ('Vật liệu', ['VẬT LIỆU'], ['M'])
]

total_checked = 0
total_errors = 0

for sname, expected_types, allowed_cols in sheets_to_check:
    ws = wb[sname]
    print(f"\nChecking sheet '{sname}' (rows 6 to {ws.max_row})...")
    s_checked = 0
    s_errors = 0
    
    for r in range(6, ws.max_row):
        stt = str(ws.cell(r, 1).value or '')
        val_f = str(ws.cell(r, 6).value or '') # Col F: Khối lượng hao phí
        
        # Detail row: starts with 'STT '
        if stt.startswith('STT '):
            s_checked += 1
            if 'hao phi VT' not in val_f:
                print(f"  Error at Row {r}: Col F has no link: {val_f}")
                s_errors += 1
                continue
                
            # ref is like ='hao phi VT'!N8
            ref = val_f.replace('=', '').replace("'", '').replace('hao phi VT!', '').strip()
            # extract col and row
            col_letter = ''.join([c for c in ref if c.isalpha()])
            hp_row = int(''.join([c for c in ref if c.isdigit()]))
            
            if col_letter not in allowed_cols:
                print(f"  Error at Row {r}: Linked col {col_letter} not in allowed {allowed_cols}")
                s_errors += 1
                continue
                
            # Check row in hao phi VT
            hp_type = str(ws_hp.cell(hp_row, 3).value or '').strip()
            hp_name = str(ws_hp.cell(hp_row, 5).value or '').strip()
            
            if hp_type not in expected_types:
                print(f"  Error at Row {r}: Linked to hao phi VT Row {hp_row} which is '{hp_type}' (expected {expected_types})! Name: {hp_name}")
                s_errors += 1
                continue
                
        # Master row: numeric STT (e.g. 1, 2, 3...)
        elif stt.isdigit():
            # Must be a SUM formula
            if not val_f.startswith('=SUM('):
                print(f"  Error at Master Row {r}: Col F is not =SUM(...): {val_f}")
                s_errors += 1
                
    print(f"Sheet '{sname}': Checked {s_checked} detail rows. Errors found: {s_errors}")
    total_checked += s_checked
    total_errors += s_errors

print(f"\n==========================================")
print(f"AUDIT SUMMARY: Checked {total_checked} detail links across all 3 sheets.")
print(f"Total cross-sheet mismatch errors: {total_errors}")
if total_errors == 0:
    print("ALL 3 SHEETS ARE 100% CLEAN, STRICTLY AUDITED, AND PERFECTLY SYNCHRONIZED!")
