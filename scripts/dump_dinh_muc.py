import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_dm = wb['Định mức áp dụng']

print(f"Total rows in 'Định mức áp dụng': {ws_dm.max_row}")

with open('dinh_muc_rows_current.txt', 'w', encoding='utf-8') as f:
    for r in range(1, ws_dm.max_row + 1):
        a = ws_dm.cell(r, 1).value
        b = ws_dm.cell(r, 2).value
        c = ws_dm.cell(r, 3).value
        d = ws_dm.cell(r, 4).value
        e = ws_dm.cell(r, 5).value
        f_val = ws_dm.cell(r, 6).value
        
        line = f"Row {r:3d} | A: {str(a or ''):12s} | B: {str(b or ''):70s} | C: {str(c or ''):10s} | D: {str(d or ''):10s} | E: {str(e or ''):15s} | F: {str(f_val or '')}"
        f.write(line + '\n')

print("Wrote dinh_muc_rows_current.txt successfully.")
