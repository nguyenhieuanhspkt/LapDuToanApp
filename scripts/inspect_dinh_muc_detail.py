import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=False)
ws_dm = wb['Định mức áp dụng']

print("Analyzing all rows in 'Định mức áp dụng':")
for r in range(45, ws_dm.max_row + 1):
    a = ws_dm.cell(r, 1).value
    b = ws_dm.cell(r, 2).value
    c = ws_dm.cell(r, 3).value
    d = ws_dm.cell(r, 4).value
    e = ws_dm.cell(r, 5).value
    f = ws_dm.cell(r, 6).value
    
    # print if it's a code or resource
    if a and any(a.startswith(x) for x in ['II.', 'III.', 'IV.', 'Bước', 'Tổng']):
        print(f"Row {r:3d} | A={str(a):10s} | B={str(b)[:60]:60s} | C={str(c):8s} | D={str(d):8s} | E={str(e):10s}")
    elif b and ('•' in str(b) or '-' in str(b) or 'Nhân công' in str(b) or 'Máy thi công' in str(b) or 'Vật liệu' in str(b)):
        print(f"Row {r:3d} | A={str(a):10s} | B={str(b)[:60]:60s} | C={str(c):8s} | D={str(d):8s} | E={str(e):10s}")
