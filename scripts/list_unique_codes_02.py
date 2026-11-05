import openpyxl

wb = openpyxl.load_workbook('BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=True)
ws_02 = wb['02_Chi_Tiet_856_Dich_Vu']

codes = set()
for r in range(4, ws_02.max_row + 1):
    c = ws_02.cell(r, 3).value
    if c:
        codes.add(str(c).strip())

print(f"Total unique codes in 02_Chi_Tiet_856_Dich_Vu: {len(codes)}")
for c in sorted(codes):
    print(f"  - {c}")
