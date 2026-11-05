import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
wb = openpyxl.load_workbook('d:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=True)
ws = wb['Định mức áp dụng']

print("=== CELL LOCATIONS IN 'Định mức áp dụng' ===")
for r in range(45, ws.max_row + 1):
    val_a = ws.cell(r, 1).value
    val_b = str(ws.cell(r, 2).value)
    val_c = ws.cell(r, 3).value
    val_d = ws.cell(r, 4).value
    
    # Check headers and items
    if val_d is not None and val_d != 'Mức hao phí\nđịnh mức':
        print(f"Row {r:3d}: col D = {val_d:<6} | col C = {str(val_c):<8} | {val_b[:55]}")
