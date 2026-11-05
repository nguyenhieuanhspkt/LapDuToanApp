import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
wb = openpyxl.load_workbook('d:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx', data_only=True)
ws = wb['Định mức áp dụng']

print(f"Total rows in Định mức áp dụng: {ws.max_row}")
for r in range(45, ws.max_row + 1):
    val_a = str(ws.cell(r, 1).value)
    val_b = str(ws.cell(r, 2).value)
    val_c = str(ws.cell(r, 3).value)
    val_d = str(ws.cell(r, 4).value)
    if val_a.startswith('II.') or val_a.startswith('III.') or val_a.startswith('IV.') or '•' in val_b or 'Bước' in val_a:
        print(f"R{r:3d}: colA={val_a:10s} | colD={val_d:8s} | colC={val_c:8s} | {val_b[:55]}")
