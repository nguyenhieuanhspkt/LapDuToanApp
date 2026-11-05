import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb_path = 'd:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
wb = openpyxl.load_workbook(wb_path)

ws_02 = wb['02_Chi_Tiet_856_Dich_Vu']
print(f"Sheet 02 has {ws_02.max_row} rows.")

# Inspect first 5 tasks from sheet 02
for r in range(5, 10):
    stt = ws_02.cell(r, 1).value
    khu_vuc = ws_02.cell(r, 2).value
    ten_ct = ws_02.cell(r, 3).value
    q = ws_02.cell(r, 4).value
    dvt = ws_02.cell(r, 5).value
    ma_dm = ws_02.cell(r, 9).value
    print(f"STT {stt}: Q={q} {dvt} | Code={ma_dm} | Name={ten_ct[:40]}")
