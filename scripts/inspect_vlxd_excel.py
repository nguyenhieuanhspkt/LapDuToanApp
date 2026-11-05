import openpyxl

file_path = r'D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng\Vật liệu\Gia_VLXD_LamDong_T08_2026.xlsx'

wb = openpyxl.load_workbook(file_path, data_only=True)
print(f"Sheets in {file_path}:")
for s in wb.sheetnames:
    ws = wb[s]
    print(f"\n--- SHEET: {s} ({ws.max_row} rows, {ws.max_column} cols) ---")
    for r in range(1, min(12, ws.max_row + 1)):
        vals = [str(ws.cell(r, c).value or '') for c in range(1, min(10, ws.max_column + 1))]
        if any(vals):
            print(f"  Row {r:2d}: " + ' | '.join(vals))
