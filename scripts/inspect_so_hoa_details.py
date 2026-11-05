import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook('C:/Users/Admin/Downloads/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx', data_only=True)
ws_nc = wb['Hao_Phi_Nhan_Cong']
ws_vl = wb['Hao_Phi_Vat_Lieu']
ws_dc = wb['Hao_Phi_Dung_Cu']
ws_tb = wb['Hao_Phi_Thiet_Bi']

print("VL unique codes:", len(set(ws_vl.cell(r, 1).value for r in range(2, ws_vl.max_row+1) if ws_vl.cell(r, 1).value)))
print("DC unique codes:", len(set(ws_dc.cell(r, 1).value for r in range(2, ws_dc.max_row+1) if ws_dc.cell(r, 1).value)))
print("TB unique codes:", len(set(ws_tb.cell(r, 1).value for r in range(2, ws_tb.max_row+1) if ws_tb.cell(r, 1).value)))

# Let's inspect 1KK1a and 1KK1b in VL
print("\n--- VL for 1KK1a ---")
for r in range(2, ws_vl.max_row+1):
    if ws_vl.cell(r, 1).value == '1KK1a':
        print(f"Name: {ws_vl.cell(r, 3).value}, Unit: {ws_vl.cell(r, 4).value}, Rate: {ws_vl.cell(r, 5).value}")

print("\n--- DC for 1KK1a ---")
for r in range(2, ws_dc.max_row+1):
    if ws_dc.cell(r, 1).value == '1KK1a':
        print(f"Name: {ws_dc.cell(r, 6).value}, Unit: {ws_dc.cell(r, 7).value}, Rate: {ws_dc.cell(r, 8).value}")

print("\n--- TB for 1KK1a ---")
for r in range(2, ws_tb.max_row+1):
    if ws_tb.cell(r, 1).value == '1KK1a':
        print(f"Name: {ws_tb.cell(r, 4).value}, Unit: {ws_tb.cell(r, 6).value}, Rate: {ws_tb.cell(r, 7).value}")
