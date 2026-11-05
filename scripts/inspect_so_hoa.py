import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook('C:/Users/Admin/Downloads/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx', data_only=True)
ws_nc = wb['Hao_Phi_Nhan_Cong']
ws_vl = wb['Hao_Phi_Vat_Lieu']
ws_dc = wb['Hao_Phi_Dung_Cu']
ws_tb = wb['Hao_Phi_Thiet_Bi']

print("=== NC HEADERS ===")
print([ws_nc.cell(1, c).value for c in range(1, ws_nc.max_column+1)])
print("=== VL HEADERS ===")
print([ws_vl.cell(1, c).value for c in range(1, ws_vl.max_column+1)])
print("=== DC HEADERS ===")
print([ws_dc.cell(1, c).value for c in range(1, ws_dc.max_column+1)])
print("=== TB HEADERS ===")
print([ws_tb.cell(1, c).value for c in range(1, ws_tb.max_column+1)])

print("\n=== SAMPLE PARAMETERS IN NC ===")
for r in range(2, ws_nc.max_row+1):
    val = ws_nc.cell(r, 1).value
    if val:
        print(f"Row {r}: Code={val}, Name={ws_nc.cell(r, 2).value}, Type={ws_nc.cell(r, 3).value}, DinhBien={ws_nc.cell(r, 4).value}, Cong={ws_nc.cell(r, 5).value}")
