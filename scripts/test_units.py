import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook('C:/Users/Admin/Downloads/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx', data_only=True)
ws_nc = wb['Hao_Phi_Nhan_Cong']
ws_vl = wb['Hao_Phi_Vat_Lieu']
ws_dc = wb['Hao_Phi_Dung_Cu']
ws_tb = wb['Hao_Phi_Thiet_Bi']

# Map each code to its data
nc_dict = {}
for r in range(2, ws_nc.max_row+1):
    code = ws_nc.cell(r, 1).value
    if code:
        nc_dict[code] = {
            'name': ws_nc.cell(r, 2).value,
            'type': ws_nc.cell(r, 3).value,
            'dinh_bien': ws_nc.cell(r, 4).value,
            'cong': ws_nc.cell(r, 5).value,
            'standard': ws_nc.cell(r, 6).value,
            'ke_thua_dc': ws_nc.cell(r, 7).value,
            'ke_thua_tb': ws_nc.cell(r, 8).value,
            'ke_thua_vl': ws_nc.cell(r, 9).value,
        }

print(f"Loaded {len(nc_dict)} parameters from NC.")

# Check unit for parameter:
# Where did '0C' and '%' come from in the image?
# In image:
# KK1a: Nhiệt độ -> Đơn vị: 0C
# KK1b: Độ ẩm -> Đơn vị: %
# Is there a sheet or table with units?
print("DM_Hoc_Phan:")
ws_hp = wb['DM_Hoc_Phan']
for r in range(1, ws_hp.max_row+1):
    print([ws_hp.cell(r, c).value for c in range(1, ws_hp.max_column+1)])

print("\nQuy_Trinh_Text_Nho:")
ws_qt = wb['Quy_Trinh_Text_Nho']
for r in range(1, ws_qt.max_row+1):
    print([ws_qt.cell(r, c).value for c in range(1, ws_qt.max_column+1)])
