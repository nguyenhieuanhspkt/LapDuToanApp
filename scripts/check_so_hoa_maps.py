import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook('C:/Users/Admin/Downloads/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx', data_only=True)
ws_nc = wb['Hao_Phi_Nhan_Cong']
ws_vl = wb['Hao_Phi_Vat_Lieu']
ws_dc = wb['Hao_Phi_Dung_Cu']
ws_tb = wb['Hao_Phi_Thiet_Bi']

# 1. Read VL items per code
vl_map = {}
for r in range(2, ws_vl.max_row + 1):
    c = ws_vl.cell(r, 1).value
    if not c: continue
    c_clean = c.replace('1', '', 1) if c.startswith('1') else c
    item = {
        'name': ws_vl.cell(r, 3).value,
        'unit': ws_vl.cell(r, 4).value,
        'rate': ws_vl.cell(r, 5).value,
        'note': ws_vl.cell(r, 6).value
    }
    vl_map.setdefault(c, []).append(item)
    vl_map.setdefault(c_clean, []).append(item)

# 2. Read DC items per code
dc_map = {}
for r in range(2, ws_dc.max_row + 1):
    c = ws_dc.cell(r, 1).value
    if not c: continue
    c_clean = c.replace('1', '', 1) if c.startswith('1') else c
    item = {
        'name': ws_dc.cell(r, 6).value, # In DC: col 6 is 'Tên dụng cụ' or col 4? Let's check!
        'unit': ws_dc.cell(r, 7).value,
        'rate': ws_dc.cell(r, 8).value,
        'note': ws_dc.cell(r, 9).value
    }
    dc_map.setdefault(c, []).append(item)
    dc_map.setdefault(c_clean, []).append(item)

# 3. Read TB items per code
tb_map = {}
for r in range(2, ws_tb.max_row + 1):
    c = ws_tb.cell(r, 1).value
    if not c: continue
    c_clean = c.replace('1', '', 1) if c.startswith('1') else c
    item = {
        'name': ws_tb.cell(r, 4).value,
        'unit': ws_tb.cell(r, 6).value,
        'rate': ws_tb.cell(r, 7).value,
        'note': ws_tb.cell(r, 8).value
    }
    tb_map.setdefault(c, []).append(item)
    tb_map.setdefault(c_clean, []).append(item)

print(f"VL map keys: {len(vl_map)}, DC map keys: {len(dc_map)}, TB map keys: {len(tb_map)}")
print("Sample DC 1KK1a:", dc_map.get('1KK1a'))
print("Sample TB 1KK1a:", tb_map.get('1KK1a'))
