import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

source_file = 'C:/Users/Admin/Downloads/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx'
target_pccc = 'd:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
target_pccc_dl = 'C:/Users/Admin/Downloads/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
target_so_hoa_dl = 'C:/Users/Admin/Downloads/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx'
target_so_hoa_local = 'd:/TaskApp_kiet/LapDuToanApp/So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx'

print("Loading data from So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx...")
wb_src = openpyxl.load_workbook(source_file, data_only=True)
ws_nc = wb_src['Hao_Phi_Nhan_Cong']
ws_vl = wb_src['Hao_Phi_Vat_Lieu']
ws_dc = wb_src['Hao_Phi_Dung_Cu']
ws_tb = wb_src['Hao_Phi_Thiet_Bi']

# Read parameters
# Unit map for parameters
unit_map = {
    'KK1a': '0C',
    'KK1b': '%',
    'KK2a': 'm/s',
    'KK2b': 'hướng',
    'KK3': 'hPa',
    'KK4a': 'mg/m3',
    'KK4b': 'mg/m3',
    'KK4c': 'mg/m3',
    'KK4d': 'mg/m3',
    'KK5a': 'mg/m3',
    'KK5b': 'mg/m3',
    'KK5c': 'mg/m3',
    'KK6': 'mg/m3',
    'KK7': 'mg/m3',
    'KK8': 'mg/m3',
    'KK9': 'mg/m3',
    'KK10': 'mg/m3',
    'KK11a': 'mg/m3',
    'KK11b': 'mg/m3',
    'KK11c': 'mg/m3',
    'KK11d': 'mg/m3',
    'KK11đ': 'mg/m3',
    'KK12a': 'mg/m3',
    'KK12b': 'mg/m3',
    'KK12c': 'mg/m3',
    'KK12d': 'mg/m3',
    'TO1a': 'dBA',
    'TO1b': 'dBA',
    'TO2': 'xe/h',
    'TO3a': 'dBA',
    'TO3b': 'dBA',
    'TO3c': 'dBA',
    'TO4': 'dBA',
    'ĐR01': 'dB'
}

# 1. Read NC
parameters = []
for r in range(2, ws_nc.max_row + 1):
    raw_code = ws_nc.cell(r, 1).value
    if not raw_code:
        continue
    # Clean code: 1KK1a -> KK1a, 2KK4a -> KK4a (or keep 2KK4a if noi nghiep)
    code_display = raw_code.replace('1', '', 1) if raw_code.startswith('1') else raw_code
    
    p = {
        'raw_code': raw_code,
        'code': code_display,
        'name': ws_nc.cell(r, 2).value,
        'stage': ws_nc.cell(r, 3).value,
        'nc_title': ws_nc.cell(r, 4).value,
        'nc_rate': ws_nc.cell(r, 5).value,
        'standard': ws_nc.cell(r, 6).value,
        'unit': unit_map.get(code_display, 'mẫu' if 'Nội nghiệp' in str(ws_nc.cell(r, 3).value) else 'điểm')
    }
    parameters.append(p)

print(f"Total parameters loaded: {len(parameters)}")

# 2. Read VL
vl_dict = {}
for r in range(2, ws_vl.max_row + 1):
    c = ws_vl.cell(r, 1).value
    if not c: continue
    vl_dict.setdefault(c, []).append({
        'name': ws_vl.cell(r, 3).value,
        'unit': ws_vl.cell(r, 4).value,
        'rate': ws_vl.cell(r, 5).value,
        'note': ws_vl.cell(r, 6).value
    })

# 3. Read DC
dc_dict = {}
for r in range(2, ws_dc.max_row + 1):
    c = ws_dc.cell(r, 1).value
    if not c: continue
    dc_dict.setdefault(c, []).append({
        'name': ws_dc.cell(r, 6).value,
        'unit': 'ca', # In template image unit is 'ca'
        'rate': ws_dc.cell(r, 8).value,
        'note': ws_dc.cell(r, 9).value
    })

# 4. Read TB
tb_dict = {}
for r in range(2, ws_tb.max_row + 1):
    c = ws_tb.cell(r, 1).value
    if not c: continue
    tb_dict.setdefault(c, []).append({
        'name': ws_tb.cell(r, 4).value,
        'unit': 'ca', # In template image unit is 'ca'
        'rate': ws_tb.cell(r, 7).value,
        'note': ws_tb.cell(r, 8).value
    })

def populate_hao_phi_vt_sheet(ws):
    ws.views.sheetView[0].showGridLines = True
    
    # Fonts
    font_tbl_hdr = Font(name='Times New Roman', size=9.5, bold=True, color='000000')
    font_cong_tac = Font(name='Times New Roman', size=10, bold=True, color='002060')
    font_item = Font(name='Times New Roman', size=9.5, color='000000')
    
    # Fills
    fill_tbl_hdr = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_cong_tac = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    # Borders
    border_thin = Side(border_style='thin', color='B0B0B0')
    box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')
    
    # Title / Top space (Row 1-4)
    ws.row_dimensions[1].height = 15
    ws.row_dimensions[2].height = 15
    ws.row_dimensions[3].height = 15
    ws.row_dimensions[4].height = 15
    
    # Row 5-6: Headers matching user screenshot exactly
    # Merges:
    ws.merge_cells('A5:A6')
    ws.merge_cells('B5:B6')
    ws.merge_cells('C5:C6')
    ws.merge_cells('D5:D6')
    ws.merge_cells('E5:E6')
    ws.merge_cells('F5:F6')
    ws.merge_cells('G5:G6')
    ws.merge_cells('H5:K5') # MỨC HAO PHÍ
    ws.merge_cells('L5:L6')
    ws.merge_cells('M5:P5') # KHỐI LƯỢNG HAO PHÍ
    ws.merge_cells('Q5:Q6')
    
    ws['A5'] = 'STT'
    ws['B5'] = 'Mã hiệu TT20'
    ws['C5'] = 'Phân loại'
    ws['D5'] = 'Giai đoạn'
    ws['E5'] = 'Tên công tác / Thành phần hao phí chi tiết'
    ws['F5'] = 'Đơn vị'
    ws['G5'] = 'Khối lượng công trình (Q)'
    
    ws['H5'] = 'MỨC HAO PHÍ'
    ws['H6'] = 'Vật liệu'
    ws['I6'] = 'Nhân công'
    ws['J6'] = 'Công cụ dụng\ncụ'
    ws['K6'] = 'Máy thi công'
    
    ws['L5'] = 'Hệ số'
    
    ws['M5'] = 'KHỐI LƯỢNG HAO PHÍ'
    ws['M6'] = 'Vật liệu'
    ws['N6'] = 'Nhân công'
    ws['O6'] = 'Công cụ dụng cụ'
    ws['P6'] = 'Máy thi công'
    
    ws['Q5'] = 'Căn cứ pháp lý định mức TT20'
    
    ws.row_dimensions[5].height = 24
    ws.row_dimensions[6].height = 24
    
    for r in [5, 6]:
        for c in range(1, 18):
            cell = ws.cell(r, c)
            cell.font = font_tbl_hdr
            cell.fill = fill_tbl_hdr
            cell.border = box_border
            cell.alignment = align_center

    curr_r = 7
    stt_idx = 1
    
    for p in parameters:
        q_row = curr_r
        raw_code = p['raw_code']
        code_disp = p['code']
        stage = p['stage']
        
        # 1. DÒNG CÔNG TÁC
        ws.row_dimensions[curr_r].height = 20
        ws.cell(row=curr_r, column=1, value=stt_idx).alignment = align_center
        ws.cell(row=curr_r, column=2, value=code_disp).alignment = align_center
        ws.cell(row=curr_r, column=3, value='CÔNG TÁC').alignment = align_center
        ws.cell(row=curr_r, column=4, value='').alignment = align_center
        ws.cell(row=curr_r, column=5, value=p['name']).alignment = align_left
        ws.cell(row=curr_r, column=6, value=p['unit']).alignment = align_center
        ws.cell(row=curr_r, column=7, value=10).alignment = align_center # Default Q=10
        ws.cell(row=curr_r, column=17, value='Thông tư 20/2017/TT-BTNMT').alignment = align_left
        
        for c in range(1, 18):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = font_cong_tac
            cell.border = box_border
            
        curr_r += 1
        
        # 2. DÒNG NHÂN CÔNG
        if p['nc_rate'] is not None:
            ws.row_dimensions[curr_r].height = 19
            ws.cell(row=curr_r, column=1, value='').alignment = align_center
            ws.cell(row=curr_r, column=2, value='NC_LDKT').alignment = align_center
            ws.cell(row=curr_r, column=3, value='NHÂN CÔNG').alignment = align_center
            ws.cell(row=curr_r, column=4, value=stage).alignment = align_left
            ws.cell(row=curr_r, column=5, value=f"- {p['nc_title']}").alignment = align_left
            ws.cell(row=curr_r, column=6, value='công').alignment = align_center
            
            c_rate = ws.cell(row=curr_r, column=9, value=float(p['nc_rate']))
            c_rate.alignment = align_right
            c_rate.number_format = '0.0000'
            
            c_qty = ws.cell(row=curr_r, column=14, value=f"=G${q_row}*I{curr_r}")
            c_qty.alignment = align_right
            c_qty.number_format = '0.0000'
            
            ws.cell(row=curr_r, column=17, value='Bảng 1 TT20 & TTLT 56/2015').alignment = align_left
            
            for c in range(1, 18):
                cell = ws.cell(row=curr_r, column=c)
                cell.font = font_item
                cell.border = box_border
            curr_r += 1
            
        # 3. CÁC DÒNG VẬT LIỆU
        vl_items = vl_dict.get(raw_code, [])
        for vl in vl_items:
            ws.row_dimensions[curr_r].height = 19
            ws.cell(row=curr_r, column=1, value='').alignment = align_center
            ws.cell(row=curr_r, column=2, value=f"VL_{code_disp}").alignment = align_center
            ws.cell(row=curr_r, column=3, value='VẬT LIỆU').alignment = align_center
            ws.cell(row=curr_r, column=4, value=stage).alignment = align_left
            ws.cell(row=curr_r, column=5, value=f"- {vl['name']}").alignment = align_left
            ws.cell(row=curr_r, column=6, value=vl['unit']).alignment = align_center
            
            r_val = float(vl['rate']) if vl['rate'] is not None else 0.0
            c_rate = ws.cell(row=curr_r, column=8, value=r_val)
            c_rate.alignment = align_right
            c_rate.number_format = '0.0000'
            
            c_qty = ws.cell(row=curr_r, column=13, value=f"=G${q_row}*H{curr_r}")
            c_qty.alignment = align_right
            c_qty.number_format = '0.0000'
            
            ws.cell(row=curr_r, column=17, value='Bảng 5 TT20/2017').alignment = align_left
            
            for c in range(1, 18):
                cell = ws.cell(row=curr_r, column=c)
                cell.font = font_item
                cell.border = box_border
            curr_r += 1
            
        # 4. CÁC DÒNG DỤNG CỤ
        dc_items = dc_dict.get(raw_code, [])
        for dc in dc_items:
            ws.row_dimensions[curr_r].height = 19
            ws.cell(row=curr_r, column=1, value='').alignment = align_center
            ws.cell(row=curr_r, column=2, value=f"DC_{code_disp}").alignment = align_center
            ws.cell(row=curr_r, column=3, value='DỤNG CỤ').alignment = align_center
            ws.cell(row=curr_r, column=4, value=stage).alignment = align_left
            ws.cell(row=curr_r, column=5, value=f"- {dc['name']}").alignment = align_left
            ws.cell(row=curr_r, column=6, value=dc['unit']).alignment = align_center
            
            r_val = float(dc['rate']) if dc['rate'] is not None else 0.0
            c_rate = ws.cell(row=curr_r, column=10, value=r_val)
            c_rate.alignment = align_right
            c_rate.number_format = '0.0000'
            
            c_qty = ws.cell(row=curr_r, column=15, value=f"=G${q_row}*J{curr_r}")
            c_qty.alignment = align_right
            c_qty.number_format = '0.0000'
            
            ws.cell(row=curr_r, column=17, value='Bảng 3 TT20/2017').alignment = align_left
            
            for c in range(1, 18):
                cell = ws.cell(row=curr_r, column=c)
                cell.font = font_item
                cell.border = box_border
            curr_r += 1
            
        # 5. CÁC DÒNG THIẾT BỊ / MÁY
        tb_items = tb_dict.get(raw_code, [])
        for tb in tb_items:
            ws.row_dimensions[curr_r].height = 19
            ws.cell(row=curr_r, column=1, value='').alignment = align_center
            ws.cell(row=curr_r, column=2, value=f"M_{code_disp}").alignment = align_center
            ws.cell(row=curr_r, column=3, value='THIẾT BỊ').alignment = align_center
            ws.cell(row=curr_r, column=4, value=stage).alignment = align_left
            ws.cell(row=curr_r, column=5, value=f"- {tb['name']}").alignment = align_left
            ws.cell(row=curr_r, column=6, value=tb['unit']).alignment = align_center
            
            r_val = float(tb['rate']) if tb['rate'] is not None else 0.0
            c_rate = ws.cell(row=curr_r, column=11, value=r_val)
            c_rate.alignment = align_right
            c_rate.number_format = '0.0000'
            
            c_qty = ws.cell(row=curr_r, column=16, value=f"=G${q_row}*K{curr_r}")
            c_qty.alignment = align_right
            c_qty.number_format = '0.0000'
            
            ws.cell(row=curr_r, column=17, value='Bảng 4 TT20 & TT 45/2013').alignment = align_left
            
            for c in range(1, 18):
                cell = ws.cell(row=curr_r, column=c)
                cell.font = font_item
                cell.border = box_border
            curr_r += 1
            
        stt_idx += 1

    # Set Column widths
    ws.column_dimensions['A'].width = 7
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 13
    ws.column_dimensions['E'].width = 38
    ws.column_dimensions['F'].width = 9
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 11
    ws.column_dimensions['I'].width = 11
    ws.column_dimensions['J'].width = 13
    ws.column_dimensions['K'].width = 12
    ws.column_dimensions['L'].width = 8
    ws.column_dimensions['M'].width = 13
    ws.column_dimensions['N'].width = 13
    ws.column_dimensions['O'].width = 15
    ws.column_dimensions['P'].width = 13
    ws.column_dimensions['Q'].width = 28
    
    ws.freeze_panes = 'F7'
    return curr_r

# Process 1: Update So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx
print("\n--- Updating So_Hoa_Phan_II_1_Khong_Khi_Tieng_On.xlsx ---")
wb_so_hoa = openpyxl.load_workbook(source_file)
if 'hao phi VT' in wb_so_hoa.sheetnames:
    wb_so_hoa.remove(wb_so_hoa['hao phi VT'])
ws_sh = wb_so_hoa.create_sheet('hao phi VT')
max_r_sh = populate_hao_phi_vt_sheet(ws_sh)
print(f"Sheet 'hao phi VT' created in So_Hoa with {max_r_sh} rows.")

try:
    wb_so_hoa.save(target_so_hoa_dl)
    print(f"Saved to {target_so_hoa_dl}")
except Exception as e:
    print(f"Could not save to {target_so_hoa_dl}: {e}")

try:
    wb_so_hoa.save(target_so_hoa_local)
    print(f"Saved to {target_so_hoa_local}")
except Exception as e:
    print(f"Could not save to {target_so_hoa_local}: {e}")

# Process 2: Also add sheet 'hao phi VT' to BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx
print("\n--- Adding 'hao phi VT' to BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx ---")
wb_pccc = openpyxl.load_workbook(target_pccc)
if 'hao phi VT' in wb_pccc.sheetnames:
    wb_pccc.remove(wb_pccc['hao phi VT'])
ws_pccc = wb_pccc.create_sheet('hao phi VT')
max_r_pccc = populate_hao_phi_vt_sheet(ws_pccc)
print(f"Sheet 'hao phi VT' added to PCCC with {max_r_pccc} rows.")

wb_pccc.save(target_pccc)
print(f"Saved to local PCCC: {target_pccc}")

try:
    wb_pccc.save(target_pccc_dl)
    print(f"Saved to Downloads PCCC: {target_pccc_dl}")
except Exception as e:
    print(f"Could not save to Downloads PCCC: {e}")

print("\nDone all!")
