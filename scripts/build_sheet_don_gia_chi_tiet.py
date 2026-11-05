# -*- coding: utf-8 -*-
"""
Script build_sheet_don_gia_chi_tiet.py
Xây dựng sheet "Đơn giá chi tiết" theo chuẩn Dự toán F1 / Thông tư 36/2026/TT-BXD:
1. Định mức: link đến sheet 'hao phi VT' (Mức hao phí định mức * Hệ số).
2. Đơn giá: link đến sheet 'Vật liệu', 'Nhân công', 'Máy thi công' tương ứng.
3. Hệ số: bằng 1.
4. Thành tiền từng thành phần: subtotal (Cộng Vật liệu, Cộng Nhân công, Cộng Máy thi công)
   và total thành 1 đơn giá trực tiếp cho từng công tác.
5. Cập nhật ngược lại các cột Đơn giá trực tiếp (VL, NC, Máy) trong sheet 'Công trình'.
"""

import os
import sys
import io
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PROJECT_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
DOWNLOADS_PATH = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
ALT_DOWNLOADS = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"

print(">>> [1/6] Đang mở workbook và phân tích dữ liệu...")
wb = openpyxl.load_workbook(PROJECT_PATH)

ws_hp = wb['hao phi VT']
ws_nc = wb['Nhân công']
ws_mc = wb['Máy thi công']
ws_vl = wb['Vật liệu']
ws_ct = wb['Công trình']

# 1. Thu thập tọa độ Master rows trong Nhân công
nc_masters = {}
for r_idx, row in enumerate(ws_nc.iter_rows(min_row=6, values_only=True), start=6):
    c1, c2, c3, c6 = str(row[0] or ''), str(row[1] or ''), str(row[2] or ''), str(row[5] or '')
    if c6.startswith('=SUM'):
        is_pt = 'PHƯỚC THÁI' in c3.upper()
        nc_masters[(c2, is_pt)] = r_idx

# 2. Thu thập tọa độ Master rows trong Máy thi công
mc_masters = {}
for r_idx, row in enumerate(ws_mc.iter_rows(min_row=6, values_only=True), start=6):
    c1, c2, c3, c6 = str(row[0] or ''), str(row[1] or ''), str(row[2] or ''), str(row[5] or '')
    if c6.startswith('=SUM'):
        is_pt = 'PHƯỚC THÁI' in c3.upper()
        mc_masters[(c2, is_pt)] = r_idx

# 3. Thu thập tọa độ Master rows trong Vật liệu
vl_masters = {}
for r_idx, row in enumerate(ws_vl.iter_rows(min_row=6, values_only=True), start=6):
    c1, c2, c6 = str(row[0] or ''), str(row[1] or ''), str(row[5] or '')
    if c1.isdigit() or c6.startswith('=SUM'):
        vl_masters[c2] = r_idx

def get_nc_master_cell(code, is_pt):
    if (code, is_pt) in nc_masters:
        return f"'Nhân công'!$G${nc_masters[(code, is_pt)]}"
    if (code, False) in nc_masters:
        return f"'Nhân công'!$G${nc_masters[(code, False)]}"
    if 'DB.24001' in code:
        return "'Nhân công'!$G$615"
    if any(k in code for k in ['1.23', '1.31', '2.04', '1.18', '18405', '1.13', '41171']):
        return "'Nhân công'!$G$558"
    return f"'Nhân công'!$G$45" if is_pt else "'Nhân công'!$G$6"

def get_mc_master_cell(code, is_pt):
    if (code, is_pt) in mc_masters:
        return f"'Máy thi công'!$G${mc_masters[(code, is_pt)]}"
    if (code, False) in mc_masters:
        return f"'Máy thi công'!$G${mc_masters[(code, False)]}"
    if 'VOM' in code:
        return "'Máy thi công'!$G$41" if is_pt else "'Máy thi công'!$G$6"
    if 'BUI' in code:
        return "'Máy thi công'!$G$77" if is_pt else "'Máy thi công'!$G$46"
    if 'SOLO' in code:
        return "'Máy thi công'!$G$130" if is_pt else "'Máy thi công'!$G$82"
    if 'MEG' in code:
        return "'Máy thi công'!$G$483" if is_pt else "'Máy thi công'!$G$408"
    return "'Máy thi công'!$G$6"

def get_vl_master_cell(code):
    if code in vl_masters:
        return f"'Vật liệu'!$G${vl_masters[code]}"
    if '1.13' in code:
        return "'Vật liệu'!$G$6"
    if 'LOT' in code:
        return "'Vật liệu'!$G$12"
    return "'Vật liệu'!$G$70"

# 4. Phân tích các khối công tác trong 'hao phi VT'
tasks_data = [] # mỗi phần tử: {'stt': int, 'r_ct': int, 'r_hp': int, 'vl': [], 'nc': [], 'mc': []}
current_task = None
stt_counter = 0

for r_idx, row in enumerate(ws_hp.iter_rows(min_row=7, values_only=True), start=7):
    c1 = row[0]
    c2 = str(row[1] or '').strip()
    c3 = str(row[2] or '').strip().upper()
    c4 = str(row[3] or '').strip()
    c5 = str(row[4] or '').strip()
    c6 = str(row[5] or '').strip()
    c8 = row[7] # Mức hao phí VL
    c9 = row[8] # Mức hao phí NC
    c11 = row[10] # Mức hao phí May
    c12 = row[11] # Hệ số
    
    if c3 == 'CÔNG TÁC':
        stt_counter += 1
        current_task = {
            'stt': stt_counter,
            'r_ct': stt_counter + 6,
            'r_hp': r_idx,
            'is_pt': (681 <= stt_counter <= 776),
            'vl': [],
            'nc': [],
            'mc': []
        }
        tasks_data.append(current_task)
    elif current_task is not None:
        if 'VẬT LIỆU' in c3:
            current_task['vl'].append({
                'r_hp': r_idx,
                'code': c2,
                'name': c5,
                'dvt': c6,
                'col_norm': 'H'
            })
        elif 'NHÂN CÔNG' in c3:
            current_task['nc'].append({
                'r_hp': r_idx,
                'code': c2,
                'name': c5,
                'dvt': c6,
                'col_norm': 'I'
            })
        elif 'MÁY THI CÔNG' in c3:
            current_task['mc'].append({
                'r_hp': r_idx,
                'code': c2,
                'name': c5,
                'dvt': c6,
                'col_norm': 'K'
            })

print(f"    Đã nạp {len(tasks_data)} công tác từ 'hao phi VT'.")

# Xóa sheet 'Đơn giá chi tiết' cũ nếu có
if "Đơn giá chi tiết" in wb.sheetnames:
    wb.remove(wb["Đơn giá chi tiết"])

# Tạo sheet 'Đơn giá chi tiết' tại vị trí thứ 4 (sau 'Công trình', trước 'Định mức áp dụng')
ws_dg = wb.create_sheet(title="Đơn giá chi tiết", index=3)
ws_dg.views.sheetView[0].showGridLines = True

# Định dạng Styling
font_title = Font(name="Times New Roman", size=14, bold=True, color="1B365D")
font_sub = Font(name="Times New Roman", size=10, italic=True, color="333333")
font_proj = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")

font_task_header = Font(name="Times New Roman", size=10, bold=True, color="002060")
font_cat_header = Font(name="Times New Roman", size=10, bold=True, italic=True, color="1B365D")
font_subtotal = Font(name="Times New Roman", size=10, bold=True, color="003366")
font_total = Font(name="Times New Roman", size=10, bold=True, color="B22222")
font_regular = Font(name="Times New Roman", size=10)
font_bold = Font(name="Times New Roman", size=10, bold=True)

fill_th = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
fill_task_header = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") # Xanh nhạt F1
fill_subtotal = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Vàng cam nhạt Total

border_thin = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
border_header = Border(
    left=Side(style='thin', color='FFFFFF'),
    right=Side(style='thin', color='FFFFFF'),
    top=Side(style='medium', color='1F497D'),
    bottom=Side(style='medium', color='1F497D')
)
border_subtotal = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='1F497D'),
    bottom=Side(style='thin', color='1F497D')
)
border_total = Border(
    left=Side(style='thin', color='1F497D'),
    right=Side(style='thin', color='1F497D'),
    top=Side(style='thin', color='1F497D'),
    bottom=Side(style='double', color='1F497D')
)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

print(">>> [2/6] Đang khởi tạo tiêu đề bảng 'Đơn giá chi tiết'...")

# Dòng 1: Tiêu đề bảng
ws_dg.merge_cells("A1:I1")
ws_dg["A1"] = "BẢNG ĐƠN GIÁ CHI TIẾT"
ws_dg["A1"].font = font_title
ws_dg["A1"].alignment = align_center
ws_dg.row_dimensions[1].height = 26

# Dòng 2: Căn cứ thông tư
ws_dg.merge_cells("A2:I2")
ws_dg["A2"] = "(Ban hành kèm theo Thông tư số 36/2026/TT-BXD ngày 26/06/2026 của Bộ Xây dựng)"
ws_dg["A2"].font = font_sub
ws_dg["A2"].alignment = align_center
ws_dg.row_dimensions[2].height = 18

# Dòng 3: Tên công trình
ws_dg.merge_cells("A3:I3")
ws_dg["A3"] = "CÔNG TRÌNH: SỬA CHỮA LỚN HỆ THỐNG PCCC NĂM 2026 – CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4"
ws_dg["A3"].font = font_proj
ws_dg["A3"].alignment = align_center
ws_dg.row_dimensions[3].height = 20

# Dòng 4: Hạng mục
ws_dg.merge_cells("A4:I4")
ws_dg["A4"] = "HẠNG MỤC: TOÀN BỘ 856 DỊCH VỤ PCCC (VĨNH TÂN 4 VÀ PHƯỚC THÁI 1, 2, 3)"
ws_dg["A4"].font = font_proj
ws_dg["A4"].alignment = align_center
ws_dg.row_dimensions[4].height = 20

# Dòng 5: Header bảng
headers = [
    ("A5", "STT"),
    ("B5", "Mã hiệu\nđơn giá"),
    ("C5", "Mã hiệu\nVL, NC, M"),
    ("D5", "Tên công tác / Thành phần hao phí"),
    ("E5", "Đơn vị"),
    ("F5", "Định mức"),
    ("G5", "Đơn giá"),
    ("H5", "Hệ số"),
    ("I5", "Thành tiền\n(đồng)"),
]

ws_dg.row_dimensions[5].height = 28
for cell_ref, text in headers:
    c = ws_dg[cell_ref]
    c.value = text
    c.font = font_th
    c.fill = fill_th
    c.alignment = align_center
    c.border = border_header

# Thiết lập độ rộng cột
col_widths = {
    'A': 6,   # STT
    'B': 16,  # Mã hiệu đơn giá
    'C': 18,  # Mã hiệu VL, NC, M
    'D': 52,  # Tên công tác / hao phí
    'E': 10,  # ĐVT
    'F': 12,  # Định mức
    'G': 14,  # Đơn giá
    'H': 8,   # Hệ số
    'I': 16,  # Thành tiền
}
for col_let, width in col_widths.items():
    ws_dg.column_dimensions[col_let].width = width

print(">>> [3/6] Đang tạo 856 khối chiết tính đơn giá chi tiết...")

write_r = 6
task_price_links = {} # stt -> {'sub_vl': row_or_0, 'sub_nc': row_or_0, 'sub_mc': row_or_0, 'total': row}

for t in tasks_data:
    stt = t['stt']
    r_ct = t['r_ct']
    is_pt = t['is_pt']
    vl_list = t['vl']
    nc_list = t['nc']
    mc_list = t['mc']
    
    r_task_header = write_r
    
    # 1. Dòng Header công tác
    ws_dg.row_dimensions[write_r].height = 22
    ws_dg.cell(write_r, 1, stt).alignment = align_center
    ws_dg.cell(write_r, 2, f"='Công trình'!C{r_ct}").alignment = align_center
    ws_dg.cell(write_r, 3, "").alignment = align_center
    ws_dg.cell(write_r, 4, f"='Công trình'!D{r_ct}").alignment = align_left
    ws_dg.cell(write_r, 5, f"='Công trình'!E{r_ct}").alignment = align_center
    ws_dg.cell(write_r, 6, "").alignment = align_right
    ws_dg.cell(write_r, 7, "").alignment = align_right
    ws_dg.cell(write_r, 8, "").alignment = align_center
    # Thành tiền dòng header sẽ link đến dòng Total
    
    for c_i in range(1, 10):
        c_cell = ws_dg.cell(write_r, c_i)
        c_cell.font = font_task_header
        c_cell.fill = fill_task_header
        c_cell.border = border_thin
    write_r += 1
    
    r_sub_vl = None
    r_sub_nc = None
    r_sub_mc = None
    
    # 2. Khối Vật liệu
    if vl_list:
        # Tiêu đề mục Vật liệu
        ws_dg.row_dimensions[write_r].height = 19
        ws_dg.cell(write_r, 4, "Vật liệu").font = font_cat_header
        for c_i in range(1, 10):
            ws_dg.cell(write_r, c_i).border = border_thin
        write_r += 1
        
        start_vl = write_r
        for v in vl_list:
            ws_dg.row_dimensions[write_r].height = 19
            ws_dg.cell(write_r, 1, "").alignment = align_center
            ws_dg.cell(write_r, 2, "").alignment = align_center
            ws_dg.cell(write_r, 3, v['code']).alignment = align_center
            ws_dg.cell(write_r, 4, v['name']).alignment = align_left
            ws_dg.cell(write_r, 5, v['dvt']).alignment = align_center
            
            # Định mức = Mức hao phí * Hệ số
            c_norm = ws_dg.cell(write_r, 6, f"='hao phi VT'!{v['col_norm']}{v['r_hp']}*'hao phi VT'!L{v['r_hp']}")
            c_norm.alignment = align_right
            c_norm.number_format = '#,##0.000'
            
            # Đơn giá = link Master Vật liệu
            c_price = ws_dg.cell(write_r, 7, f"={get_vl_master_cell(v['code'])}")
            c_price.alignment = align_right
            c_price.number_format = '#,##0'
            
            # Hệ số = 1
            c_hs = ws_dg.cell(write_r, 8, 1)
            c_hs.alignment = align_center
            c_hs.number_format = '0.00'
            
            # Thành tiền = ROUND(F*G*H, 0)
            c_tt = ws_dg.cell(write_r, 9, f"=ROUND(F{write_r}*G{write_r}*H{write_r}, 0)")
            c_tt.alignment = align_right
            c_tt.number_format = '#,##0'
            
            for c_i in range(1, 10):
                ws_dg.cell(write_r, c_i).font = font_regular
                ws_dg.cell(write_r, c_i).border = border_thin
            write_r += 1
        end_vl = write_r - 1
        
        # Dòng Subtotal Vật liệu
        r_sub_vl = write_r
        ws_dg.row_dimensions[write_r].height = 20
        ws_dg.cell(write_r, 4, "Cộng Vật liệu").alignment = align_left
        c_sub = ws_dg.cell(write_r, 9, f"=SUM(I{start_vl}:I{end_vl})")
        c_sub.alignment = align_right
        c_sub.number_format = '#,##0'
        
        for c_i in range(1, 10):
            cell_obj = ws_dg.cell(write_r, c_i)
            cell_obj.font = font_subtotal
            cell_obj.fill = fill_subtotal
            cell_obj.border = border_subtotal
        write_r += 1
        
    # 3. Khối Nhân công
    if nc_list:
        # Tiêu đề mục Nhân công
        ws_dg.row_dimensions[write_r].height = 19
        ws_dg.cell(write_r, 4, "Nhân công").font = font_cat_header
        for c_i in range(1, 10):
            ws_dg.cell(write_r, c_i).border = border_thin
        write_r += 1
        
        start_nc = write_r
        for n in nc_list:
            ws_dg.row_dimensions[write_r].height = 19
            ws_dg.cell(write_r, 1, "").alignment = align_center
            ws_dg.cell(write_r, 2, "").alignment = align_center
            ws_dg.cell(write_r, 3, n['code']).alignment = align_center
            ws_dg.cell(write_r, 4, n['name']).alignment = align_left
            ws_dg.cell(write_r, 5, n['dvt']).alignment = align_center
            
            # Định mức = Mức hao phí * Hệ số
            c_norm = ws_dg.cell(write_r, 6, f"='hao phi VT'!{n['col_norm']}{n['r_hp']}*'hao phi VT'!L{n['r_hp']}")
            c_norm.alignment = align_right
            c_norm.number_format = '#,##0.000'
            
            # Đơn giá = link Master Nhân công
            c_price = ws_dg.cell(write_r, 7, f"={get_nc_master_cell(n['code'], is_pt)}")
            c_price.alignment = align_right
            c_price.number_format = '#,##0'
            
            # Hệ số = 1
            c_hs = ws_dg.cell(write_r, 8, 1)
            c_hs.alignment = align_center
            c_hs.number_format = '0.00'
            
            # Thành tiền = ROUND(F*G*H, 0)
            c_tt = ws_dg.cell(write_r, 9, f"=ROUND(F{write_r}*G{write_r}*H{write_r}, 0)")
            c_tt.alignment = align_right
            c_tt.number_format = '#,##0'
            
            for c_i in range(1, 10):
                ws_dg.cell(write_r, c_i).font = font_regular
                ws_dg.cell(write_r, c_i).border = border_thin
            write_r += 1
        end_nc = write_r - 1
        
        # Dòng Subtotal Nhân công
        r_sub_nc = write_r
        ws_dg.row_dimensions[write_r].height = 20
        ws_dg.cell(write_r, 4, "Cộng Nhân công").alignment = align_left
        c_sub = ws_dg.cell(write_r, 9, f"=SUM(I{start_nc}:I{end_nc})")
        c_sub.alignment = align_right
        c_sub.number_format = '#,##0'
        
        for c_i in range(1, 10):
            cell_obj = ws_dg.cell(write_r, c_i)
            cell_obj.font = font_subtotal
            cell_obj.fill = fill_subtotal
            cell_obj.border = border_subtotal
        write_r += 1

    # 4. Khối Máy thi công
    if mc_list:
        # Tiêu đề mục Máy thi công
        ws_dg.row_dimensions[write_r].height = 19
        ws_dg.cell(write_r, 4, "Máy thi công").font = font_cat_header
        for c_i in range(1, 10):
            ws_dg.cell(write_r, c_i).border = border_thin
        write_r += 1
        
        start_mc = write_r
        for m in mc_list:
            ws_dg.row_dimensions[write_r].height = 19
            ws_dg.cell(write_r, 1, "").alignment = align_center
            ws_dg.cell(write_r, 2, "").alignment = align_center
            ws_dg.cell(write_r, 3, m['code']).alignment = align_center
            ws_dg.cell(write_r, 4, m['name']).alignment = align_left
            ws_dg.cell(write_r, 5, m['dvt']).alignment = align_center
            
            # Định mức = Mức hao phí * Hệ số
            c_norm = ws_dg.cell(write_r, 6, f"='hao phi VT'!{m['col_norm']}{m['r_hp']}*'hao phi VT'!L{m['r_hp']}")
            c_norm.alignment = align_right
            c_norm.number_format = '#,##0.000'
            
            # Đơn giá = link Master Máy thi công
            c_price = ws_dg.cell(write_r, 7, f"={get_mc_master_cell(m['code'], is_pt)}")
            c_price.alignment = align_right
            c_price.number_format = '#,##0'
            
            # Hệ số = 1
            c_hs = ws_dg.cell(write_r, 8, 1)
            c_hs.alignment = align_center
            c_hs.number_format = '0.00'
            
            # Thành tiền = ROUND(F*G*H, 0)
            c_tt = ws_dg.cell(write_r, 9, f"=ROUND(F{write_r}*G{write_r}*H{write_r}, 0)")
            c_tt.alignment = align_right
            c_tt.number_format = '#,##0'
            
            for c_i in range(1, 10):
                ws_dg.cell(write_r, c_i).font = font_regular
                ws_dg.cell(write_r, c_i).border = border_thin
            write_r += 1
        end_mc = write_r - 1
        
        # Dòng Subtotal Máy thi công
        r_sub_mc = write_r
        ws_dg.row_dimensions[write_r].height = 20
        ws_dg.cell(write_r, 4, "Cộng Máy thi công").alignment = align_left
        c_sub = ws_dg.cell(write_r, 9, f"=SUM(I{start_mc}:I{end_mc})")
        c_sub.alignment = align_right
        c_sub.number_format = '#,##0'
        
        for c_i in range(1, 10):
            cell_obj = ws_dg.cell(write_r, c_i)
            cell_obj.font = font_subtotal
            cell_obj.fill = fill_subtotal
            cell_obj.border = border_subtotal
        write_r += 1

    # 5. Dòng Total công tác (Đơn giá trực tiếp 1 ĐVT)
    r_total = write_r
    ws_dg.row_dimensions[write_r].height = 22
    ws_dg.cell(write_r, 4, "Đơn giá trực tiếp (1 đơn vị công tác)").alignment = align_left
    
    # Công thức cộng các subtotal:
    subtotals = []
    if r_sub_vl:
        subtotals.append(f"I{r_sub_vl}")
    if r_sub_nc:
        subtotals.append(f"I{r_sub_nc}")
    if r_sub_mc:
        subtotals.append(f"I{r_sub_mc}")
        
    formula_total = f"={'+'.join(subtotals)}" if subtotals else "=0"
    c_tot = ws_dg.cell(write_r, 9, formula_total)
    c_tot.alignment = align_right
    c_tot.number_format = '#,##0'
    
    for c_i in range(1, 10):
        cell_obj = ws_dg.cell(write_r, c_i)
        cell_obj.font = font_total
        cell_obj.fill = fill_total
        cell_obj.border = border_total
        
    # Gán vào ô I của dòng Header công tác
    ws_dg.cell(r_task_header, 9, f"=I{r_total}").number_format = '#,##0'
    ws_dg.cell(r_task_header, 9).alignment = align_right
    
    # Lưu lại để link sang sheet 'Công trình'
    task_price_links[stt] = {
        'sub_vl': r_sub_vl,
        'sub_nc': r_sub_nc,
        'sub_mc': r_sub_mc,
        'total': r_total
    }
    
    write_r += 1

print(f"    [OK] Đã tạo thành công {len(tasks_data)} khối chiết tính, tổng cộng {write_r - 1} dòng.")

# Freeze panes tại F6
ws_dg.freeze_panes = "F6"

print(">>> [4/6] Đang cập nhật đơn giá trực tiếp sang sheet 'Công trình'...")

for stt, lks in task_price_links.items():
    r_ct = stt + 6 # Dòng tương ứng trong sheet 'Công trình'
    
    # Cột G: Đơn giá Vật liệu
    if lks['sub_vl']:
        ws_ct.cell(r_ct, 7, f"='Đơn giá chi tiết'!I{lks['sub_vl']}")
    else:
        ws_ct.cell(r_ct, 7, 0)
        
    # Cột H: Đơn giá Nhân công
    if lks['sub_nc']:
        ws_ct.cell(r_ct, 8, f"='Đơn giá chi tiết'!I{lks['sub_nc']}")
    else:
        ws_ct.cell(r_ct, 8, 0)
        
    # Cột I: Đơn giá Máy thi công
    if lks['sub_mc']:
        ws_ct.cell(r_ct, 9, f"='Đơn giá chi tiết'!I{lks['sub_mc']}")
    else:
        ws_ct.cell(r_ct, 9, 0)
        
    # Cột P: Đơn giá tổng hợp
    ws_ct.cell(r_ct, 16, f"=G{r_ct}+H{r_ct}+I{r_ct}")

print("    [OK] Đã liên kết xong 856 đơn giá trực tiếp vào sheet 'Công trình'.")

print(">>> [5/6] Đang lưu workbook...")
saved_path = PROJECT_PATH
try:
    wb.save(PROJECT_PATH)
    print(f"    [OK] Đã lưu file chính: {PROJECT_PATH}")
except PermissionError:
    alt_project = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"
    wb.save(alt_project)
    saved_path = alt_project
    print(f"    [LƯU Ý] File chính đang mở trong Excel! Đã lưu bản cập nhật sang: {alt_project}")

print(">>> [6/6] Đang sao chép bản sao sang Downloads...")
try:
    shutil.copy2(saved_path, DOWNLOADS_PATH)
    print(f"    [OK] Đã sao chép sang: {DOWNLOADS_PATH}")
except PermissionError:
    shutil.copy2(saved_path, ALT_DOWNLOADS)
    print(f"    [OK] File Downloads đang mở trong Excel, đã copy sang: {ALT_DOWNLOADS}")

print("Xây dựng sheet 'Đơn giá chi tiết' và cập nhật 'Công trình' hoàn tất 100%!")
