# -*- coding: utf-8 -*-
"""
Script build_sheet_xuat_toan_nc.py
Xây dựng sheet 'Xuất toán Nhân công' trong workbook:
- STT (1..856)
- Tên công tác
- ĐVT, Khối lượng (Q)
- Giá trị số công (công thức link trực tiếp từ 'hao phi VT' cột N)
- Giá trị thành tiền của công (công thức link từ 'Công trình' cột K)
- Tọa độ ô số công tại 'hao phi VT'
- Formula nguyên bản của giá trị số công tại 'hao phi VT'
"""
import os
import sys
import io
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PROJECT_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
DOWNLOADS_PATH = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
ALT_DOWNLOADS = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"

print(">>> [1/5] Đang mở workbook và trích xuất dữ liệu...")
wb = openpyxl.load_workbook(PROJECT_PATH)

ws_ct = wb['Công trình']
ws_hp = wb['hao phi VT']

# Thu thập các dòng nhân công cho từng công tác từ 'hao phi VT'
tasks_data = {} # stt -> {'r_ct': r, 'name': ..., 'unit': ..., 'qty': ..., 'nc_rows': [{'r_hp': ..., 'formula_n': ...}]}

for r in range(7, 863):
    stt = ws_ct.cell(r, 1).value
    name = ws_ct.cell(r, 4).value
    unit = ws_ct.cell(r, 5).value
    qty = ws_ct.cell(r, 6).value
    if stt:
        stt_int = int(stt)
        tasks_data[stt_int] = {
            'stt': stt_int,
            'r_ct': r,
            'name': str(name or '').strip(),
            'unit': str(unit or '').strip(),
            'qty': qty,
            'nc_rows': []
        }

current_stt = 0
for r in range(7, ws_hp.max_row + 1):
    c_type = str(ws_hp.cell(r, 3).value or '').strip().upper()
    if c_type == 'CÔNG TÁC':
        current_stt += 1
        continue
    if current_stt in tasks_data and 'NHÂN CÔNG' in c_type:
        formula_n = str(ws_hp.cell(r, 14).value or '')
        hs = ws_hp.cell(r, 12).value
        tasks_data[current_stt]['nc_rows'].append({
            'r_hp': r,
            'formula_n': formula_n,
            'hs': hs
        })

print(f"    Đã nạp {len(tasks_data)} công tác với đầy đủ tọa độ dòng nhân công.")

# Xóa sheet cũ nếu có
for s_old in ["Xuất toán Nhân công", "XuatToan_NhanCong", "XuatToan_NC"]:
    if s_old in wb.sheetnames:
        wb.remove(wb[s_old])

# Tạo sheet 'Xuất toán Nhân công' ngay sau 'Quy đổi ĐVT'
idx_pos = wb.sheetnames.index("Quy đổi ĐVT") + 1 if "Quy đổi ĐVT" in wb.sheetnames else len(wb.sheetnames)
ws_xt = wb.create_sheet(title="Xuất toán Nhân công", index=idx_pos)
ws_xt.views.sheetView[0].showGridLines = True

# Định dạng Fonts
font_title = Font(name="Times New Roman", size=13, bold=True, color="1B365D")
font_proj = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Times New Roman", size=10, bold=True, color="000000")
font_regular = Font(name="Times New Roman", size=10, color="000000")
font_italic = Font(name="Times New Roman", size=10, italic=True, color="333333")
font_formula = Font(name="Consolas", size=9.5, color="002060")

# Fills
fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_subtotal = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

# Borders
border_thin = Border(
    left=Side(style='thin', color='A6A6A6'),
    right=Side(style='thin', color='A6A6A6'),
    top=Side(style='thin', color='A6A6A6'),
    bottom=Side(style='thin', color='A6A6A6')
)
border_header = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='medium', color='000000'),
    bottom=Side(style='medium', color='000000')
)
border_total = Border(
    left=Side(style='thin', color='000000'),
    right=Side(style='thin', color='000000'),
    top=Side(style='thin', color='000000'),
    bottom=Side(style='double', color='000000')
)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

print(">>> [2/5] Đang thiết lập cấu trúc sheet 'Xuất toán Nhân công'...")

# Tiêu đề
ws_xt.merge_cells("A1:G1")
ws_xt["A1"] = "BẢNG XUẤT TOÁN GIÁ TRỊ NHÂN CÔNG VÀ CÔNG THỨC HAO PHÍ (856 DỊCH VỤ PCCC)"
ws_xt["A1"].font = font_title
ws_xt["A1"].alignment = align_center
ws_xt.row_dimensions[1].height = 25

ws_xt.merge_cells("A2:G2")
ws_xt["A2"] = "CÔNG TRÌNH: SỬA CHỮA LỚN HỆ THỐNG PCCC NĂM 2026 – CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4"
ws_xt["A2"].font = font_proj
ws_xt["A2"].alignment = align_center
ws_xt.row_dimensions[2].height = 20

ws_xt.cell(3, 7, "Đơn vị tính số công: ngày công; thành tiền: đồng").alignment = align_right
ws_xt.cell(3, 7).font = font_italic
ws_xt.row_dimensions[3].height = 16

# Header bảng (Dòng 4)
headers = [
    ("A4", "STT"),
    ("B4", "Danh mục công tác PCCC"),
    ("C4", "ĐVT"),
    ("D4", "Khối lượng (Q)"),
    ("E4", "Giá trị số công\n(ngày công)"),
    ("F4", "Giá trị thành tiền của công\n(đồng)"),
    ("G4", "Formula của giá trị số công tại sheet 'hao phi VT'"),
]

ws_xt.row_dimensions[4].height = 30
for cell_ref, text in headers:
    c = ws_xt[cell_ref]
    c.value = text
    c.font = font_th
    c.fill = fill_header
    c.alignment = align_center
    c.border = border_header

# Thiết lập độ rộng cột
col_widths = {
    'A': 8,   # STT
    'B': 55,  # Tên công tác
    'C': 10,  # ĐVT
    'D': 14,  # Khối lượng Q
    'E': 18,  # Số công
    'F': 22,  # Thành tiền công
    'G': 40,  # Formula tại hao phi VT
}
for col_let, width in col_widths.items():
    ws_xt.column_dimensions[col_let].width = width

print(">>> [3/5] Đang điền dữ liệu 856 dịch vụ...")

write_r = 5
for stt in range(1, 857):
    t = tasks_data[stt]
    r_ct = t['r_ct']
    nc_list = t['nc_rows']
    
    ws_xt.row_dimensions[write_r].height = 20
    is_even = (stt % 2 == 0)
    row_fill = fill_zebra if is_even else None
    
    # Cột A: STT
    c_a = ws_xt.cell(write_r, 1, stt)
    c_a.alignment = align_center
    c_a.font = font_bold
    c_a.border = border_thin
    if row_fill: c_a.fill = row_fill
    
    # Cột B: Tên công tác (link từ 'Công trình')
    c_b = ws_xt.cell(write_r, 2, f"='Công trình'!D{r_ct}")
    c_b.alignment = align_left
    c_b.font = font_regular
    c_b.border = border_thin
    if row_fill: c_b.fill = row_fill
    
    # Cột C: ĐVT (link từ 'Công trình')
    c_c = ws_xt.cell(write_r, 3, f"='Công trình'!E{r_ct}")
    c_c.alignment = align_center
    c_c.font = font_regular
    c_c.border = border_thin
    if row_fill: c_c.fill = row_fill
    
    # Cột D: Khối lượng Q (link từ 'Công trình')
    c_d = ws_xt.cell(write_r, 4, f"='Công trình'!F{r_ct}")
    c_d.alignment = align_right
    c_d.font = font_regular
    c_d.number_format = '#,##0.##'
    c_d.border = border_thin
    if row_fill: c_d.fill = row_fill
    
    # Cột E: Giá trị số công
    # Link từ cột N của 'hao phi VT'
    if not nc_list:
        formula_days = 0
        formula_str_display = "Không có nhân công"
    elif len(nc_list) == 1:
        r_hp = nc_list[0]['r_hp']
        formula_days = f"='hao phi VT'!N{r_hp}"
        formula_str_display = f"='hao phi VT'!N{r_hp} [{nc_list[0]['formula_n']}]"
    else:
        # Tổng các dòng nhân công
        terms = [f"'hao phi VT'!N{item['r_hp']}" for item in nc_list]
        formula_days = f"={'+'.join(terms)}"
        formulas_disp = " + ".join([item['formula_n'] for item in nc_list if item['formula_n']])
        formula_str_display = f"={' + '.join(terms)} [{formulas_disp}]"
        
    c_e = ws_xt.cell(write_r, 5, formula_days)
    c_e.alignment = align_right
    c_e.font = font_bold
    c_e.number_format = '#,##0.00'
    c_e.border = border_thin
    if row_fill: c_e.fill = row_fill
    
    # Cột F: Giá trị thành tiền của công (link từ 'Công trình' cột K)
    c_f = ws_xt.cell(write_r, 6, f"='Công trình'!K{r_ct}")
    c_f.alignment = align_right
    c_f.font = font_bold
    c_f.number_format = '#,##0'
    c_f.border = border_thin
    if row_fill: c_f.fill = row_fill
    
    # Cột G: Formula của giá trị số công tại sheet 'hao phi VT'
    c_g = ws_xt.cell(write_r, 7, formula_str_display)
    c_g.alignment = align_left
    c_g.font = font_formula
    c_g.border = border_thin
    if row_fill: c_g.fill = row_fill
    
    write_r += 1

# Dòng Tổng cộng
ws_xt.row_dimensions[write_r].height = 24
ws_xt.merge_cells(f"A{write_r}:D{write_r}")
c_tot_lbl = ws_xt.cell(write_r, 1, "TỔNG CỘNG:")
c_tot_lbl.alignment = align_right
c_tot_lbl.font = font_bold

# Tổng số công: SUM(E5:E860)
c_tot_e = ws_xt.cell(write_r, 5, f"=SUM(E5:E{write_r-1})")
c_tot_e.alignment = align_right
c_tot_e.font = font_bold
c_tot_e.number_format = '#,##0.00'
c_tot_e.fill = fill_subtotal

# Tổng thành tiền công: SUM(F5:F860)
c_tot_f = ws_xt.cell(write_r, 6, f"=SUM(F5:F{write_r-1})")
c_tot_f.alignment = align_right
c_tot_f.font = font_bold
c_tot_f.number_format = '#,##0'
c_tot_f.fill = fill_subtotal

c_tot_g = ws_xt.cell(write_r, 7, "")

for col_idx in range(1, 8):
    ws_xt.cell(write_r, col_idx).border = border_total

print(">>> [4/5] Đang lưu workbook...")
saved_path = PROJECT_PATH
try:
    wb.save(PROJECT_PATH)
    print(f"    [OK] Đã lưu file chính: {PROJECT_PATH}")
except PermissionError:
    alt_project = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"
    wb.save(alt_project)
    saved_path = alt_project
    print(f"    [LƯU Ý] File chính đang mở trong Excel! Đã lưu bản cập nhật sang: {alt_project}")

print(">>> [5/5] Đang sao chép bản sao sang Downloads...")
try:
    shutil.copy2(saved_path, DOWNLOADS_PATH)
    print(f"    [OK] Đã sao chép sang: {DOWNLOADS_PATH}")
except PermissionError:
    shutil.copy2(saved_path, ALT_DOWNLOADS)
    print(f"    [OK] File Downloads đang mở trong Excel, đã copy sang: {ALT_DOWNLOADS}")

print("Hoàn tất xuất toán sheet 'Xuất toán Nhân công' 100%!")
