# -*- coding: utf-8 -*-
"""
Script build_sheet_thdt.py
Xây dựng sheet tổng hợp "THDT" (BẢNG TỔNG HỢP DỰ TOÁN HẠNG MỤC theo chuẩn Dự toán F1 / Thông tư 36/2026/TT-BXD):
1. Kết cấu bảng gồm các chi phí:
   - I. Chi phí trực tiếp (T = VL + NC + M)
     - 1. Chi phí vật liệu (VL = VLHT) -> Theo bảng tổng hợp VL (='Công trình'!J863)
     - 2. Chi phí nhân công (NC = NCHT) -> Theo bảng tổng hợp NC (='Công trình'!K863)
     - 3. Chi phí máy và thiết bị thi công (M = MHT) -> Theo bảng tổng hợp M (='Công trình'!L863)
   - II. Chi phí gián tiếp (GT = C + TT)
     - 1. Chi phí chung (C = T x 7,3%)
     - 2. Chi phí một số công việc không xác định được khối lượng từ thiết kế (TT = T x 2,5%)
   - III. Thu nhập chịu thuế tính trước (TL = (T + GT) x 5,5%)
   - Chi phí xây dựng trước thuế (G = T + GT + TL)
   - IV. Thuế giá trị gia tăng (GTGT = G x 8%)
   - CHI PHÍ XÂY DỰNG SAU THUẾ (Gxd = G + GTGT)
   - V. Chi phí nhà tạm để ở và điều hành thi công (LT = G x 1,1% x (1 + 8%))
   - Tổng cộng (Gxd + LT)
   - Làm tròn (=ROUND(Tổng cộng, -3))
   - Bằng chữ
   - Phần ký duyệt: Người lập, Người chủ trì...
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

print(">>> [1/4] Đang mở workbook...")
wb = openpyxl.load_workbook(PROJECT_PATH)

# Xóa sheet 'THDT' hoặc 'THDT hạng mục' cũ nếu có
for s_old in ["THDT", "THDT hạng mục", "THDT_HangMuc"]:
    if s_old in wb.sheetnames:
        wb.remove(wb[s_old])

# Tạo sheet 'THDT' ở vị trí thứ 3 (trước 'Công trình')
ws_thdt = wb.create_sheet(title="THDT", index=2)
ws_thdt.views.sheetView[0].showGridLines = True

# Định dạng Fonts
font_title = Font(name="Times New Roman", size=13, bold=True, color="000000")
font_proj = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_unit = Font(name="Times New Roman", size=10, italic=True, color="000000")
font_th = Font(name="Times New Roman", size=10, bold=True, color="000000")
font_bold = Font(name="Times New Roman", size=10, bold=True, color="000000")
font_regular = Font(name="Times New Roman", size=10, color="000000")
font_italic = Font(name="Times New Roman", size=10, italic=True, color="000000")
font_sign_title = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_sign_sub = Font(name="Times New Roman", size=10, italic=True, color="000000")
font_sign_name = Font(name="Times New Roman", size=11, bold=True, color="000000")

# Fills
fill_header = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # Xám sáng chuẩn F1
fill_highlight_total = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid") # Xanh mềm như hình

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
    top=Side(style='medium', color='000000'),
    bottom=Side(style='medium', color='000000')
)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

print(">>> [2/4] Đang thiết lập bảng Tổng hợp dự toán hạng mục...")

# Tiêu đề
ws_thdt.merge_cells("A1:E1")
ws_thdt["A1"] = "BẢNG TỔNG HỢP DỰ TOÁN HẠNG MỤC"
ws_thdt["A1"].font = font_title
ws_thdt["A1"].alignment = align_center
ws_thdt.row_dimensions[1].height = 24

ws_thdt.merge_cells("A2:E2")
ws_thdt["A2"] = "CÔNG TRÌNH: SỬA CHỮA LỚN HỆ THỐNG PCCC NĂM 2026 – CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4"
ws_thdt["A2"].font = font_proj
ws_thdt["A2"].alignment = align_center
ws_thdt.row_dimensions[2].height = 20

ws_thdt.merge_cells("A3:E3")
ws_thdt["A3"] = "HẠNG MỤC: TOÀN BỘ 856 DỊCH VỤ PCCC (VĨNH TÂN 4 VÀ PHƯỚC THÁI 1, 2, 3)"
ws_thdt["A3"].font = font_proj
ws_thdt["A3"].alignment = align_center
ws_thdt.row_dimensions[3].height = 20

ws_thdt.cell(4, 5, "Đơn vị tính: đồng").alignment = align_right
ws_thdt.cell(4, 5).font = font_unit
ws_thdt.row_dimensions[4].height = 18

# Header Bảng (Dòng 5)
headers = [
    ("A5", "STT"),
    ("B5", "NỘI DUNG CHI PHÍ"),
    ("C5", "CÁCH TÍNH"),
    ("D5", "GIÁ TRỊ"),
    ("E5", "KÝ HIỆU"),
]
ws_thdt.row_dimensions[5].height = 26
for cell_ref, text in headers:
    c = ws_thdt[cell_ref]
    c.value = text
    c.font = font_th
    c.fill = fill_header
    c.alignment = align_center
    c.border = border_header

# Dữ liệu các dòng chi phí (Dòng 6 đến 22)
rows_config = [
    # (r, stt, noidung, cachtinh, giatri_formula, kyhieu, is_bold, is_highlight_total)
    (6,  "I", "Chi phí trực tiếp", "VL + NC + M", "=D7+D9+D11", "T", True, False),
    (7,  "1", "Chi phí vật liệu", "VLHT", "=D8", "VL", True, False),
    (8,  "",  "  - Đơn giá vật liệu", "Theo bảng tổng hợp vật liệu", "='Công trình'!J863", "VLHT", False, False),
    (9,  "2", "Chi phí nhân công", "NCHT", "=D10", "NC", True, False),
    (10, "",  "  - Đơn giá nhân công", "Theo bảng tổng hợp nhân công", "='Công trình'!K863", "NCHT", False, False),
    (11, "3", "Chi phí máy và thiết bị thi công", "MHT", "=D12", "M", True, False),
    (12, "",  "  - Đơn giá máy thi công", "Theo bảng tổng hợp máy thi công", "='Công trình'!L863", "MHT", False, False),
    (13, "II", "Chi phí gián tiếp", "C + TT", "=D14+D15", "GT", True, False),
    (14, "1", "Chi phí chung", "T x 7,3%", "=ROUND(D6*7.3%, 0)", "C", False, False),
    (15, "2", "Chi phí một số công việc không xác định được khối lượng từ thiết kế", "T x 2,5%", "=ROUND(D6*2.5%, 0)", "TT", False, False),
    (16, "III", "Thu nhập chịu thuế tính trước", "(T + GT) x 5,5%", "=ROUND((D6+D13)*5.5%, 0)", "TL", True, False),
    (17, "",  "Chi phí xây dựng trước thuế", "T + GT + TL", "=D6+D13+D16", "G", True, False),
    (18, "IV", "Thuế giá trị gia tăng", "G x 8%", "=ROUND(D17*8%, 0)", "GTGT", True, False),
    (19, "",  "CHI PHÍ XÂY DỰNG SAU THUẾ", "G + GTGT", "=D17+D18", "Gxd", True, False),
    (20, "V", "Chi phí nhà tạm để ở và điều hành thi công", "G x 1,1% x (1 + 8%)", "=ROUND(D17*1.1%*(1+8%), 0)", "LT", True, False),
    (21, "",  "Tổng cộng", "Gxd + LT", "=D19+D20", "", True, True),
    (22, "",  "Làm tròn", "", "=ROUND(D21, -3)", "", True, False),
]

for r, stt, noidung, cachtinh, formula, kyhieu, is_bold, is_hl in rows_config:
    ws_thdt.row_dimensions[r].height = 22
    
    # Cột A: STT
    c_a = ws_thdt.cell(r, 1, stt)
    c_a.alignment = align_center
    c_a.font = font_bold if is_bold else font_regular
    c_a.border = border_thin
    
    # Cột B: Nội dung
    c_b = ws_thdt.cell(r, 2, noidung)
    c_b.alignment = align_right if r in [21, 22] else align_left
    c_b.font = font_bold if is_bold else font_regular
    c_b.border = border_thin
    
    # Cột C: Cách tính
    c_c = ws_thdt.cell(r, 3, cachtinh)
    c_c.alignment = align_center
    c_c.font = font_bold if is_bold else font_regular
    c_c.border = border_thin
    if is_hl:
        c_c.fill = fill_highlight_total
        
    # Cột D: Giá trị
    c_d = ws_thdt.cell(r, 4, formula)
    c_d.alignment = align_right
    c_d.font = font_bold if is_bold else font_regular
    c_d.number_format = '#,##0'
    c_d.border = border_total if is_hl else border_thin
    if is_hl:
        c_d.fill = fill_highlight_total
        
    # Cột E: Ký hiệu
    c_e = ws_thdt.cell(r, 5, kyhieu)
    c_e.alignment = align_center
    c_e.font = font_bold if is_bold else font_regular
    c_e.border = border_thin

# Dòng 23: Bằng chữ
ws_thdt.row_dimensions[23].height = 24
ws_thdt.merge_cells("A23:E23")
ws_thdt["A23"] = "Bằng chữ: (Xem số tiền chi tiết làm tròn tại ô D22 đồng./.)"
ws_thdt["A23"].font = font_italic
ws_thdt["A23"].alignment = align_left

# Phần ký duyệt (Dòng 25 đến 33)
ws_thdt.row_dimensions[25].height = 20
ws_thdt.merge_cells("D25:E25")
ws_thdt["D25"] = ".... , ngày .... tháng .... năm ......"
ws_thdt["D25"].font = font_italic
ws_thdt["D25"].alignment = align_center

ws_thdt.row_dimensions[27].height = 20
ws_thdt.cell(27, 2, "NGƯỜI LẬP").alignment = align_center
ws_thdt.cell(27, 2).font = font_sign_title

ws_thdt.merge_cells("D27:E27")
ws_thdt["D27"] = "NGƯỜI CHỦ TRÌ"
ws_thdt["D27"].font = font_sign_title
ws_thdt["D27"].alignment = align_center

ws_thdt.row_dimensions[31].height = 18
ws_thdt.cell(31, 2, "(Ký, họ tên)").alignment = align_center
ws_thdt.cell(31, 2).font = font_sign_sub

ws_thdt.merge_cells("D31:E31")
ws_thdt["D31"] = "(Ký, họ tên)"
ws_thdt["D31"].font = font_sign_sub
ws_thdt["D31"].alignment = align_center

ws_thdt.row_dimensions[32].height = 20
ws_thdt.cell(32, 2, "NGUYỄN VĂN A").alignment = align_center
ws_thdt.cell(32, 2).font = font_sign_name

ws_thdt.merge_cells("D32:E32")
ws_thdt["D32"] = "NGUYỄN VĂN B"
ws_thdt["D32"].font = font_sign_name
ws_thdt["D32"].alignment = align_center

ws_thdt.row_dimensions[33].height = 18
ws_thdt.merge_cells("D33:E33")
ws_thdt["D33"] = "Chứng chỉ KS định giá XD hạng ..., số ..."
ws_thdt["D33"].font = font_sign_sub
ws_thdt["D33"].alignment = align_center

# Độ rộng cột chuẩn
col_widths = {
    'A': 8,   # STT
    'B': 46,  # Nội dung chi phí
    'C': 32,  # Cách tính
    'D': 22,  # Giá trị
    'E': 12,  # Ký hiệu
}
for col_let, width in col_widths.items():
    ws_thdt.column_dimensions[col_let].width = width

print(">>> [3/4] Đang lưu workbook...")
saved_path = PROJECT_PATH
try:
    wb.save(PROJECT_PATH)
    print(f"    [OK] Đã lưu file chính: {PROJECT_PATH}")
except PermissionError:
    alt_project = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"
    wb.save(alt_project)
    saved_path = alt_project
    print(f"    [LƯU Ý] File chính đang mở trong Excel! Đã lưu bản cập nhật sang: {alt_project}")

print(">>> [4/4] Đang sao chép bản sao sang Downloads...")
try:
    shutil.copy2(saved_path, DOWNLOADS_PATH)
    print(f"    [OK] Đã sao chép sang: {DOWNLOADS_PATH}")
except PermissionError:
    shutil.copy2(saved_path, ALT_DOWNLOADS)
    print(f"    [OK] File Downloads đang mở trong Excel, đã copy sang: {ALT_DOWNLOADS}")

print("Xây dựng sheet tổng hợp THDT hoàn tất 100%!")
