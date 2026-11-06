# -*- coding: utf-8 -*-
"""
Script build_sheet_quy_doi_dvt.py
Xây dựng sheet 'Quy đổi ĐVT' và thiết lập liên kết kiểm soát ĐVT:
1. Tạo sheet 'Quy đổi ĐVT' tại vị trí thứ 3 (ngay sau 02_Chi_Tiet_856_Dich_Vu và trước THDT).
2. Lập bảng so sánh, thuyết minh phân tích kỹ thuật và xác định hệ số quy đổi k_UOM cho các công tác chiếm tỷ trọng lớn:
   - STT 262: Bảo trì dây nhiệt TH78 tại LACP 9 (5.839 m -> quy đổi về 150 bộ module giám sát Zone).
   - STT 290: Bảo trì dây nhiệt TH78 tại LACP 10 (1.296 m -> quy đổi về 56 bộ module giám sát Zone).
3. Thiết lập liên kết động từ sheet 'Quy đổi ĐVT' sang sheet 'hao phi VT' (cột Hệ số L1029 và L1135).
4. Tự động cập nhật toàn bộ dây chuyền tính toán qua 'Đơn giá chi tiết', 'Công trình', và 'THDT'.
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

print(">>> [1/5] Đang mở workbook...")
wb = openpyxl.load_workbook(PROJECT_PATH)

# Xóa sheet 'Quy đổi ĐVT' cũ nếu có
for s_old in ["Quy đổi ĐVT", "QuyDoi_DVT", "Quy_Doi_DVT"]:
    if s_old in wb.sheetnames:
        wb.remove(wb[s_old])

# Tạo sheet 'Quy đổi ĐVT' tại vị trí thứ 3 (trước THDT)
ws_qd = wb.create_sheet(title="Quy đổi ĐVT", index=2)
ws_qd.views.sheetView[0].showGridLines = True

# Định dạng Fonts
font_main_title = Font(name="Times New Roman", size=13, bold=True, color="1B365D")
font_proj_title = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_sec_title = Font(name="Times New Roman", size=11, bold=True, color="002060")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Times New Roman", size=10, bold=True, color="000000")
font_regular = Font(name="Times New Roman", size=10, color="000000")
font_italic = Font(name="Times New Roman", size=10, italic=True, color="333333")
font_red_bold = Font(name="Times New Roman", size=10, bold=True, color="C00000")
font_green_bold = Font(name="Times New Roman", size=10, bold=True, color="008000")
font_diff_bold = Font(name="Times New Roman", size=10, bold=True, color="002060")

# Fills
fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid") # Xanh đậm tiêu chuẩn EVN
fill_highlight_save = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Xanh lá nhạt tiết kiệm
fill_highlight_warn = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Cam nhạt cảnh báo
fill_subtotal = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Xanh dương nhạt

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

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

print(">>> [2/5] Đang thiết lập cấu trúc sheet 'Quy đổi ĐVT'...")

# Tiêu đề
ws_qd.merge_cells("A1:N1")
ws_qd["A1"] = "BẢNG TỔNG HỢP VÀ THUYẾT MINH QUY ĐỔI ĐƠN VỊ TÍNH (ĐVT) KIỂM SOÁT CHI PHÍ NHÂN CÔNG"
ws_qd["A1"].font = font_main_title
ws_qd["A1"].alignment = align_center
ws_qd.row_dimensions[1].height = 25

ws_qd.merge_cells("A2:N2")
ws_qd["A2"] = "CÔNG TRÌNH: SỬA CHỮA LỚN HỆ THỐNG PCCC NĂM 2026 – CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4"
ws_qd["A2"].font = font_proj_title
ws_qd["A2"].alignment = align_center
ws_qd.row_dimensions[2].height = 20

ws_qd.merge_cells("A3:N3")
ws_qd["A3"] = "HẠNG MỤC: KIỂM SOÁT ĐỊNH MỨC & QUY ĐỔI THỨ NGUYÊN ĐVT CHO CÁC DỊCH VỤ CHI PHỐI TRỌNG YẾU"
ws_qd["A3"].font = font_italic
ws_qd["A3"].alignment = align_center
ws_qd.row_dimensions[3].height = 18

ws_qd.cell(4, 14, "Đơn vị tính: đồng").alignment = align_right
ws_qd.cell(4, 14).font = font_italic
ws_qd.row_dimensions[4].height = 16

# Headers Bảng Quy đổi (Dòng 5)
headers = [
    ("A5", "STT\nCT"),
    ("B5", "Mã hiệu\nđịnh mức"),
    ("C5", "Danh mục dịch vụ PCCC chi tiết"),
    ("D5", "ĐVT\nThiết kế"),
    ("E5", "Khối lượng\nThiết kế (Q_TK)"),
    ("F5", "ĐVT\nĐịnh mức"),
    ("G5", "Căn cứ & Cơ sở kỹ thuật quy đổi thứ nguyên"),
    ("H5", "Hệ số\nquy đổi (k_UOM)"),
    ("I5", "Khối lượng\nquy đổi (Q_QĐ)"),
    ("J5", "Định mức\n(công/ĐV)"),
    ("K5", "Đơn giá\n(đồng/công)"),
    ("L5", "Chi phí NC\nTRƯỚC quy đổi (đồng)"),
    ("M5", "Chi phí NC\nSAU quy đổi (đồng)"),
    ("N5", "Chênh lệch giảm trừ\n(Tiết kiệm CP ảo)"),
]

ws_qd.row_dimensions[5].height = 36
for cell_ref, text in headers:
    c = ws_qd[cell_ref]
    c.value = text
    c.font = font_th
    c.fill = fill_header
    c.alignment = align_center
    c.border = border_header

# Dòng 6: Tiêu đề nhóm
ws_qd.merge_cells("A6:N6")
ws_qd["A6"] = "I. CÁC CÔNG TÁC BẢO TRÌ DÂY NHIỆT BÁO CHÁY TUYẾN TÍNH PROTECTOWIRE (CHIẾM > 57% CHI PHÍ NHÂN CÔNG)"
ws_qd["A6"].font = font_sec_title
ws_qd["A6"].fill = fill_subtotal
ws_qd.row_dimensions[6].height = 24
for col_idx in range(1, 15):
    ws_qd.cell(6, col_idx).border = border_thin

# Dòng 7: STT 262 (LACP 9)
# 5839 m dây nhiệt Protectowire kết nối về 150 bộ module giám sát Apollo 55000-760 (STT 260)
ws_qd.row_dimensions[7].height = 42
ws_qd.cell(7, 1, 262).alignment = align_center
ws_qd.cell(7, 2, "E6C.3.03").alignment = align_center
ws_qd.cell(7, 3, "Bảo trì dây nhiệt TH78 của Hệ thống báo cháy tự động khu vực nhà máy chính NMNĐVT4 – Khu vực Các tháp và băng tải-LACP 9").alignment = align_left
ws_qd.cell(7, 4, "M").alignment = align_center
ws_qd.cell(7, 5, 5839).alignment = align_right
ws_qd.cell(7, 5).number_format = '#,##0'
ws_qd.cell(7, 6, "1 bộ").alignment = align_center
ws_qd.cell(7, 7, "Dây nhiệt Protectowire TH78 kết nối về 150 bộ module giám sát Apollo 55000-760 (STT 260) tại LACP 9. Quy đổi về 150 bộ module Zone theo đúng chỉ đạo định mức: k = 150 / 5839.").alignment = align_left
# Cột H: Hệ số quy đổi
c_h7 = ws_qd.cell(7, 8, "=150/E7")
c_h7.alignment = align_right
c_h7.number_format = '0.000000'
# Cột I: Khối lượng sau quy đổi = E7 * H7 = 150 bộ
c_i7 = ws_qd.cell(7, 9, "=ROUND(E7*H7, 2)")
c_i7.alignment = align_right
c_i7.number_format = '#,##0.00'
# Cột J: Định mức = 'Định mức áp dụng'!D134
c_j7 = ws_qd.cell(7, 10, "='Định mức áp dụng'!D134")
c_j7.alignment = align_right
c_j7.number_format = '0.00'
# Cột K: Đơn giá NC = 'Nhân công'!$G$498
c_k7 = ws_qd.cell(7, 11, "='Nhân công'!$G$498")
c_k7.alignment = align_right
c_k7.number_format = '#,##0'
# Cột L: Trước quy đổi = E7 * J7 * K7
c_l7 = ws_qd.cell(7, 12, "=ROUND(E7*J7*K7, 0)")
c_l7.alignment = align_right
c_l7.number_format = '#,##0'
c_l7.fill = fill_highlight_warn
c_l7.font = font_red_bold
# Cột M: Sau quy đổi = I7 * J7 * K7
c_m7 = ws_qd.cell(7, 13, "=ROUND(I7*J7*K7, 0)")
c_m7.alignment = align_right
c_m7.number_format = '#,##0'
c_m7.fill = fill_highlight_save
c_m7.font = font_green_bold
# Cột N: Chênh lệch giảm trừ = L7 - M7
c_n7 = ws_qd.cell(7, 14, "=L7-M7")
c_n7.alignment = align_right
c_n7.number_format = '#,##0'
c_n7.font = font_diff_bold

for col_idx in range(1, 15):
    ws_qd.cell(7, col_idx).border = border_thin
    if col_idx not in [12, 13, 14]:
        ws_qd.cell(7, col_idx).font = font_regular

# Dòng 8: STT 290 (LACP 10)
# 1296 m dây nhiệt Protectowire kết nối về 56 bộ module giám sát Apollo 55000-760 (STT 288)
ws_qd.row_dimensions[8].height = 42
ws_qd.cell(8, 1, 290).alignment = align_center
ws_qd.cell(8, 2, "E6C.3.03").alignment = align_center
ws_qd.cell(8, 3, "Bảo trì dây nhiệt TH78 của Hệ thống báo cháy tự động khu vực nhà máy chính NMNĐVT4 – Khu vực nhà LEB 3, tháp 1, 2, tháp mở rộng và kho than MR LACP 10").alignment = align_left
ws_qd.cell(8, 4, "M").alignment = align_center
ws_qd.cell(8, 5, 1296).alignment = align_right
ws_qd.cell(8, 5).number_format = '#,##0'
ws_qd.cell(8, 6, "1 bộ").alignment = align_center
ws_qd.cell(8, 7, "Dây nhiệt Protectowire TH78 kết nối về 56 bộ module giám sát Apollo 55000-760 (STT 288) tại LACP 10. Quy đổi về 56 bộ module Zone theo đúng chỉ đạo định mức: k = 56 / 1296.").alignment = align_left
# Cột H: Hệ số quy đổi
c_h8 = ws_qd.cell(8, 8, "=56/E8")
c_h8.alignment = align_right
c_h8.number_format = '0.000000'
# Cột I: Khối lượng sau quy đổi = E8 * H8 = 56 bộ
c_i8 = ws_qd.cell(8, 9, "=ROUND(E8*H8, 2)")
c_i8.alignment = align_right
c_i8.number_format = '#,##0.00'
# Cột J: Định mức = 'Định mức áp dụng'!D134
c_j8 = ws_qd.cell(8, 10, "='Định mức áp dụng'!D134")
c_j8.alignment = align_right
c_j8.number_format = '0.00'
# Cột K: Đơn giá NC = 'Nhân công'!$G$498
c_k8 = ws_qd.cell(8, 11, "='Nhân công'!$G$498")
c_k8.alignment = align_right
c_k8.number_format = '#,##0'
# Cột L: Trước quy đổi = E8 * J8 * K8
c_l8 = ws_qd.cell(8, 12, "=ROUND(E8*J8*K8, 0)")
c_l8.alignment = align_right
c_l8.number_format = '#,##0'
c_l8.fill = fill_highlight_warn
c_l8.font = font_red_bold
# Cột M: Sau quy đổi = I8 * J8 * K8
c_m8 = ws_qd.cell(8, 13, "=ROUND(I8*J8*K8, 0)")
c_m8.alignment = align_right
c_m8.number_format = '#,##0'
c_m8.fill = fill_highlight_save
c_m8.font = font_green_bold
# Cột N: Chênh lệch giảm trừ = L8 - M8
c_n8 = ws_qd.cell(8, 14, "=L8-M8")
c_n8.alignment = align_right
c_n8.number_format = '#,##0'
c_n8.font = font_diff_bold

for col_idx in range(1, 15):
    ws_qd.cell(8, col_idx).border = border_thin
    if col_idx not in [12, 13, 14]:
        ws_qd.cell(8, col_idx).font = font_regular

# Dòng 9: Tổng cộng
ws_qd.row_dimensions[9].height = 26
ws_qd.merge_cells("A9:D9")
ws_qd["A9"] = "TỔNG CỘNG CHÊNH LỆCH QUY ĐỔI ĐVT:"
ws_qd["A9"].font = font_bold
ws_qd["A9"].alignment = align_right

ws_qd.cell(9, 5, "=E7+E8").number_format = '#,##0'
ws_qd.cell(9, 5).font = font_bold
ws_qd.cell(9, 5).alignment = align_right

ws_qd.cell(9, 6, "")
ws_qd.cell(9, 7, "")
ws_qd.cell(9, 8, "")

ws_qd.cell(9, 9, "=I7+I8").number_format = '#,##0.00'
ws_qd.cell(9, 9).font = font_bold
ws_qd.cell(9, 9).alignment = align_right

ws_qd.cell(9, 10, "")
ws_qd.cell(9, 11, "")

ws_qd.cell(9, 12, "=L7+L8").number_format = '#,##0'
ws_qd.cell(9, 12).font = font_red_bold
ws_qd.cell(9, 12).alignment = align_right
ws_qd.cell(9, 12).fill = fill_highlight_warn

ws_qd.cell(9, 13, "=M7+M8").number_format = '#,##0'
ws_qd.cell(9, 13).font = font_green_bold
ws_qd.cell(9, 13).alignment = align_right
ws_qd.cell(9, 13).fill = fill_highlight_save

ws_qd.cell(9, 14, "=N7+N8").number_format = '#,##0'
ws_qd.cell(9, 14).font = font_diff_bold
ws_qd.cell(9, 14).alignment = align_right
ws_qd.cell(9, 14).fill = fill_subtotal

for col_idx in range(1, 15):
    ws_qd.cell(9, col_idx).border = border_total

# Phần II: Thuyết minh chi tiết (Dòng 11 trở đi)
ws_qd.row_dimensions[11].height = 22
ws_qd.merge_cells("A11:N11")
ws_qd["A11"] = "II. NGUYÊN TẮC VÀ CĂN CỨ KỸ THUẬT QUY ĐỔI ĐƠN VỊ TÍNH (THUYẾT MINH THẨM ĐỊNH)"
ws_qd["A11"].font = font_sec_title
ws_qd["A11"].alignment = align_left

notes = [
    "1. Bản chất kỹ thuật hệ thống báo cháy tháp & băng tải than (LACP 9 và LACP 10):",
    "   - Cảm biến nhiệt sử dụng loại cáp báo cháy tuyến tính Protectowire TH78 (dây nhiệt) trải dọc theo các dàn con lăn băng tải than.",
    "   - Toàn bộ tín hiệu của tuyến cáp nhiệt TH78 được kết nối và giám sát bởi các Hộp Module giám sát chuyển mạch mini Apollo 55000-760 (Mini Switch Monitor Module).",
    "   - Tại phân hệ LACP 9: Có 150 bộ module giám sát (tại STT 260) tương ứng quản lý 5.839 mét cáp nhiệt TH78 (bình quân ~38,9 m/module Zone).",
    "   - Tại phân hệ LACP 10: Có 56 bộ module giám sát (tại STT 288) tương ứng quản lý 1.296 mét cáp nhiệt TH78 (bình quân ~23,1 m/module Zone).",
    "",
    "2. Phân tích sự sai lệch do chưa quy đổi ĐVT trước đây:",
    "   - Mã định mức E6C.3.03 (Quyết định 45/QĐ-HĐTV và 177/QĐ-HĐTV của EVN) được xây dựng cho '1 Bộ' module giám sát Loop (định mức 2,71 ngày công/1 bộ gồm 11 bước công nghệ).",
    "   - Đơn vị lập dự toán đã lấy trực tiếp khối lượng chiều dài mét (5.839 m và 1.296 m) nhân với định mức 2,71 công/bộ, dẫn đến 1 mét dây nhiệt tốn tới 2,71 ngày công (sai lệch thứ nguyên nghiêm trọng).",
    "   - Việc này làm chi phí nhân công của riêng 2 công tác này bị đội khống lên tới 6.042.259.766 đồng (chiếm 57,05% tổng nhân công toàn bộ 856 dịch vụ!).",
    "",
    "3. Giải pháp quy đổi chuẩn hóa kỹ thuật và pháp lý (Áp dụng liên kết động vào hồ sơ):",
    "   - Quy đổi số lượng mét cáp nhiệt về đúng số bộ Module Zone bảo vệ thực tế theo ghi chú chuẩn: 'Quy đổi theo số bộ module':",
    "     + Tại STT 262: Hệ số quy đổi k_UOM = 150 bộ / 5.839 m = 0,02568933. Khối lượng sau quy đổi = 150 bộ.",
    "     + Tại STT 290: Hệ số quy đổi k_UOM = 56 bộ / 1.296 m = 0,04320988. Khối lượng sau quy đổi = 56 bộ.",
    "   - Chi phí nhân công sau quy đổi: STT 262 = 127.027.185 đồng; STT 290 = 47.423.482 đồng (Tổng cộng = 174.450.667 đồng).",
    "   - Tiết giảm trực tiếp chi phí nhân công ảo: 5.867.809.099 đồng (~ 5,87 TỶ ĐỒNG).",
    "   - Tổng giá trị dự toán toàn bộ công trình sau thuế và chi phí khác giảm từ ~ 13,405 tỷ đồng xuống ~ 5,984 tỷ đồng (Tiết kiệm hơn 7,42 TỶ ĐỒNG cho Chủ đầu tư).",
    "",
    "4. Tính liên kết hệ thống:",
    "   - Ô H7 và H8 tại sheet 'Quy đổi ĐVT' này được liên kết trực tiếp vào cột Hệ số (cột L) tại sheet 'hao phi VT' (dòng 1029 và 1135).",
    "   - Hệ thống công thức tại 'Đơn giá chi tiết', 'Công trình' và 'THDT' tự động cập nhật đồng bộ 100%, bảo đảm không đứt gãy liên kết."
]

for idx, note_line in enumerate(notes, start=12):
    ws_qd.row_dimensions[idx].height = 18
    ws_qd.merge_cells(f"A{idx}:N{idx}")
    cell = ws_qd[f"A{idx}"]
    cell.value = note_line
    cell.alignment = align_left
    if note_line.startswith(("1.", "2.", "3.", "4.")):
        cell.font = font_bold
    elif "Tiết giảm" in note_line or "Tiết kiệm" in note_line:
        cell.font = font_red_bold
    else:
        cell.font = font_regular

# Độ rộng các cột
col_widths = {
    'A': 8,   # STT
    'B': 14,  # Mã định mức
    'C': 42,  # Danh mục
    'D': 10,  # ĐVT TK
    'E': 14,  # Q_TK
    'F': 12,  # ĐVT ĐM
    'G': 46,  # Căn cứ quy đổi
    'H': 14,  # k_UOM
    'I': 14,  # Q_QĐ
    'J': 12,  # ĐM
    'K': 14,  # ĐG
    'L': 20,  # Trước QĐ
    'M': 20,  # Sau QĐ
    'N': 22,  # Giảm trừ
}
for col_let, width in col_widths.items():
    ws_qd.column_dimensions[col_let].width = width

print(">>> [3/5] Đang thiết lập liên kết động từ 'Quy đổi ĐVT' sang 'hao phi VT'...")
ws_hp = wb['hao phi VT']

# STT 262: dòng 1029 (NC), 1030 (VOM), 1031 (MEG)
# Cột L là cột 12 (Hệ số)
ws_hp.cell(1029, 12, "='Quy đổi ĐVT'!$H$7")
ws_hp.cell(1030, 12, "='Quy đổi ĐVT'!$H$7")
ws_hp.cell(1031, 12, "='Quy đổi ĐVT'!$H$7")

# STT 290: dòng 1135 (NC), 1136 (VOM), 1137 (MEG)
ws_hp.cell(1135, 12, "='Quy đổi ĐVT'!$H$8")
ws_hp.cell(1136, 12, "='Quy đổi ĐVT'!$H$8")
ws_hp.cell(1137, 12, "='Quy đổi ĐVT'!$H$8")

print("    [OK] Đã gắn công thức liên kết Hệ số k_UOM vào 'hao phi VT' dòng 1029-1031 và 1135-1137!")

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

print("Hoàn tất xây dựng sheet 'Quy đổi ĐVT' và thiết lập liên kết động 100%!")
