# -*- coding: utf-8 -*-
"""
Script apply_khanh_hoa_for_phuoc_thai.py
Áp dụng Đơn giá Khánh Hòa (QĐ 646/QĐ-SXD ngày 03/7/2026) cho các công tác thuộc NMĐMT Phước Thái 1, 2, 3
trong file BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx:
1. Tạo các sheet Đơn giá Khánh Hòa:
   - KH_00_ThuyetMinh: Căn cứ pháp lý, phân vùng Ninh Phước thuộc KV3 (Vùng III).
   - KH_01_DonGia_NhanCong: Bảng giá NC chuẩn QĐ 646 & Thang bảng lương 7 bậc chi tiết.
   - KH_02_DonGia_CaMay: Bảng giá Ca máy Khánh Hòa.
2. Cập nhật Sheet Nhân công & Máy thi công:
   - Tách Master Row cho Vĩnh Tân 4 (Link Lâm Đồng Vùng IV) và Phước Thái 1, 2, 3 (Link Khánh Hòa KV3 Ninh Phước).
   - Bảo toàn 100% công thức link detail row sang 'hao phi VT'.
   - Master Row SUM các Detail rows tương ứng.
   - Thành tiền = ROUND(F*G, 0).
"""

import os
import sys
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PROJECT_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
DOWNLOADS_PATH = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"

print(">>> [1/6] Đang mở workbook Dự toán...")
wb = openpyxl.load_workbook(PROJECT_PATH, data_only=False)

# Styling
font_title = Font(name="Times New Roman", size=14, bold=True, color="1B365D")
font_sub = Font(name="Times New Roman", size=11, italic=True, color="333333")
font_sec = Font(name="Times New Roman", size=12, bold=True, color="0D47A1")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Times New Roman", size=10, bold=True)
font_regular = Font(name="Times New Roman", size=10)
font_italic = Font(name="Times New Roman", size=9, italic=True)

fill_primary = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
fill_teal = PatternFill(start_color="005F73", end_color="005F73", fill_type="solid")
fill_accent = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
fill_accent_pt = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid") # Peach soft for Phuoc Thai
fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

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

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# -------------------------------------------------------------
# 1. SHEET KH_00_ThuyetMinh
# -------------------------------------------------------------
print(">>> [2/6] Đang tạo Sheet KH_00_ThuyetMinh...")
if "KH_00_ThuyetMinh" in wb.sheetnames:
    wb.remove(wb["KH_00_ThuyetMinh"])

ws_kh0 = wb.create_sheet(title="KH_00_ThuyetMinh")
ws_kh0.views.sheetView[0].showGridLines = True

ws_kh0.merge_cells("A1:G1")
ws_kh0["A1"] = "CĂN CỨ ĐƠN GIÁ TỈNH KHÁNH HÒA ÁP DỤNG CHO DỰ ÁN NMĐMT PHƯỚC THÁI 1, 2, 3"
ws_kh0["A1"].font = font_title
ws_kh0["A1"].alignment = align_center
ws_kh0.row_dimensions[1].height = 28

ws_kh0.merge_cells("A2:G2")
ws_kh0["A2"] = "Dự án SCL Hệ thống PCCC năm 2026 – Các hạng mục tại Nhà máy điện mặt trời Phước Thái 1, 2, 3 (Ninh Phước)"
ws_kh0["A2"].font = font_sub
ws_kh0["A2"].alignment = align_center

kh_intro = [
    ("I. CĂN CỨ PHÁP LÝ ĐƠN GIÁ KHÁNH HÒA", "", font_sec, fill_accent),
    ("1. Đơn giá Nhân công", "Quyết định số 646/QĐ-SXD ngày 03/7/2026 của Giám đốc Sở Xây dựng tỉnh Khánh Hòa về việc công bố đơn giá nhân công xây dựng chuyển đổi trên địa bàn tỉnh Khánh Hòa theo Nghị định số 206/2026/NĐ-CP, Thông tư số 36/2026/TT-BXD, Thông tư số 37/2026/TT-BXD và Thông tư số 38/2026/TT-BXD của Bộ Xây dựng.", font_regular, None),
    ("2. Bảng giá Ca máy & TBTC", "Quyết định số 45/QĐ-SXD ngày 15/01/2026 và Quyết định số 486/QĐ-SXD của Sở Xây dựng Khánh Hòa công bố Bảng giá ca máy và thiết bị thi công trên địa bàn tỉnh Khánh Hòa.", font_regular, None),
    ("II. PHÂN VÙNG ĐỊA BÀN VÀ CĂN CỨ ÁP DỤNG CHO PHƯỚC THÁI", "", font_sec, fill_accent),
    ("Khu vực 3 (Vùng III)", "Theo Phụ lục ban hành kèm theo Quyết định số 646/QĐ-SXD ngày 03/7/2026 của Sở Xây dựng Khánh Hòa:\n'Khu vực 3 (vùng III), gồm các xã Đại Lãnh, Tu Bông, Vạn Thắng, Vạn Ninh, Vạn Hưng, Diên Khánh, Diên Lạc, Diên Điền, Suối Hiệp, Diên Thọ, Diên Lâm, Cam Lâm, Suối Dầu, Cam Hiệp, Cam An, Ninh Phước, Phước Hữu, Phước Hậu, Phước Dinh, Ninh Hải, Xuân Hải, Vĩnh Hải, Ninh Sơn, Lâm Sơn, Anh Dũng, Mỹ Sơn, Thuận Nam, Cà Ná, Phước Hà.'\n=> Địa bàn huyện Ninh Phước (nơi đặt Nhà máy điện mặt trời Phước Thái 1, 2, 3) áp dụng mức giá của KHU VỰC 3 (VÙNG III) theo QĐ 646/QĐ-SXD.", font_regular, None),
    ("Khu vực 2 (Vùng II)", "Gồm các phường thuộc Nha Trang, Cam Ranh, Ninh Hòa, Phan Rang, Đông Hải, Ninh Chử, Bảo An, Đô Vinh và các xã vùng đồng bằng phụ cận.", font_regular, None),
    ("Khu vực 4 (Vùng IV)", "Gồm các xã, phường vùng sâu còn lại trên địa bàn tỉnh Khánh Hòa.", font_regular, None),
    ("III. NGUYÊN TẮC ÁP DỤNG CHO 856 DỊCH VỤ", "", font_sec, fill_accent),
    ("Phần việc tại Vĩnh Tân 4", "Áp dụng Bộ Đơn giá tỉnh Lâm Đồng VÙNG IV (Văn bản số 5898/SXD-QLXD và Quyết định số 21/QĐ-SXD), link sang sheet '01_DonGia_NhanCong' và '02_DonGia_CaMay'.", font_regular, None),
    ("Phần việc tại Phước Thái 1, 2, 3", "Gồm 96 công tác dịch vụ PCCC tại NMĐMT Phước Thái 1 (STT 681-704), Phước Thái 2 (STT 705-739) và Phước Thái 3 (STT 740-776). Áp dụng Bộ Đơn giá tỉnh Khánh Hòa KHU VỰC 3 (VÙNG III - Ninh Phước), link sang sheet 'KH_01_DonGia_NhanCong' và 'KH_02_DonGia_CaMay'.", font_regular, None),
]

r_idx = 4
for item in kh_intro:
    title, desc, f_style, fill_s = item
    if not desc:
        ws_kh0.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=7)
        c = ws_kh0.cell(r_idx, 1, title)
        c.font = f_style
        if fill_s: c.fill = fill_s
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws_kh0.row_dimensions[r_idx].height = 24
    else:
        c1 = ws_kh0.cell(r_idx, 1, title)
        c1.font = font_bold
        c1.alignment = Alignment(horizontal='left', vertical='top')
        ws_kh0.merge_cells(start_row=r_idx, start_column=2, end_row=r_idx, end_column=7)
        c2 = ws_kh0.cell(r_idx, 2, desc)
        c2.font = font_regular
        c2.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws_kh0.row_dimensions[r_idx].height = 54 if len(desc) > 150 else 24
    r_idx += 1

ws_kh0.column_dimensions['A'].width = 28
for col_letter in ['B', 'C', 'D', 'E', 'F', 'G']:
    ws_kh0.column_dimensions[col_letter].width = 18

# -------------------------------------------------------------
# 2. SHEET KH_01_DonGia_NhanCong
# -------------------------------------------------------------
print(">>> [3/6] Đang tạo Sheet KH_01_DonGia_NhanCong...")
if "KH_01_DonGia_NhanCong" in wb.sheetnames:
    wb.remove(wb["KH_01_DonGia_NhanCong"])

ws_kh1 = wb.create_sheet(title="KH_01_DonGia_NhanCong")
ws_kh1.views.sheetView[0].showGridLines = True

ws_kh1.merge_cells("A1:G1")
ws_kh1["A1"] = "BẢNG ĐƠN GIÁ NHÂN CÔNG XÂY DỰNG TỈNH KHÁNH HÒA NĂM 2026"
ws_kh1["A1"].font = font_title
ws_kh1["A1"].alignment = align_center

ws_kh1.merge_cells("A2:G2")
ws_kh1["A2"] = "(Ban hành kèm theo Quyết định số 646/QĐ-SXD ngày 03/7/2026 của Sở Xây dựng Khánh Hòa)"
ws_kh1["A2"].font = font_sub
ws_kh1["A2"].alignment = align_center

# Phần A
ws_kh1.merge_cells("A4:G4")
ws_kh1["A4"] = "PHẦN A: BẢNG GIÁ NHÂN CÔNG CHUYỂN ĐỔI THEO NHÓM THEO QĐ SỐ 646/QĐ-SXD (BẬC BÌNH QUÂN 3,5/7)"
ws_kh1["A4"].font = font_sec
ws_kh1["A4"].fill = fill_accent
ws_kh1["A4"].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws_kh1.row_dimensions[4].height = 24

headers_kh1 = ["STT", "Nhóm nhân công", "Tên nhân công (Loại công tác áp dụng)", "ĐVT", "Khu vực 2 (Vùng II)", "Khu vực 3 (Vùng III - Ninh Phước)", "Khu vực 4 (Vùng IV)"]
ws_kh1.row_dimensions[5].height = 28
for c_idx, h in enumerate(headers_kh1, 1):
    cell = ws_kh1.cell(5, c_idx, h)
    cell.font = font_th
    cell.fill = fill_primary
    cell.alignment = align_center
    cell.border = border_header

kh_nc_data = [
    ("I", "NHÓM NHÂN CÔNG XÂY DỰNG", "", "", None, None, None, True),
    ("1", "Nhóm 1", "Chặt cây địa hình bằng phẳng, phá dỡ kết cấu bằng máy; đào móng thủ công; đắp đất nền móng, nền đường; xói hút đất từ tàu hút...", "công", 232470, 227549, 219374, False),
    ("2", "Nhóm 2", "Đóng cọc bê tông cốt thép; bê tông tường, dầm; cốt thép tường, dầm; ván khuôn thép; thảm bê tông nhựa; xây tường gạch...", "công", 289152, 273521, 260101, False),
    ("3", "Nhóm 3", "Gia công vì kèo, cột thép; trát tường, trụ, cột; ốp gạch, lát nền, sàn; lắp đặt hệ thống điện: quạt, đèn, điều hòa, PCCC...", "công", 302504, 286151, 272111, False),
    ("4", "Nhóm 4", "Rải căng dây đường dây tải điện; lắp đặt máy biến áp; thí nghiệm độ chặt; lắp đặt máy và thiết bị cần cẩu, cầu trục, đo lường...", "công", 325024, 307450, 293351, False),
    ("II", "NHÓM NHÂN CÔNG VẬN HÀNH, ĐIỀU KHIỂN MÁY", "", "", None, None, None, True),
    ("1", "Lái máy", "Nhân công vận hành máy, điều khiển máy thi công xây dựng", "công", 330984, 310681, 298428, False),
    ("2", "Lái xe", "Lái xe các loại", "công", 310500, 291454, 279959, False),
    ("3", "Thủy thủ / Thợ điện máy", "Thủy thủ, thợ máy, thợ điện", "công", 477000, 443000, 422000, False),
    ("4", "Tàu sông", "Máy trưởng, máy I, máy II, điện trưởng, kỹ thuật viên cuốc I, II", "công", 477000, 443000, 422000, False),
    ("5", "Thuyền trưởng", "Thuyền trưởng, thuyền phó", "công", 477000, 443000, 422000, False),
    ("III", "NHÓM NHÂN CÔNG THỰC HIỆN CÁC CÔNG TÁC KHÁC", "", "", None, None, None, True),
    ("1", "Kỹ sư KS-TN", "Kỹ sư thực hiện khảo sát, thí nghiệm", "công", 310818, 296438, 287342, False),
    ("2", "Thợ lặn", "Thợ lặn", "công", 628000, 587000, 562000, False),
    ("3", "Nghệ nhân", "Nghệ nhân", "công", 568000, 527000, 502000, False),
]

curr_r = 6
for row in kh_nc_data:
    stt, grp, desc, unit, v2, v3, v4, is_sec = row
    ws_kh1.row_dimensions[curr_r].height = 24 if is_sec else 28
    if is_sec:
        ws_kh1.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
        c = ws_kh1.cell(curr_r, 1, f"{stt}. {grp}")
        c.font = font_bold
        c.fill = fill_zebra
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        for c_idx in range(1, 8): ws_kh1.cell(curr_r, c_idx).border = border_thin
    else:
        ws_kh1.cell(curr_r, 1, stt).alignment = align_center
        ws_kh1.cell(curr_r, 2, grp).alignment = align_left
        ws_kh1.cell(curr_r, 3, desc).alignment = align_left
        ws_kh1.cell(curr_r, 4, unit).alignment = align_center
        for c_idx, val in enumerate([v2, v3, v4], 5):
            c_val = ws_kh1.cell(curr_r, c_idx, val)
            c_val.number_format = '#,##0'
            c_val.alignment = align_right
            if c_idx == 6: c_val.font = font_bold # Highlight KV3 Ninh Phuoc
        for c_idx in range(1, 8):
            cell = ws_kh1.cell(curr_r, c_idx)
            cell.font = font_bold if c_idx == 6 and not is_sec else font_regular
            cell.border = border_thin
    curr_r += 1

# Phần B: Thang bảng lương 7 bậc chi tiết
curr_r += 2
ws_kh1.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=8)
c_hdr2 = ws_kh1.cell(curr_r, 1, "PHẦN B: BẢNG QUY ĐỔI ĐƠN GIÁ NHÂN CÔNG TỈNH KHÁNH HÒA TỪ BẬC 1/7 ĐẾN BẬC 7/7 THEO THÔNG TƯ 38/2026/TT-BXD")
c_hdr2.font = font_sec
c_hdr2.fill = fill_accent
c_hdr2.alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws_kh1.row_dimensions[curr_r].height = 24
curr_r += 1

headers_kh2 = ["STT", "Nhóm nhân công", "Cấp bậc thợ", "Hệ số cấp bậc", "Hệ số quy đổi (so 3,5)", "Khu vực 2 (Vùng II)", "Khu vực 3 (VÙNG III - NINH PHƯỚC)", "Khu vực 4 (Vùng IV)"]
ws_kh1.row_dimensions[curr_r].height = 28
for c_idx, h in enumerate(headers_kh2, 1):
    cell = ws_kh1.cell(curr_r, c_idx, h)
    cell.font = font_th
    cell.fill = fill_teal
    cell.alignment = align_center
    cell.border = border_header
curr_r += 1

grades_scale = [
    ("Bậc 1,0/7", 1.00, 1.00/1.52),
    ("Bậc 2,0/7", 1.18, 1.18/1.52),
    ("Bậc 3,0/7", 1.39, 1.39/1.52),
    ("Bậc 3,5/7 (Chuẩn)", 1.52, 1.52/1.52),
    ("Bậc 4,0/7", 1.65, 1.65/1.52),
    ("Bậc 5,0/7", 1.95, 1.95/1.52),
    ("Bậc 6,0/7", 2.31, 2.31/1.52),
    ("Bậc 7,0/7", 2.73, 2.73/1.52),
]

kh_groups = [
    ("Nhóm 1", 232470, 227549, 219374),
    ("Nhóm 2", 289152, 273521, 260101),
    ("Nhóm 3", 302504, 286151, 272111),
    ("Nhóm 4", 325024, 307450, 293351),
    ("Lái máy", 330984, 310681, 298428),
    ("Lái xe", 310500, 291454, 279959),
]

stt_c = 1
for grp_name, base_v2, base_v3, base_v4 in kh_groups:
    for g_title, g_coef, g_ratio in grades_scale:
        is_std = "Chuẩn" in g_title
        ws_kh1.cell(curr_r, 1, stt_c).alignment = align_center
        ws_kh1.cell(curr_r, 2, grp_name).alignment = align_left
        ws_kh1.cell(curr_r, 3, g_title).alignment = align_center
        
        c_k = ws_kh1.cell(curr_r, 4, g_coef)
        c_k.number_format = '0.00'
        c_k.alignment = align_center
        
        c_r = ws_kh1.cell(curr_r, 5, g_ratio)
        c_r.number_format = '0.00000'
        c_r.alignment = align_center
        
        val_v2 = round(base_v2 * g_ratio)
        val_v3 = round(base_v3 * g_ratio)
        val_v4 = round(base_v4 * g_ratio)
        
        for c_idx, val in enumerate([val_v2, val_v3, val_v4], 6):
            c_val = ws_kh1.cell(curr_r, c_idx, val)
            c_val.number_format = '#,##0'
            c_val.alignment = align_right
            if c_idx == 7: # KV3 Ninh Phuoc
                c_val.font = font_bold
                if is_std: c_val.fill = fill_highlight
                
        for c_idx in range(1, 9):
            cell = ws_kh1.cell(curr_r, c_idx)
            cell.border = border_thin
            if is_std and c_idx != 7:
                cell.fill = fill_highlight
                cell.font = font_bold
            elif c_idx != 7:
                cell.font = font_regular
                
        ws_kh1.row_dimensions[curr_r].height = 20
        curr_r += 1
        stt_c += 1

ws_kh1.column_dimensions['A'].width = 8
ws_kh1.column_dimensions['B'].width = 16
ws_kh1.column_dimensions['C'].width = 40
ws_kh1.column_dimensions['D'].width = 12
ws_kh1.column_dimensions['E'].width = 18
ws_kh1.column_dimensions['F'].width = 22
ws_kh1.column_dimensions['G'].width = 28
ws_kh1.column_dimensions['H'].width = 22
ws_kh1.freeze_panes = "A6"

# -------------------------------------------------------------
# 3. SHEET KH_02_DonGia_CaMay
# -------------------------------------------------------------
print(">>> [4/6] Đang tạo Sheet KH_02_DonGia_CaMay...")
if "KH_02_DonGia_CaMay" in wb.sheetnames:
    wb.remove(wb["KH_02_DonGia_CaMay"])

ws_cm_ld = wb["02_DonGia_CaMay"]
ws_kh2 = wb.copy_worksheet(ws_cm_ld)
ws_kh2.title = "KH_02_DonGia_CaMay"

ws_kh2["A1"] = "BẢNG GIÁ CA MÁY VÀ THIẾT BỊ THI CÔNG TỈNH KHÁNH HÒA NĂM 2026"
ws_kh2["A2"] = "(Căn cứ Quyết định công bố giá Ca máy tỉnh Khánh Hòa – Áp dụng cho NMĐMT Phước Thái 1, 2, 3 tại Ninh Phước)"

ws_kh2.cell(4, 13, "GIÁ CA MÁY KHU VỰC 2 (đồng)")
ws_kh2.cell(4, 14, "GIÁ CA MÁY KHU VỰC 3 (NINH PHƯỚC)")
ws_kh2.cell(4, 15, "GIÁ CA MÁY KHU VỰC 4 (đồng)")

# -------------------------------------------------------------
# 4. TÁCH MASTER ROWS TRONG SHEET NHÂN CÔNG THEO KHU VỰC
# -------------------------------------------------------------
print(">>> [5/6] Đang tái cấu trúc Sheet 'Nhân công' & 'Máy thi công' tách bạch Vĩnh Tân 4 và Phước Thái...")

# Function to rebuild Master-Detail sheet with Regional Split
def rebuild_resource_sheet(sheet_name, is_nc=True):
    old_ws = wb[sheet_name]
    
    # Read existing rows into structured blocks
    # Block: {'master_info': (stt, code, name, loc, unit, g_val, h_val, note), 'vt_details': [], 'pt_details': []}
    blocks = []
    current_block = None
    
    for r in range(6, old_ws.max_row + 1):
        c1 = str(old_ws.cell(r, 1).value or '')
        c2 = str(old_ws.cell(r, 2).value or '')
        c3 = str(old_ws.cell(r, 3).value or '')
        c4 = str(old_ws.cell(r, 4).value or '')
        c5 = str(old_ws.cell(r, 5).value or '')
        c6_formula = str(old_ws.cell(r, 6).value or '')
        c7_price = str(old_ws.cell(r, 7).value or '')
        c8_total = str(old_ws.cell(r, 8).value or '')
        c9_note = str(old_ws.cell(r, 9).value or '')
        
        if c1.isdigit():
            # Master row
            current_block = {
                'stt': int(c1),
                'code': c2,
                'name': c3,
                'loc': c4,
                'unit': c5,
                'vt_details': [],
                'pt_details': []
            }
            blocks.append(current_block)
        else:
            # Detail row
            detail_data = (c1, c2, c3, c4, c5, c6_formula, c7_price, c8_total, c9_note)
            if 'phước thái' in c3.lower() or 'phước thái' in c4.lower():
                current_block['pt_details'].append(detail_data)
            else:
                current_block['vt_details'].append(detail_data)
                
    print(f"    Loaded {len(blocks)} master blocks from '{sheet_name}'.")
    
    # Clear data rows from row 6 downwards
    old_ws.delete_rows(6, old_ws.max_row - 5)
    
    # Mapping for Khanh Hoa KV3 Ninh Phuoc (Row coordinates in KH_01_DonGia_NhanCong Table B, Col G):
    # Nhóm 3:
    # Bậc 3,0/7: G44 (261.678 đ)
    # Bậc 3,5/7: G45 (286.151 đ)
    # Bậc 4,0/7: G46 (310.624 đ)
    # Bậc 5,0/7: G47 (367.102 đ)
    kh_nc_map = {
        "NC_E6C.3.05": ("='KH_01_DonGia_NhanCong'!G47", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 5,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G47"),
        "NC_E6C.3.02_B5": ("='KH_01_DonGia_NhanCong'!G47", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 5,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G47"),
        "NC_E6C.3.02_B1-11": ("='KH_01_DonGia_NhanCong'!G47", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 5,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G47"),
        "NC_BD.41111": ("='KH_01_DonGia_NhanCong'!G45", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 3,5/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G45"),
        "NC_E6C.3.01_B5": ("='KH_01_DonGia_NhanCong'!G47", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 5,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G47"),
        "NC_E6C.3.01_B1-11": ("='KH_01_DonGia_NhanCong'!G47", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 5,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G47"),
        "NC_BD.41131": ("='KH_01_DonGia_NhanCong'!G46", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 4,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G46"),
        "NC_BD.41141": ("='KH_01_DonGia_NhanCong'!G44", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 3,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G44"),
        "NC_E6C.3.03": ("='KH_01_DonGia_NhanCong'!G47", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 5,0/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G47"),
        "NC_BD.41121": ("='KH_01_DonGia_NhanCong'!G45", "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3, Bậc 3,5/7 KV3 Ninh Phước) - Link 'KH_01_DonGia_NhanCong'!G45"),
    }
    
    # Lâm Đồng Vùng IV mapping for NC (Col H in 01_DonGia_NhanCong):
    ld_nc_map = {
        "NC_E6C.3.05": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_E6C.3.02_B5": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_E6C.3.02_B1-11": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_BD.41111": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
        "NC_E6C.3.01_B5": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_E6C.3.01_B1-11": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_BD.41131": ("='01_DonGia_NhanCong'!H46", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 4,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H46"),
        "NC_BD.41141": ("='01_DonGia_NhanCong'!H44", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H44"),
        "NC_E6C.3.03": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_BD.41121": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
        "NC_E6C.1.23": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
        "NC_E6C.1.31": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
        "NC_E6C.2.04": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
        "NC_E6C.1.18": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
        "NC_BA.18405": ("='01_DonGia_NhanCong'!H52", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 3,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H52"),
        "NC_E6C.3.04": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_E6B.1.13.01": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
        "NC_E6C.3.06": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_DB.24001": ("='01_DonGia_NhanCong'!G19", "VB 5898/SXD-QLXD Lâm Đồng (Kỹ sư KS-TN VÙNG IV) - Link '01_DonGia_NhanCong'!G19"),
        "NC_E6BOP.1.03": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
        "NC_E6BOP.1.01": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
        "NC_BD.41171": ("='01_DonGia_NhanCong'!H52", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 3,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H52"),
        "NC_SB.82633": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
        "NC_BB.86601": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
        "NC_BB.86103": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
        "NC_BB.85201": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
        "NC_SB.82634": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    }
    
    # Mapping for Ca máy
    kh_cm_map = {
        "M_E6C.3.05_VOM": ("='KH_02_DonGia_CaMay'!N677", "QĐ Ca máy Khánh Hòa (M203.0019: Máy đo vạn năng KV3) - Link 'KH_02_DonGia_CaMay'!N677"),
        "M_E6C.3.05_BUI": ("='KH_02_DonGia_CaMay'!N459", "QĐ Ca máy Khánh Hòa (M112.4301: Máy cầm tay KV3) - Link 'KH_02_DonGia_CaMay'!N459"),
        "M_E6C.3.02_SOLO": ("='KH_02_DonGia_CaMay'!N684", "QĐ Ca máy Khánh Hòa (M203.0026: TB kiểm tra thử nghiệm KV3) - Link 'KH_02_DonGia_CaMay'!N684"),
        "M_E6C.3.02_VOM": ("='KH_02_DonGia_CaMay'!N677", "QĐ Ca máy Khánh Hòa (M203.0019: Máy đo vạn năng KV3) - Link 'KH_02_DonGia_CaMay'!N677"),
        "M_BD.41111": ("='KH_02_DonGia_CaMay'!N684", "QĐ Ca máy Khánh Hòa (M203.0026: TB kiểm tra áp lực KV3) - Link 'KH_02_DonGia_CaMay'!N684"),
        "M_E6C.3.01_MEG": ("='KH_02_DonGia_CaMay'!N683", "QĐ Ca máy Khánh Hòa (M203.0025: Mê gôm mét KV3) - Link 'KH_02_DonGia_CaMay'!N683"),
        "M_BD.41131": ("='KH_02_DonGia_CaMay'!N685", "QĐ Ca máy Khánh Hòa (M203.0027: TB tạo dòng điện KV3) - Link 'KH_02_DonGia_CaMay'!N685"),
        "M_E6C.3.03_MEG": ("='KH_02_DonGia_CaMay'!N683", "QĐ Ca máy Khánh Hòa (M203.0025: Mê gôm mét KV3) - Link 'KH_02_DonGia_CaMay'!N683"),
        "M_BD.41121": ("='KH_02_DonGia_CaMay'!N684", "QĐ Ca máy Khánh Hòa (M203.0026: TB kiểm tra áp lực KV3) - Link 'KH_02_DonGia_CaMay'!N684"),
    }
    
    ld_cm_map = {
        "M_E6C.3.05_VOM": ("='02_DonGia_CaMay'!O677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng VÙNG IV) - Link '02_DonGia_CaMay'!O677"),
        "M_E6C.3.05_BUI": ("='02_DonGia_CaMay'!O459", "QĐ 21/QĐ-SXD Lâm Đồng (M112.4301: Máy cầm tay VÙNG IV) - Link '02_DonGia_CaMay'!O459"),
        "M_E6C.3.02_SOLO": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra thử nghiệm VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
        "M_E6C.3.02_VOM": ("='02_DonGia_CaMay'!O677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng VÙNG IV) - Link '02_DonGia_CaMay'!O677"),
        "M_BD.41111": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
        "M_E6C.3.01_MEG": ("='02_DonGia_CaMay'!O683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Mê gôm mét VÙNG IV) - Link '02_DonGia_CaMay'!O683"),
        "M_BD.41131": ("='02_DonGia_CaMay'!O685", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0027: TB tạo dòng điện VÙNG IV) - Link '02_DonGia_CaMay'!O685"),
        "M_E6C.3.03_MEG": ("='02_DonGia_CaMay'!O683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Mê gôm mét VÙNG IV) - Link '02_DonGia_CaMay'!O683"),
        "M_BD.41121": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
        "M_E6C.1.23_BOM": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
        "M_E6C.1.31_GAS": ("='02_DonGia_CaMay'!O664", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0006: Hợp bộ phân tích hàm lượng khí VÙNG IV) - Link '02_DonGia_CaMay'!O664"),
        "M_E6C.2.04_VOM": ("='02_DonGia_CaMay'!O677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng VÙNG IV) - Link '02_DonGia_CaMay'!O677"),
        "M_E6C.1.18_VOM": ("='02_DonGia_CaMay'!O677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng VÙNG IV) - Link '02_DonGia_CaMay'!O677"),
        "M_E6C.3.04_SOLO": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra thử nghiệm VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
        "DC_E6C.3.04_S200": ("='02_DonGia_CaMay'!O683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Dụng cụ thử nghiệm VÙNG IV) - Link '02_DonGia_CaMay'!O683"),
        "M_E6B.1.13.01_TR": ("='02_DonGia_CaMay'!O672", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0014: Máy đo điện trở một chiều VÙNG IV) - Link '02_DonGia_CaMay'!O672"),
        "M_E6C.3.06_CD": ("='02_DonGia_CaMay'!O683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Mê gôm mét đo điện trở CĐ VÙNG IV) - Link '02_DonGia_CaMay'!O683"),
        "M_E6C.3.06_BAN": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
        "M_DB.24001": ("='02_DonGia_CaMay'!O673", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0015: Máy đo điện trở tiếp địa VÙNG IV) - Link '02_DonGia_CaMay'!O673"),
        "M_E6BOP.1.03": ("='02_DonGia_CaMay'!O148", "QĐ 21/QĐ-SXD Lâm Đồng (M102.1308: Hệ kích nâng / Palăng VÙNG IV) - Link '02_DonGia_CaMay'!O148"),
        "M_E6BOP.1.01": ("='02_DonGia_CaMay'!O684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực VÙNG IV) - Link '02_DonGia_CaMay'!O684"),
    }
    
    write_r = 6
    stt_master_counter = 1
    
    for b in blocks:
        stt_old = b['stt']
        code = b['code']
        base_name = b['name']
        unit = b['unit']
        vt_details = b['vt_details']
        pt_details = b['pt_details']
        
        # 1. Master row for Vĩnh Tân 4
        if vt_details:
            r_master_vt = write_r
            clean_name = base_name.replace('[LOẠI ', '').split(']', 1)[-1].strip()
            title_vt = f"[LOẠI {stt_master_counter}.1 - VĨNH TÂN 4] {clean_name}"
            loc_vt = f"Nhà máy chính NMNĐ Vĩnh Tân 4 ({len(vt_details)} công tác - Đơn giá Lâm Đồng Vùng IV)"
            
            old_ws.row_dimensions[write_r].height = 24
            old_ws.cell(write_r, 1, f"{stt_old}.1").alignment = align_center
            old_ws.cell(write_r, 2, code).alignment = align_center
            old_ws.cell(write_r, 3, title_vt).alignment = align_left
            old_ws.cell(write_r, 4, loc_vt).alignment = align_left
            old_ws.cell(write_r, 5, unit).alignment = align_center
            
            # Start and End detail rows for VT
            start_det_vt = write_r + 1
            end_det_vt = write_r + len(vt_details)
            old_ws.cell(write_r, 6, f"=SUM(F{start_det_vt}:F{end_det_vt})").alignment = align_right
            old_ws.cell(write_r, 6).number_format = '#,##0.000' if is_nc else '#,##0.00'
            
            # Unit Price & Formula
            price_map = ld_nc_map if is_nc else ld_cm_map
            p_formula, p_note = price_map.get(code, ("0", "Chưa xác định"))
            old_ws.cell(write_r, 7, p_formula).alignment = align_right
            old_ws.cell(write_r, 7).number_format = '#,##0'
            
            # Total
            old_ws.cell(write_r, 8, f"=ROUND(F{r_master_vt}*G{r_master_vt}, 0)").alignment = align_right
            old_ws.cell(write_r, 8).number_format = '#,##0'
            old_ws.cell(write_r, 9, p_note).alignment = align_left
            
            for c_i in range(1, 10):
                c_cell = old_ws.cell(write_r, c_i)
                c_cell.font = font_bold
                c_cell.fill = fill_accent
                c_cell.border = border_thin
                
            write_r += 1
            
            # Write VT detail rows
            for det in vt_details:
                d_stt, d_code, d_name, d_loc, d_unit, d_formula, _, _, _ = det
                old_ws.row_dimensions[write_r].height = 20
                old_ws.cell(write_r, 1, d_stt).alignment = align_center
                old_ws.cell(write_r, 2, d_code).alignment = align_center
                old_ws.cell(write_r, 3, d_name).alignment = align_left
                old_ws.cell(write_r, 4, d_loc).alignment = align_left
                old_ws.cell(write_r, 5, d_unit).alignment = align_center
                
                c_f = old_ws.cell(write_r, 6, d_formula)
                c_f.alignment = align_right
                c_f.number_format = '#,##0.000' if is_nc else '#,##0.00'
                
                old_ws.cell(write_r, 7, "").alignment = align_right
                old_ws.cell(write_r, 8, "").alignment = align_right
                old_ws.cell(write_r, 9, "").alignment = align_left
                
                for c_i in range(1, 10):
                    c_cell = old_ws.cell(write_r, c_i)
                    c_cell.font = font_regular
                    c_cell.border = border_thin
                write_r += 1
                
        # 2. Master row for Phước Thái 1, 2, 3
        if pt_details:
            r_master_pt = write_r
            clean_name = base_name.replace('[LOẠI ', '').split(']', 1)[-1].strip()
            sub_id = f"{stt_old}.2" if vt_details else f"{stt_old}"
            title_pt = f"[LOẠI {sub_id} - PHƯỚC THÁI] {clean_name}"
            loc_pt = f"NMĐMT Phước Thái 1, 2, 3 ({len(pt_details)} công tác - Đơn giá Khánh Hòa KV3 Ninh Phước)"
            
            old_ws.row_dimensions[write_r].height = 24
            old_ws.cell(write_r, 1, sub_id).alignment = align_center
            old_ws.cell(write_r, 2, code).alignment = align_center
            old_ws.cell(write_r, 3, title_pt).alignment = align_left
            old_ws.cell(write_r, 4, loc_pt).alignment = align_left
            old_ws.cell(write_r, 5, unit).alignment = align_center
            
            start_det_pt = write_r + 1
            end_det_pt = write_r + len(pt_details)
            old_ws.cell(write_r, 6, f"=SUM(F{start_det_pt}:F{end_det_pt})").alignment = align_right
            old_ws.cell(write_r, 6).number_format = '#,##0.000' if is_nc else '#,##0.00'
            
            price_map = kh_nc_map if is_nc else kh_cm_map
            p_formula, p_note = price_map.get(code, ("0", "Chưa xác định"))
            old_ws.cell(write_r, 7, p_formula).alignment = align_right
            old_ws.cell(write_r, 7).number_format = '#,##0'
            
            old_ws.cell(write_r, 8, f"=ROUND(F{r_master_pt}*G{r_master_pt}, 0)").alignment = align_right
            old_ws.cell(write_r, 8).number_format = '#,##0'
            old_ws.cell(write_r, 9, p_note).alignment = align_left
            
            for c_i in range(1, 10):
                c_cell = old_ws.cell(write_r, c_i)
                c_cell.font = font_bold
                c_cell.fill = fill_accent_pt # Soft peach highlight for Phuoc Thai
                c_cell.border = border_thin
                
            write_r += 1
            
            # Write PT detail rows
            for det in pt_details:
                d_stt, d_code, d_name, d_loc, d_unit, d_formula, _, _, _ = det
                old_ws.row_dimensions[write_r].height = 20
                old_ws.cell(write_r, 1, d_stt).alignment = align_center
                old_ws.cell(write_r, 2, d_code).alignment = align_center
                old_ws.cell(write_r, 3, d_name).alignment = align_left
                old_ws.cell(write_r, 4, d_loc).alignment = align_left
                old_ws.cell(write_r, 5, d_unit).alignment = align_center
                
                c_f = old_ws.cell(write_r, 6, d_formula)
                c_f.alignment = align_right
                c_f.number_format = '#,##0.000' if is_nc else '#,##0.00'
                
                old_ws.cell(write_r, 7, "").alignment = align_right
                old_ws.cell(write_r, 8, "").alignment = align_right
                old_ws.cell(write_r, 9, "").alignment = align_left
                
                for c_i in range(1, 10):
                    c_cell = old_ws.cell(write_r, c_i)
                    c_cell.font = font_regular
                    c_cell.border = border_thin
                write_r += 1
                
        stt_master_counter += 1
        
    print(f"    Rebuilt '{sheet_name}' successfully: now has {write_r - 1} rows.")

# Rebuild both sheets
rebuild_resource_sheet("Nhân công", is_nc=True)
rebuild_resource_sheet("Máy thi công", is_nc=False)

print(">>> [6/6] Đang lưu file và xác thực liên kết...")
wb.save(PROJECT_PATH)
file_size_mb = os.path.getsize(PROJECT_PATH) / (1024 * 1024)
print(f"    [OK] Đã lưu file: {PROJECT_PATH} ({file_size_mb:.2f} MB)")

try:
    wb.save(DOWNLOADS_PATH)
    print(f"    [OK] Đã lưu bản sao sang: {DOWNLOADS_PATH}")
except PermissionError:
    alt_downloads = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"
    wb.save(alt_downloads)
    print(f"    [LƯU Ý] File tại Downloads đang được mở trong Excel, đã lưu bản sao sang: {alt_downloads}")

print("Hoàn tất 100% việc áp dụng Đơn giá Khánh Hòa cho Phước Thái 1, 2, 3!")
