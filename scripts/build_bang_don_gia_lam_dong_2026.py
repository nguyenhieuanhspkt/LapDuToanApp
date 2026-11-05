# -*- coding: utf-8 -*-
"""
Script build_bang_don_gia_lam_dong_2026.py
Tổng hợp dữ liệu Đơn giá Nhân công, Đơn giá Ca máy, Đơn giá Công bố Vật liệu tỉnh Lâm Đồng năm 2026.
Nguồn dữ liệu:
1. Nhân công: Văn bản số 5898/SXD-QLXD ngày 02/7/2026 của Sở Xây dựng Lâm Đồng
2. Ca máy: Quyết định số 21/QĐ-SXD ngày 14/01/2026 của Sở Xây dựng Lâm Đồng
3. Vật liệu: Công bố giá VLXD Tháng 8/2026 (Gia_VLXD_LamDong_T08_2026.xlsx)
"""

import os
import re
import sys
import io
import pdfplumber
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

DIR_LAM_DONG = r"D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Năm 2026\Thẩm định 360_hieuna\Định mức QĐ\Đơn giá lâm đồng"
PDF_CA_MAY = os.path.join(DIR_LAM_DONG, "Ca máy", "ca máy 21 LĐ.pdf")
EXCEL_VLXD = os.path.join(DIR_LAM_DONG, "Vật liệu", "Gia_VLXD_LamDong_T08_2026.xlsx")

OUTPUT_PATHS = [
    os.path.join(DIR_LAM_DONG, "BANG_DON_GIA_LAM_DONG_2026.xlsx"),
    r"d:\TaskApp_kiet\LapDuToanApp\BANG_DON_GIA_LAM_DONG_2026.xlsx",
    r"C:\Users\Admin\Downloads\BANG_DON_GIA_LAM_DONG_2026.xlsx"
]

def clean_num(s):
    if not s or str(s).strip() in ['-', '', 'None', '—']:
        return None
    s = re.sub(r'\s+', '', str(s)).replace('.', '').replace(',', '.')
    try:
        val = float(s)
        return int(val) if val.is_integer() else val
    except:
        return s

print(">>> [1/4] Đang trích xuất dữ liệu Ca máy từ PDF QĐ 21/QĐ-SXD...")
ca_may_records = []
current_chapter = ""
current_group = ""

with pdfplumber.open(PDF_CA_MAY) as pdf:
    for pno in range(9, len(pdf.pages)):
        tables = pdf.pages[pno].extract_tables()
        for t in tables:
            for r in t:
                if not r or len(r) < 3:
                    continue
                c0 = str(r[0] or '').strip().replace('\n', ' ')
                c1 = str(r[1] or '').strip().replace('\n', ' ')
                c2 = str(r[2] or '').strip().replace('\n', ' ')
                
                # Check for chapter / group header
                if 'CHƯƠNG' in c1 or 'CHƢƠNG' in c1 or 'CHƯƠNG' in c0 or 'CHƢƠNG' in c0:
                    current_chapter = f"{c1}: {c2}".strip(': ')
                    continue
                if re.match(r'^M\d{3}\.0000$', c1):
                    current_group = f"{c1}: {c2}".strip(': ')
                    continue
                if re.match(r'^M\d{3}\.\d{2}00$', c1) or (not c0.isdigit() and c1.startswith('M')):
                    current_group = f"{c1}: {c2}".strip(': ')
                    continue
                
                # Machine row: c0 is digit and c1 starts with M
                if re.match(r'^M\d{3}\.\d{4}[a-zA-Z]?$', c1) and c0.isdigit():
                    spec = c2
                    group_prefix = current_group.split(':', 1)[-1].strip() if ':' in current_group else current_group
                    group_clean = re.sub(r'-\s*(công suất|dung tích|trọng lượng|tải trọng|sức nâng|đường kính|bề rộng).*$', '', group_prefix, flags=re.IGNORECASE).strip(' :')
                    full_name = f"{group_prefix} {spec}".strip() if group_prefix and group_prefix not in spec else spec
                    
                    row_data = {
                        'stt': int(c0),
                        'code': c1,
                        'chapter': current_chapter,
                        'group': group_clean,
                        'spec': spec,
                        'full_name': full_name,
                        'fuel_norm': clean_num(r[3]) if len(r) > 3 else None,
                        'fuel_type': str(r[4] or '').strip().replace('\n', ' ') if len(r) > 4 else '',
                        'fuel_cost': clean_num(r[5]) if len(r) > 5 else None,
                        'driver_rank': str(r[6] or '').strip().replace('\n', ' ') if len(r) > 6 else '',
                        'driver_wage_v2': clean_num(r[7]) if len(r) > 7 else None,
                        'driver_wage_v3': clean_num(r[8]) if len(r) > 8 else None,
                        'driver_wage_v4': clean_num(r[9]) if len(r) > 9 else None,
                        'price_v2': clean_num(r[10]) if len(r) > 10 else None,
                        'price_v3': clean_num(r[11]) if len(r) > 11 else None,
                        'price_v4': clean_num(r[12]) if len(r) > 12 else None,
                    }
                    ca_may_records.append(row_data)

print(f"    => Đã trích xuất thành công {len(ca_may_records)} máy thi công.")

print(">>> [2/4] Đang khởi tạo Workbook mới và thiết lập các Sheet...")
wb = openpyxl.Workbook()
# Remove default sheet
wb.remove(wb.active)

# Color Palette & Styles
font_title = Font(name="Times New Roman", size=15, bold=True, color="1B365D")
font_sub = Font(name="Times New Roman", size=11, italic=True, color="333333")
font_sec = Font(name="Times New Roman", size=12, bold=True, color="0D47A1")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Times New Roman", size=10, bold=True)
font_regular = Font(name="Times New Roman", size=10)
font_italic = Font(name="Times New Roman", size=10, italic=True)

fill_primary = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")  # Dark Blue
fill_teal = PatternFill(start_color="005F73", end_color="005F73", fill_type="solid")     # Teal Blue
fill_accent = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")   # Light soft blue
fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")    # Very light grey
fill_highlight = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")# Light gold

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
border_total = Border(
    top=Side(style='thin', color='1F497D'),
    bottom=Side(style='double', color='1F497D')
)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# -------------------------------------------------------------
# SHEET 0: THUYẾT MINH & CĂN CỨ PHÁP LÝ
# -------------------------------------------------------------
print("    -> Đang tạo Sheet 00_ThuyetMinh_CanCu...")
ws0 = wb.create_sheet(title="00_ThuyetMinh_CanCu")
ws0.views.sheetView[0].showGridLines = True

ws0.merge_cells("A1:G1")
ws0["A1"] = "BẢNG ĐƠN GIÁ NHÂN CÔNG - GIÁ CA MÁY - GIÁ VẬT LIỆU XÂY DỰNG TỈNH LÂM ĐỒNG"
ws0["A1"].font = font_title
ws0["A1"].alignment = align_center
ws0.row_dimensions[1].height = 28

ws0.merge_cells("A2:G2")
ws0["A2"] = "Hệ thống dữ liệu thẩm định và lập dự toán chi phí đầu tư xây dựng trên địa bàn tỉnh Lâm Đồng năm 2026"
ws0["A2"].font = font_sub
ws0["A2"].alignment = align_center
ws0.row_dimensions[2].height = 20

intro_data = [
    ("I. CĂN CỨ PHÁP LÝ VÀ NGUỒN SỐ LIỆU", "", font_sec, fill_accent),
    ("1. Đơn giá Nhân công", "Văn bản số 5898/SXD-QLXD ngày 02/7/2026 của Sở Xây dựng Lâm Đồng về việc hướng dẫn một số nội dung về lập và quản lý chi phí ĐTXD theo Nghị định số 206/2026/NĐ-CP của Chính phủ; công bố bảng giá nhân công chuyển đổi theo nhóm nhân công quy định tại Thông tư 37/2026/TT-BXD và Thông tư 38/2026/TT-BXD của Bộ Xây dựng.", font_regular, None),
    ("2. Bảng giá Ca máy & TBTC", "Quyết định số 21/QĐ-SXD ngày 14/01/2026 của Giám đốc Sở Xây dựng tỉnh Lâm Đồng ban hành Bảng giá ca máy và thiết bị thi công xây dựng trên địa bàn tỉnh Lâm Đồng năm 2026.", font_regular, None),
    ("3. Công bố giá Vật liệu XD", "Công bố giá vật liệu xây dựng Tháng 08 năm 2026 trên địa bàn tỉnh Lâm Đồng của Liên Sở Xây dựng - Tài chính (File dữ liệu chi tiết Gia_VLXD_LamDong_T08_2026.xlsx đã số hóa chuẩn gồm 17 nhóm vật liệu, công bố giá toàn tỉnh tại 3 khu vực, giá tại mỏ/nơi sản xuất và giá tại nơi bán).", font_regular, None),
    ("II. PHÂN VÙNG ĐỊA BÀN HÀNH CHÍNH ÁP DỤNG", "", font_sec, fill_accent),
    ("Vùng II", "Thành phố Đà Lạt, Thành phố Bảo Lộc.", font_regular, None),
    ("Vùng III", "Huyện Đức Trọng, Huyện Di Linh.", font_regular, None),
    ("Vùng IV", "Gồm các huyện: Lạc Dương, Đam Rông, Lâm Hà, Bảo Lâm, Cát Tiên, Đạ Huoai, Đạ Tẻh.", font_regular, None),
    ("Khu vực Đặc khu", "Đơn giá nhân công được điều chỉnh với hệ số 1,08 so với đơn giá nhân công của vùng tương ứng theo ghi chú tại Văn bản số 5898/SXD-QLXD.", font_regular, None),
    ("III. CẤU TRÚC WORKBOOK TRA CỨU", "", font_sec, fill_accent),
    ("Sheet 01_DonGia_NhanCong", "Gồm Bảng đơn giá chuẩn 13 nhóm công bố tại VB 5898/SXD-QLXD (Bậc 3,5/7) và Bảng thang lương đầy đủ từ Bậc 1/7 đến Bậc 7/7 tính toán chuyển đổi theo Thông tư 38/2026/TT-BXD cho cả 3 vùng và Khu vực đặc khu.", font_regular, None),
    ("Sheet 02_DonGia_CaMay", "Gồm toàn bộ 674 máy và thiết bị thi công theo QĐ 21/QĐ-SXD, thể hiện định mức tiêu hao nhiên liệu, thợ lái máy, chi phí nhiên liệu, tiền lương thợ lái và giá ca máy hoàn chỉnh cho Vùng II, III, IV.", font_regular, None),
    ("Sheet 03_VLXD_TongHop_17Nhom", "Bảng tổng hợp phạm vi 17 nhóm vật tư vật liệu chủ yếu công bố trong tháng 8/2026.", font_regular, None),
    ("Sheet 04_VLXD_ChiTiet_ToanTinh", "Bảng giá chi tiết 1.521 danh mục vật liệu toàn tỉnh tại 3 khu vực theo Phụ lục 1 Công bố giá VLXD.", font_regular, None),
    ("Sheet 05_VLXD_GiaTaiMo_NoiSX", "Bảng giá đá xây dựng, cát xây dựng khai thác tại mỏ, trạm nghiền kèm thông tin giấy phép và tọa độ GPS mỏ theo Phụ lục 3.", font_regular, None),
    ("Sheet 06_VLXD_GiaTaiNoiBan", "Bảng giá khảo sát tại các đại lý, cơ sở kinh doanh, trạm trộn theo địa bàn theo Phụ lục 4.", font_regular, None),
]

r_idx = 4
for item in intro_data:
    title, desc, f_style, fill_s = item
    if not desc:
        ws0.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=7)
        c = ws0.cell(r_idx, 1, title)
        c.font = f_style
        if fill_s: c.fill = fill_s
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        ws0.row_dimensions[r_idx].height = 24
    else:
        c1 = ws0.cell(r_idx, 1, title)
        c1.font = font_bold
        c1.alignment = Alignment(horizontal='left', vertical='top')
        ws0.merge_cells(start_row=r_idx, start_column=2, end_row=r_idx, end_column=7)
        c2 = ws0.cell(r_idx, 2, desc)
        c2.font = font_regular
        c2.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws0.row_dimensions[r_idx].height = 36 if len(desc) > 100 else 22
    r_idx += 1

ws0.column_dimensions['A'].width = 28
for col_letter in ['B', 'C', 'D', 'E', 'F', 'G']:
    ws0.column_dimensions[col_letter].width = 18

# -------------------------------------------------------------
# SHEET 1: ĐƠN GIÁ NHÂN CÔNG
# -------------------------------------------------------------
print("    -> Đang tạo Sheet 01_DonGia_NhanCong...")
ws1 = wb.create_sheet(title="01_DonGia_NhanCong")
ws1.views.sheetView[0].showGridLines = True

ws1.merge_cells("A1:G1")
ws1["A1"] = "BẢNG ĐƠN GIÁ NHÂN CÔNG XÂY DỰNG TỈNH LÂM ĐỒNG NĂM 2026"
ws1["A1"].font = font_title
ws1["A1"].alignment = align_center

ws1.merge_cells("A2:G2")
ws1["A2"] = "(Ban hành kèm theo Văn bản số 5898/SXD-QLXD ngày 02/7/2026 của Sở Xây dựng Lâm Đồng)"
ws1["A2"].font = font_sub
ws1["A2"].alignment = align_center

# Table 1: Bảng công bố chuẩn VB 5898 (Bậc bình quân 3,5/7)
ws1.merge_cells("A4:G4")
ws1["A4"] = "PHẦN A: BẢNG GIÁ NHÂN CÔNG CHUYỂN ĐỔI THEO NHÓM THEO VB SỐ 5898/SXD-QLXD (BẬC BÌNH QUÂN 3,5/7)"
ws1["A4"].font = font_sec
ws1["A4"].fill = fill_accent
ws1["A4"].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws1.row_dimensions[4].height = 24

headers_nc1 = ["STT", "Nhóm nhân công", "Tên nhân công (Loại công tác áp dụng)", "Đơn vị tính", "Đơn giá Vùng II (đồng)", "Đơn giá Vùng III (đồng)", "Đơn giá Vùng IV (đồng)"]
ws1.row_dimensions[5].height = 28
for c_idx, h in enumerate(headers_nc1, 1):
    cell = ws1.cell(5, c_idx, h)
    cell.font = font_th
    cell.fill = fill_primary
    cell.alignment = align_center
    cell.border = border_header

nc1_data = [
    # Group I
    ("I", "NHÓM NHÂN CÔNG XÂY DỰNG", "", "", None, None, None, True),
    ("1", "Nhóm 1", "Chuẩn bị mặt bằng, phá dỡ - tháo dỡ, vệ sinh; công tác phục vụ thi công (bốc xếp, vận chuyển, đóng bao, giàn giáo); công tác đất, đá, cát (đào, đắp, bơm, nạo vét, xói hút, vận chuyển và không gồm phá đá bằng nổ mìn).", "công", 276204, 275554, 268124, False),
    ("2", "Nhóm 2", "Thi công cọc, gia cố, khoan tạo lỗ, khoan giảm áp, khớp nối, phá đá bằng khoan nổ mìn, bê tông xi măng, lắp đặt cấu kiện bê tông, ván khuôn, cốt thép, bê tông nhựa, xây gạch, đá và các công tác khác chưa được quy định tại nhóm 1, 3 và 4 bảng này.", "công", 306300, 300500, 291700, False),
    ("3", "Nhóm 3", "Lắp đặt hệ thống kỹ thuật công trình; đường ray, gia công lắp dựng kết cấu gỗ, kết cấu sắt thép; cáp thép; trát, ốp, lát, láng, làm trần, mái, sơn, bả; khảo sát xây dựng.", "công", 328131, 321917, 312490, False),
    ("4", "Nhóm 4", "Lắp đặt, tháo dỡ đường dây tải điện và trạm biến áp, thiết bị công nghệ, thông tin, tín hiệu, đo lường, điều khiển, thí nghiệm.", "công", 344746, 340813, 328116, False),
    # Group II
    ("II", "NHÓM NHÂN CÔNG VẬN HÀNH, ĐIỀU KHIỂN MÁY", "", "", None, None, None, True),
    ("1", "Lái máy", "Nhân công vận hành, điều khiển máy thi công xây dựng", "công", 338057, 336191, 325325, False),
    ("2", "Lái xe", "Lái xe các loại", "công", 317136, 315385, 305192, False),
    ("3", "Thủy thủ / Thợ điện máy", "Thủy thủ, thợ máy, thợ điện", "công", 365000, 342000, 326200, False),
    ("4", "Tàu sông", "Máy trưởng, máy I, máy II, điện trưởng, kỹ thuật viên cuốc I, kỹ thuật viên cuốc II tàu sông", "công", 358400, 351200, 341900, False),
    ("5", "Tàu biển", "Máy trưởng, máy I, máy II, điện trưởng, kỹ thuật viên cuốc I, kỹ thuật viên cuốc II tàu biển", "công", 362900, 356300, 346400, False),
    ("6", "Thuyền trưởng", "Thuyền trưởng, thuyền phó", "công", 400600, 387400, 377700, False),
    # Group III
    ("III", "NHÓM NHÂN CÔNG THỰC HIỆN CÔNG TÁC KHÁC", "", "", None, None, None, True),
    ("1", "Kỹ sư KS-TN", "Kỹ sư thực hiện khảo sát, thí nghiệm", "công", 320000, 316900, 313000, False),
    ("2", "Thợ lặn", "Thợ lặn", "công", 579400, 561400, 548900, False),
    ("3", "Nghệ nhân", "Nghệ nhân", "công", 568000, 527000, 502000, False),
]

curr_r = 6
for row in nc1_data:
    stt, grp, desc, unit, v2, v3, v4, is_sec = row
    ws1.row_dimensions[curr_r].height = 24 if is_sec else (42 if len(desc) > 80 else 24)
    if is_sec:
        ws1.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=7)
        c = ws1.cell(curr_r, 1, f"{stt}. {grp}")
        c.font = font_bold
        c.fill = fill_zebra
        c.alignment = Alignment(horizontal='left', vertical='center', indent=1)
        for c_idx in range(1, 8):
            ws1.cell(curr_r, c_idx).border = border_thin
    else:
        ws1.cell(curr_r, 1, stt).alignment = align_center
        ws1.cell(curr_r, 2, grp).alignment = align_left
        ws1.cell(curr_r, 3, desc).alignment = align_left
        ws1.cell(curr_r, 4, unit).alignment = align_center
        
        c_v2 = ws1.cell(curr_r, 5, v2)
        c_v2.number_format = '#,##0'
        c_v2.alignment = align_right
        
        c_v3 = ws1.cell(curr_r, 6, v3)
        c_v3.number_format = '#,##0'
        c_v3.alignment = align_right
        
        c_v4 = ws1.cell(curr_r, 7, v4)
        c_v4.number_format = '#,##0'
        c_v4.alignment = align_right
        
        for c_idx in range(1, 8):
            cell = ws1.cell(curr_r, c_idx)
            cell.font = font_regular
            cell.border = border_thin
    curr_r += 1

# Table 2: Detailed Thang bảng lương Bậc 1/7 đến Bậc 7/7 per Thông tư 38/2026
curr_r += 2
ws1.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=10)
c_hdr2 = ws1.cell(curr_r, 1, "PHẦN B: BẢNG QUY ĐỔI ĐƠN GIÁ NHÂN CÔNG CHI TIẾT TỪ BẬC 1/7 ĐẾN BẬC 7/7 THEO THÔNG TƯ 38/2026/TT-BXD")
c_hdr2.font = font_sec
c_hdr2.fill = fill_accent
c_hdr2.alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws1.row_dimensions[curr_r].height = 24
curr_r += 1

headers_nc2 = ["STT", "Nhóm nhân công", "Cấp bậc thợ", "Hệ số cấp bậc", "Hệ số quy đổi (so Bậc 3,5)", "Vùng II (đồng/công)", "Vùng III (đồng/công)", "Vùng IV (đồng/công)", "Đặc khu Vùng II (x1,08)", "Đặc khu Vùng III (x1,08)"]
ws1.row_dimensions[curr_r].height = 28
for c_idx, h in enumerate(headers_nc2, 1):
    cell = ws1.cell(curr_r, c_idx, h)
    cell.font = font_th
    cell.fill = fill_teal
    cell.alignment = align_center
    cell.border = border_header
curr_r += 1

# Standard grade scale for 7 grades:
# Bac 1: 1.00; Bac 2: 1.18; Bac 3: 1.39; Bac 3.5: 1.52; Bac 4: 1.65; Bac 5: 1.95; Bac 6: 2.31; Bac 7: 2.73
grades = [
    ("Bậc 1,0/7", 1.00, 1.00/1.52),
    ("Bậc 2,0/7", 1.18, 1.18/1.52),
    ("Bậc 3,0/7", 1.39, 1.39/1.52),
    ("Bậc 3,5/7 (Chuẩn)", 1.52, 1.52/1.52),
    ("Bậc 4,0/7", 1.65, 1.65/1.52),
    ("Bậc 5,0/7", 1.95, 1.95/1.52),
    ("Bậc 6,0/7", 2.31, 2.31/1.52),
    ("Bậc 7,0/7", 2.73, 2.73/1.52),
]

target_groups = [
    ("Nhóm 1", 276204, 275554, 268124),
    ("Nhóm 2", 306300, 300500, 291700),
    ("Nhóm 3", 328131, 321917, 312490),
    ("Nhóm 4", 344746, 340813, 328116),
    ("Lái máy", 338057, 336191, 325325),
    ("Lái xe", 317136, 315385, 305192),
]

stt_counter = 1
for grp_name, base_v2, base_v3, base_v4 in target_groups:
    for g_title, g_coef, g_ratio in grades:
        is_standard = "Chuẩn" in g_title
        ws1.cell(curr_r, 1, stt_counter).alignment = align_center
        ws1.cell(curr_r, 2, grp_name).alignment = align_left
        ws1.cell(curr_r, 3, g_title).alignment = align_center
        
        c_k = ws1.cell(curr_r, 4, g_coef)
        c_k.number_format = '0.00'
        c_k.alignment = align_center
        
        c_r = ws1.cell(curr_r, 5, g_ratio)
        c_r.number_format = '0.00000'
        c_r.alignment = align_center
        
        # Prices
        val_v2 = round(base_v2 * g_ratio)
        val_v3 = round(base_v3 * g_ratio)
        val_v4 = round(base_v4 * g_ratio)
        val_dk_v2 = round(val_v2 * 1.08)
        val_dk_v3 = round(val_v3 * 1.08)
        
        for c_idx, val in enumerate([val_v2, val_v3, val_v4, val_dk_v2, val_dk_v3], 6):
            c_val = ws1.cell(curr_r, c_idx, val)
            c_val.number_format = '#,##0'
            c_val.alignment = align_right
            
        for c_idx in range(1, 11):
            cell = ws1.cell(curr_r, c_idx)
            cell.font = font_bold if is_standard else font_regular
            cell.border = border_thin
            if is_standard:
                cell.fill = fill_highlight
                
        ws1.row_dimensions[curr_r].height = 20
        curr_r += 1
        stt_counter += 1

# Column widths for Sheet 1
ws1.column_dimensions['A'].width = 8
ws1.column_dimensions['B'].width = 18
ws1.column_dimensions['C'].width = 45
ws1.column_dimensions['D'].width = 12
ws1.column_dimensions['E'].width = 18
ws1.column_dimensions['F'].width = 18
ws1.column_dimensions['G'].width = 18
ws1.column_dimensions['H'].width = 18
ws1.column_dimensions['I'].width = 20
ws1.column_dimensions['J'].width = 20
ws1.freeze_panes = "A6"

# -------------------------------------------------------------
# SHEET 2: ĐƠN GIÁ CA MÁY VÀ THIẾT BỊ THI CÔNG
# -------------------------------------------------------------
print("    -> Đang tạo Sheet 02_DonGia_CaMay...")
ws2 = wb.create_sheet(title="02_DonGia_CaMay")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:N1")
ws2["A1"] = "BẢNG GIÁ CA MÁY VÀ THIẾT BỊ THI CÔNG XÂY DỰNG TỈNH LÂM ĐỒNG NĂM 2026"
ws2["A1"].font = font_title
ws2["A1"].alignment = align_center

ws2.merge_cells("A2:N2")
ws2["A2"] = "(Ban hành kèm theo Quyết định số 21/QĐ-SXD ngày 14/01/2026 của Sở Xây dựng Lâm Đồng)"
ws2["A2"].font = font_sub
ws2["A2"].alignment = align_center

headers_cm = [
    "STT", "Mã hiệu máy", "Chương thiết bị", "Nhóm máy", "Loại máy và quy cách chi tiết",
    "Định mức NL (1 ca)", "Loại nhiên liệu", "Chi phí NL (đồng)", "Cấp bậc thợ lái",
    "Lương thợ Vùng II (đồng)", "Lương thợ Vùng III (đồng)", "Lương thợ Vùng IV (đồng)",
    "GIÁ CA MÁY VÙNG II (đồng)", "GIÁ CA MÁY VÙNG III (đồng)", "GIÁ CA MÁY VÙNG IV (đồng)"
]

ws2.row_dimensions[4].height = 30
for c_idx, h in enumerate(headers_cm, 1):
    cell = ws2.cell(4, c_idx, h)
    cell.font = font_th
    cell.fill = fill_primary
    cell.alignment = align_center
    cell.border = border_header

for idx, rec in enumerate(ca_may_records, 1):
    r_num = idx + 4
    ws2.row_dimensions[r_num].height = 20
    
    ws2.cell(r_num, 1, rec['stt']).alignment = align_center
    ws2.cell(r_num, 2, rec['code']).alignment = align_center
    ws2.cell(r_num, 3, rec['chapter']).alignment = align_left
    ws2.cell(r_num, 4, rec['group']).alignment = align_left
    ws2.cell(r_num, 5, rec['full_name']).alignment = align_left
    
    # Fuel norm
    c_fn = ws2.cell(r_num, 6, rec['fuel_norm'])
    c_fn.alignment = align_right
    if isinstance(rec['fuel_norm'], (int, float)): c_fn.number_format = '#,##0.0' if isinstance(rec['fuel_norm'], float) else '#,##0'
    
    ws2.cell(r_num, 7, rec['fuel_type']).alignment = align_center
    
    c_fc = ws2.cell(r_num, 8, rec['fuel_cost'])
    c_fc.alignment = align_right
    if isinstance(rec['fuel_cost'], (int, float)): c_fc.number_format = '#,##0'
    
    ws2.cell(r_num, 9, rec['driver_rank']).alignment = align_center
    
    for c_i, key in enumerate(['driver_wage_v2', 'driver_wage_v3', 'driver_wage_v4', 'price_v2', 'price_v3', 'price_v4'], 10):
        val = rec[key]
        c = ws2.cell(r_num, c_i, val)
        c.alignment = align_right
        if isinstance(val, (int, float)):
            c.number_format = '#,##0'
            if c_i >= 13:
                c.font = font_bold
            else:
                c.font = font_regular
        else:
            c.font = font_regular
            
    for col_i in range(1, 16):
        cell = ws2.cell(r_num, col_i)
        if col_i < 13: cell.font = font_regular
        cell.border = border_thin
        if idx % 2 == 0:
            cell.fill = fill_zebra

ws2.column_dimensions['A'].width = 7
ws2.column_dimensions['B'].width = 14
ws2.column_dimensions['C'].width = 24
ws2.column_dimensions['D'].width = 26
ws2.column_dimensions['E'].width = 42
ws2.column_dimensions['F'].width = 12
ws2.column_dimensions['G'].width = 14
ws2.column_dimensions['H'].width = 16
ws2.column_dimensions['I'].width = 14
ws2.column_dimensions['J'].width = 16
ws2.column_dimensions['K'].width = 16
ws2.column_dimensions['L'].width = 16
ws2.column_dimensions['M'].width = 20
ws2.column_dimensions['N'].width = 20
ws2.column_dimensions['O'].width = 20
ws2.freeze_panes = "F5"
ws2.auto_filter.ref = f"A4:O{len(ca_may_records) + 4}"

# -------------------------------------------------------------
# SHEET 3, 4, 5, 6: VẬT LIỆU XÂY DỰNG TỪ FILE GỐC
# -------------------------------------------------------------
print(">>> [3/4] Đang nạp dữ liệu Công bố Giá Vật liệu Tháng 8/2026...")
wb_vl = openpyxl.load_workbook(EXCEL_VLXD, data_only=True)

vl_sheets_map = [
    ("00_TongHop_17Nhom", "03_VLXD_TongHop_17Nhom", "TỔNG HỢP 17 NHÓM VẬT LIỆU XÂY DỰNG CÔNG BỐ THÁNG 8/2026 - TỈNH LÂM ĐỒNG"),
    ("PL1_GiaToanTinh_3KV", "04_VLXD_ChiTiet_ToanTinh", "BẢNG GIÁ VẬT LIỆU XÂY DỰNG TOÀN TỈNH TẠI 3 KHU VỰC - THÁNG 8/2026"),
    ("PL3_GiaTaiMo", "05_VLXD_GiaTaiMo_NoiSX", "BẢNG GIÁ ĐÁ, CÁT KHAI THÁC TẠI MỎ VÀ NƠI SẢN XUẤT (KÈM GPS) - THÁNG 8/2026"),
    ("PL4_GiaTaiNoiBan", "06_VLXD_GiaTaiNoiBan", "BẢNG GIÁ VẬT LIỆU XÂY DỰNG KHẢO SÁT TẠI NƠI BÁN / ĐẠI LÝ - THÁNG 8/2026")
]

for src_name, dst_name, banner_title in vl_sheets_map:
    print(f"    -> Đang sao chép và định dạng Sheet {dst_name}...")
    src_ws = wb_vl[src_name]
    dst_ws = wb.create_sheet(title=dst_name)
    dst_ws.views.sheetView[0].showGridLines = True
    
    # Header Banner
    max_c = src_ws.max_column
    col_end_letter = get_column_letter(max_c)
    dst_ws.merge_cells(f"A1:{col_end_letter}1")
    dst_ws["A1"] = banner_title
    dst_ws["A1"].font = font_title
    dst_ws["A1"].alignment = align_center
    dst_ws.row_dimensions[1].height = 28
    
    dst_ws.merge_cells(f"A2:{col_end_letter}2")
    dst_ws["A2"] = "(Kèm theo Công bố giá VLXD Tháng 8/2026 của Liên Sở Xây dựng - Tài chính tỉnh Lâm Đồng)"
    dst_ws["A2"].font = font_sub
    dst_ws["A2"].alignment = align_center
    dst_ws.row_dimensions[2].height = 18
    
    # Copy data starting at row 4
    for r in range(1, src_ws.max_row + 1):
        dst_r = r + 3
        is_header_row = (r == 1)
        dst_ws.row_dimensions[dst_r].height = 26 if is_header_row else 20
        
        for c in range(1, max_c + 1):
            val = src_ws.cell(r, c).value
            target_cell = dst_ws.cell(dst_r, c, val)
            
            if is_header_row:
                target_cell.font = font_th
                target_cell.fill = fill_teal
                target_cell.alignment = align_center
                target_cell.border = border_header
            else:
                target_cell.border = border_thin
                # Number vs Text formatting
                if isinstance(val, (int, float)):
                    target_cell.font = font_regular
                    target_cell.number_format = '#,##0' if abs(val) >= 10 else '#,##0.00'
                    target_cell.alignment = align_right
                else:
                    target_cell.font = font_regular
                    sval = str(val or '').strip()
                    if sval.isdigit() and len(sval) <= 4:
                        target_cell.alignment = align_center
                    elif any(u in sval.lower() for u in ['kg', 'm3', 'viên', 'tấn', 'cây', 'cuộn', 'bộ', 'thùng', 'm2', 'md']):
                        target_cell.alignment = align_center
                    else:
                        target_cell.alignment = align_left
                        
                if r % 2 == 1 and not is_header_row:
                    target_cell.fill = fill_zebra
                    
    # Auto-fit column widths
    for col in dst_ws.iter_cols(min_row=4, max_row=min(src_ws.max_row + 3, 20), min_col=1, max_col=max_c):
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(cell.value or '')) for cell in col)
        dst_ws.column_dimensions[col_letter].width = max(12, min(max_len + 4, 45))
        
    dst_ws.freeze_panes = "D5" if max_c > 6 else "B5"
    dst_ws.auto_filter.ref = f"A4:{col_end_letter}{src_ws.max_row + 3}"

# -------------------------------------------------------------
# SAVE WORKBOOK TO ALL LOCATIONS
# -------------------------------------------------------------
print(">>> [4/4] Đang lưu file Excel và xác thực kiểm tra...")
for p in OUTPUT_PATHS:
    os.makedirs(os.path.dirname(p), exist_ok=True)
    wb.save(p)
    file_size_mb = os.path.getsize(p) / (1024 * 1024)
    print(f"    [OK] Đã lưu thành công: {p} ({file_size_mb:.2f} MB)")

print("\nHoàn tất 100% quá trình tạo Bảng Đơn giá Lâm Đồng năm 2026!")
