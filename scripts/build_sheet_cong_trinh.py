# -*- coding: utf-8 -*-
"""
Script build_sheet_cong_trinh.py
Xây dựng sheet "Công trình" (Bảng chi tiết khối lượng công tác xây dựng theo chuẩn F1 / Thông tư 36/2026/TT-BXD):
1. Nguồn dữ liệu: 856 dịch vụ PCCC từ sheet '02_Chi_Tiet_856_Dich_Vu'.
2. Cấu trúc sheet 'Công trình' (chuẩn mẫu Dự toán F1 theo ảnh đính kèm):
   - Tiêu đề: BẢNG CHI TIẾT KHỐI LƯỢNG CÔNG TÁC XÂY DỰNG
   - Cột A: STT (1 đến 856)
   - Cột B: Ký hiệu bản vẽ
   - Cột C: Mã hiệu công tác
   - Cột D: Danh mục công tác
   - Cột E: Đơn vị
   - Cột F: Khối lượng toàn bộ (Q)
   - Cột G, H, I: Tính trực tiếp (Đơn giá Vật liệu, Nhân công, Máy thi công)
   - Cột J, K, L: Thành tiền (Vật liệu, Nhân công, Máy thi công)
   - Cột M, N, O: Hệ số công tác (VL, NC, M)
   - Cột P: Đơn giá tổng hợp (=G+H+I)
   - Cột Q: Định mức (Căn cứ định mức)
   - Cột R: Tổng thành tiền (=J+K+L)
   - Cột S: Khu vực / Hạng mục công trình
   - Dòng THM - TỔNG HẠNG MỤC
   - Phần ký duyệt: Người thực hiện, Người chủ trì, Chứng chỉ hành nghề...
3. Liên kết ngược sang sheet 'hao phi VT':
   - Với mỗi STT trong 856 dịch vụ tại sheet 'hao phi VT':
     + Cột A (STT): link ='Công trình'!A...
     + Cột B (Mã hiệu): link ='Công trình'!C...
     + Cột D (Khu vực): link ='Công trình'!S...
     + Cột E (Tên công tác/công trình): link ='Công trình'!D...
     + Cột F (Đơn vị tính): link ='Công trình'!E...
     + Cột G (Khối lượng Q): link ='Công trình'!F...
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

print(">>> [1/5] Đang mở workbook và trích xuất dữ liệu 856 dịch vụ...")
wb = openpyxl.load_workbook(PROJECT_PATH)
ws_856 = wb['02_Chi_Tiet_856_Dich_Vu']

services = []
for row in ws_856.iter_rows(min_row=5, values_only=True):
    stt = row[0]
    loc = row[1]
    name = row[2]
    q = row[3]
    dvt = row[4]
    code = row[8] if len(row) > 8 else ''
    note = row[9] if len(row) > 9 else ''
    if stt is not None:
        services.append({
            'stt': int(stt),
            'loc': str(loc or '').strip(),
            'name': str(name or '').strip(),
            'q': q,
            'dvt': str(dvt or '').strip(),
            'code': str(code or '').strip(),
            'note': str(note or '').strip()
        })

print(f"    Đã nạp {len(services)} dịch vụ từ '02_Chi_Tiet_856_Dich_Vu'.")

# Xóa sheet 'Công trình' nếu đã tồn tại để tạo mới hoàn chỉnh
if "Công trình" in wb.sheetnames:
    wb.remove(wb["Công trình"])

# Tạo sheet 'Công trình' và đặt ở vị trí thứ 3 (sau 02_Chi_Tiet_856_Dich_Vu)
ws_ct = wb.create_sheet(title="Công trình", index=2)
ws_ct.views.sheetView[0].showGridLines = True

# Định dạng Styling
font_title = Font(name="Times New Roman", size=14, bold=True, color="1B365D")
font_sub = Font(name="Times New Roman", size=10, italic=True, color="333333")
font_proj = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
font_thm = Font(name="Times New Roman", size=11, bold=True, color="002060")
font_regular = Font(name="Times New Roman", size=10)
font_bold = Font(name="Times New Roman", size=10, bold=True)
font_italic = Font(name="Times New Roman", size=9, italic=True)

fill_th_main = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
fill_th_sub = PatternFill(start_color="2A5B8C", end_color="2A5B8C", fill_type="solid")
fill_thm = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

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
border_thm = Border(
    left=Side(style='thin', color='1F497D'),
    right=Side(style='thin', color='1F497D'),
    top=Side(style='medium', color='1F497D'),
    bottom=Side(style='double', color='1F497D')
)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

print(">>> [2/5] Đang thiết lập cấu trúc bảng Dự toán F1 trong sheet 'Công trình'...")

# Dòng 1: Tiêu đề bảng
ws_ct.merge_cells("A1:S1")
ws_ct["A1"] = "BẢNG CHI TIẾT KHỐI LƯỢNG CÔNG TÁC XÂY DỰNG"
ws_ct["A1"].font = font_title
ws_ct["A1"].alignment = align_center
ws_ct.row_dimensions[1].height = 26

# Dòng 2: Căn cứ thông tư
ws_ct.merge_cells("A2:S2")
ws_ct["A2"] = "(Ban hành kèm theo Thông tư số 36/2026/TT-BXD ngày 26/06/2026 của Bộ Xây dựng)"
ws_ct["A2"].font = font_sub
ws_ct["A2"].alignment = align_center
ws_ct.row_dimensions[2].height = 18

# Dòng 3: Tên công trình
ws_ct.merge_cells("A3:S3")
ws_ct["A3"] = "CÔNG TRÌNH: SỬA CHỮA LỚN HỆ THỐNG PCCC NĂM 2026 – CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4"
ws_ct["A3"].font = font_proj
ws_ct["A3"].alignment = align_center
ws_ct.row_dimensions[3].height = 20

# Dòng 4: Hạng mục
ws_ct.merge_cells("A4:S4")
ws_ct["A4"] = "HẠNG MỤC: TOÀN BỘ 856 DỊCH VỤ PCCC TẠI NMNĐ VĨNH TÂN 4 VÀ NMĐMT PHƯỚC THÁI 1, 2, 3"
ws_ct["A4"].font = font_proj
ws_ct["A4"].alignment = align_center
ws_ct.row_dimensions[4].height = 20

# Dòng 5 & 6: Header 2 tầng chuẩn F1
ws_ct.row_dimensions[5].height = 24
ws_ct.row_dimensions[6].height = 22

# Merge vertical headers:
headers_vert = [
    ("A5:A6", "STT"),
    ("B5:B6", "Ký hiệu\nbản vẽ"),
    ("C5:C6", "Mã hiệu\ncông tác"),
    ("D5:D6", "Danh mục công tác"),
    ("E5:E6", "Đơn vị"),
    ("F5:F6", "Khối lượng\ntoàn bộ"),
    ("P5:P6", "Đơn giá\n(Tổng hợp)"),
    ("Q5:Q6", "Định mức\náp dụng"),
    ("R5:R6", "Tổng thành tiền\n(=VL+NC+M)"),
    ("S5:S6", "Khu vực / Hạng mục\nthi công")
]

for range_str, text in headers_vert:
    ws_ct.merge_cells(range_str)
    c_top = ws_ct[range_str.split(':')[0]]
    c_top.value = text
    c_top.font = font_th
    c_top.fill = fill_th_main
    c_top.alignment = align_center

# Group headers:
ws_ct.merge_cells("G5:I5")
ws_ct["G5"] = "Tính trực tiếp (Đơn giá)"
ws_ct["G5"].font = font_th
ws_ct["G5"].fill = fill_th_main
ws_ct["G5"].alignment = align_center

ws_ct.merge_cells("J5:L5")
ws_ct["J5"] = "Thành tiền (đồng)"
ws_ct["J5"].font = font_th
ws_ct["J5"].fill = fill_th_main
ws_ct["J5"].alignment = align_center

ws_ct.merge_cells("M5:O5")
ws_ct["M5"] = "Hệ số công tác"
ws_ct["M5"].font = font_th
ws_ct["M5"].fill = fill_th_main
ws_ct["M5"].alignment = align_center

# Sub headers in row 6:
sub_headers = [
    ("G6", "Vật liệu"),
    ("H6", "Nhân công"),
    ("I6", "Máy thi công"),
    ("J6", "Vật liệu"),
    ("K6", "Nhân công"),
    ("L6", "Máy thi công"),
    ("M6", "VL"),
    ("N6", "NC"),
    ("O6", "M"),
]

for cell_ref, text in sub_headers:
    c_sub = ws_ct[cell_ref]
    c_sub.value = text
    c_sub.font = font_th
    c_sub.fill = fill_th_sub
    c_sub.alignment = align_center

# Đóng viền header
for r in [5, 6]:
    for c in range(1, 20):
        ws_ct.cell(r, c).border = border_header

print(">>> [3/5] Đang ghi 856 dịch vụ vào sheet 'Công trình'...")
start_r = 7
for idx, svc in enumerate(services, start=start_r):
    ws_ct.row_dimensions[idx].height = 20
    
    ws_ct.cell(idx, 1, svc['stt']).alignment = align_center # A: STT
    ws_ct.cell(idx, 2, "").alignment = align_center          # B: Bản vẽ
    ws_ct.cell(idx, 3, svc['code']).alignment = align_center # C: Mã hiệu
    ws_ct.cell(idx, 4, svc['name']).alignment = align_left   # D: Tên công tác
    ws_ct.cell(idx, 5, svc['dvt']).alignment = align_center  # E: ĐVT
    
    # F: Khối lượng Q
    c_q = ws_ct.cell(idx, 6, svc['q'])
    c_q.alignment = align_right
    c_q.number_format = '#,##0.00' if isinstance(svc['q'], float) else '#,##0'
    
    # G, H, I: Đơn giá trực tiếp (VL, NC, M)
    c_g = ws_ct.cell(idx, 7, 0)
    c_g.alignment = align_right
    c_g.number_format = '#,##0'
    
    c_h = ws_ct.cell(idx, 8, 0)
    c_h.alignment = align_right
    c_h.number_format = '#,##0'
    
    c_i = ws_ct.cell(idx, 9, 0)
    c_i.alignment = align_right
    c_i.number_format = '#,##0'
    
    # J, K, L: Thành tiền (=ROUND(F*ĐG*Hệ số, 0))
    c_j = ws_ct.cell(idx, 10, f"=ROUND(F{idx}*G{idx}*M{idx}, 0)")
    c_j.alignment = align_right
    c_j.number_format = '#,##0'
    
    c_k = ws_ct.cell(idx, 11, f"=ROUND(F{idx}*H{idx}*N{idx}, 0)")
    c_k.alignment = align_right
    c_k.number_format = '#,##0'
    
    c_l = ws_ct.cell(idx, 12, f"=ROUND(F{idx}*I{idx}*O{idx}, 0)")
    c_l.alignment = align_right
    c_l.number_format = '#,##0'
    
    # M, N, O: Hệ số công tác (1.0000)
    c_m = ws_ct.cell(idx, 13, 1.0)
    c_m.alignment = align_center
    c_m.number_format = '0.0000'
    
    c_n = ws_ct.cell(idx, 14, 1.0)
    c_n.alignment = align_center
    c_n.number_format = '0.0000'
    
    c_o = ws_ct.cell(idx, 15, 1.0)
    c_o.alignment = align_center
    c_o.number_format = '0.0000'
    
    # P: Đơn giá tổng hợp (=G+H+I)
    c_p = ws_ct.cell(idx, 16, f"=G{idx}+H{idx}+I{idx}")
    c_p.alignment = align_right
    c_p.number_format = '#,##0'
    
    # Q: Định mức
    ws_ct.cell(idx, 17, svc['note']).alignment = align_left
    
    # R: Tổng thành tiền (=J+K+L)
    c_r = ws_ct.cell(idx, 18, f"=J{idx}+K{idx}+L{idx}")
    c_r.alignment = align_right
    c_r.number_format = '#,##0'
    
    # S: Khu vực / Hạng mục
    ws_ct.cell(idx, 19, svc['loc']).alignment = align_left
    
    # Styling cell
    for c in range(1, 20):
        cell_obj = ws_ct.cell(idx, c)
        cell_obj.font = font_regular
        cell_obj.border = border_thin
        if idx % 2 == 0:
            cell_obj.fill = fill_zebra

end_r = start_r + len(services) - 1 # Row 862
thm_r = end_r + 1                   # Row 863

# Dòng THM - TỔNG HẠNG MỤC
ws_ct.row_dimensions[thm_r].height = 26
ws_ct.cell(thm_r, 1, "").alignment = align_center
ws_ct.cell(thm_r, 2, "THM").alignment = align_center

ws_ct.merge_cells(f"C{thm_r}:E{thm_r}")
c_thm_title = ws_ct.cell(thm_r, 3, "TỔNG HẠNG MỤC")
c_thm_title.alignment = align_center

ws_ct.cell(thm_r, 6, "").alignment = align_right
ws_ct.cell(thm_r, 7, "").alignment = align_right
ws_ct.cell(thm_r, 8, "").alignment = align_right
ws_ct.cell(thm_r, 9, "").alignment = align_right

# Tổng thành tiền VL, NC, M
ws_ct.cell(thm_r, 10, f"=SUM(J{start_r}:J{end_r})").alignment = align_right
ws_ct.cell(thm_r, 10).number_format = '#,##0'

ws_ct.cell(thm_r, 11, f"=SUM(K{start_r}:K{end_r})").alignment = align_right
ws_ct.cell(thm_r, 11).number_format = '#,##0'

ws_ct.cell(thm_r, 12, f"=SUM(L{start_r}:L{end_r})").alignment = align_right
ws_ct.cell(thm_r, 12).number_format = '#,##0'

ws_ct.cell(thm_r, 13, "").alignment = align_center
ws_ct.cell(thm_r, 14, "").alignment = align_center
ws_ct.cell(thm_r, 15, "").alignment = align_center
ws_ct.cell(thm_r, 16, "").alignment = align_right
ws_ct.cell(thm_r, 17, "").alignment = align_left

# Tổng chi phí trực tiếp (=J + K + L)
ws_ct.cell(thm_r, 18, f"=SUM(R{start_r}:R{end_r})").alignment = align_right
ws_ct.cell(thm_r, 18).number_format = '#,##0'
ws_ct.cell(thm_r, 19, "").alignment = align_left

for c in range(1, 20):
    c_cell = ws_ct.cell(thm_r, c)
    c_cell.font = font_thm
    c_cell.fill = fill_thm
    c_cell.border = border_thm

# Phần chữ ký phê duyệt (Rows 865 - 872)
sign_r1 = thm_r + 2
ws_ct.row_dimensions[sign_r1].height = 22
ws_ct.merge_cells(f"B{sign_r1}:E{sign_r1}")
ws_ct[f"B{sign_r1}"] = "Người thực hiện"
ws_ct[f"B{sign_r1}"].font = font_bold
ws_ct[f"B{sign_r1}"].alignment = align_center

ws_ct.merge_cells(f"K{sign_r1}:P{sign_r1}")
ws_ct[f"K{sign_r1}"] = "Người chủ trì"
ws_ct[f"K{sign_r1}"].font = font_bold
ws_ct[f"K{sign_r1}"].alignment = align_center

sign_r2 = sign_r1 + 6
ws_ct.row_dimensions[sign_r2].height = 20
ws_ct.merge_cells(f"K{sign_r2}:S{sign_r2}")
ws_ct[f"K{sign_r2}"] = "Chứng chỉ hành nghề hoạt động xây dựng số: ......................................."
ws_ct[f"K{sign_r2}"].font = font_italic
ws_ct[f"K{sign_r2}"].alignment = align_left

sign_r3 = sign_r2 + 1
ws_ct.row_dimensions[sign_r3].height = 20
ws_ct.merge_cells(f"K{sign_r3}:S{sign_r3}")
ws_ct[f"K{sign_r3}"] = "Lĩnh vực hành nghề: Định giá xây dựng hạng ..........................................."
ws_ct[f"K{sign_r3}"].font = font_italic
ws_ct[f"K{sign_r3}"].alignment = align_left

# Cài đặt độ rộng cột
col_widths = {
    'A': 6,   # STT
    'B': 10,  # Bản vẽ
    'C': 15,  # Mã hiệu
    'D': 46,  # Tên công tác
    'E': 9,   # ĐVT
    'F': 12,  # Khối lượng Q
    'G': 12,  # ĐG VL
    'H': 12,  # ĐG NC
    'I': 12,  # ĐG Máy
    'J': 14,  # TT VL
    'K': 14,  # TT NC
    'L': 14,  # TT Máy
    'M': 8,   # HS VL
    'N': 8,   # HS NC
    'O': 8,   # HS Máy
    'P': 14,  # Đơn giá
    'Q': 20,  # Định mức
    'R': 16,  # Tổng TT
    'S': 25,  # Khu vực
}
for col_let, width in col_widths.items():
    ws_ct.column_dimensions[col_let].width = width

# Freeze panes tại F7 (Khóa từ cột E và dòng 6 trở lên)
ws_ct.freeze_panes = "F7"

print(">>> [4/5] Đang liên kết sheet 'hao phi VT' sang sheet 'Công trình'...")
ws_hp = wb['hao phi VT']

# Tạo bản đồ STT -> Dòng trong 'Công trình'
# STT 1 nằm ở dòng 7, STT 2 ở dòng 8, ..., STT i ở dòng i + 6
stt_linked_count = 0
for r_idx in range(7, ws_hp.max_row + 1):
    c_phanloai = str(ws_hp.cell(r_idx, 3).value or '').strip().upper()
    c_stt = ws_hp.cell(r_idx, 1).value
    
    if c_phanloai == 'CÔNG TÁC' and c_stt is not None:
        try:
            stt_int = int(str(c_stt).replace('=', '').replace("'Công trình'!", '').replace('A', ''))
        except ValueError:
            # Trường hợp c_stt là công thức cũ hoặc chuỗi
            stt_int = None
            # Tìm stt từ dữ liệu trước đó
        
        # Để an toàn, đọc STT từ danh sách services theo thứ tự nếu cần
        # Vì hao phi VT có đúng 856 công tác theo thứ tự STT 1 đến 856:
        stt_current = stt_linked_count + 1
        r_ct = stt_current + 6 # Dòng tương ứng trong sheet 'Công trình'
        
        # Link Cột A (STT):
        ws_hp.cell(r_idx, 1, f"='Công trình'!A{r_ct}").alignment = align_center
        
        # Link Cột B (Mã hiệu định mức):
        ws_hp.cell(r_idx, 2, f"='Công trình'!C{r_ct}").alignment = align_center
        
        # Link Cột D (Khu vực / Hạng mục):
        ws_hp.cell(r_idx, 4, f"='Công trình'!S{r_ct}").alignment = align_left
        
        # Link Cột E (Tên công tác / công trình):
        ws_hp.cell(r_idx, 5, f"='Công trình'!D{r_ct}").alignment = align_left
        
        # Link Cột F (Đơn vị tính):
        ws_hp.cell(r_idx, 6, f"='Công trình'!E{r_ct}").alignment = align_center
        
        # Link Cột G (Khối lượng công trình Q):
        ws_hp.cell(r_idx, 7, f"='Công trình'!F{r_ct}").alignment = align_right
        ws_hp.cell(r_idx, 7).number_format = '#,##0.00' if isinstance(services[stt_linked_count]['q'], float) else '#,##0'
        
        stt_linked_count += 1

print(f"    [OK] Đã liên kết thành công toàn bộ {stt_linked_count} dòng CÔNG TÁC từ 'hao phi VT' sang 'Công trình'!")

print(">>> [5/5] Đang lưu workbook và sao chép...")
saved_path = PROJECT_PATH
try:
    wb.save(PROJECT_PATH)
    print(f"    [OK] Đã lưu file chính: {PROJECT_PATH}")
except PermissionError:
    alt_project = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"
    wb.save(alt_project)
    saved_path = alt_project
    print(f"    [LƯU Ý] File chính đang được mở trong Excel! Đã lưu bản cập nhật sang: {alt_project}")

try:
    shutil.copy2(saved_path, DOWNLOADS_PATH)
    print(f"    [OK] Đã sao chép sang: {DOWNLOADS_PATH}")
except PermissionError:
    shutil.copy2(saved_path, ALT_DOWNLOADS)
    print(f"    [OK] File Downloads đang mở trong Excel, đã copy sang: {ALT_DOWNLOADS}")

print("Xây dựng sheet 'Công trình' và liên kết 'hao phi VT' hoàn tất 100%!")
