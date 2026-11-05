# -*- coding: utf-8 -*-
"""
Script update_vat_lieu_tam_tinh.py
Kiểm tra tính hợp lý của đường link đơn giá sheet 'Vật liệu' sang '04_VLXD_ChiTiet_ToanTinh':
1. [LOẠI 1] NƯỚC CẤT NẠP ACCU, MỠ BÔI CỌC BÌNH (VL_E6B.1.13.01, kg):
   - Link cũ: '04_VLXD_ChiTiet_ToanTinh'!N174 -> Là "Jothiner Joway" (dung môi pha sơn JOTON, 86.000 đ/kg).
   - Đánh giá: KHÔNG HỢP LÝ (nước cất và mỡ cọc bình không phải dung môi pha sơn).
   - Xử lý: Áp GIÁ TẠM TÍNH = 35.000 đ/kg, BÔI MÀU ĐỎ.

2. [LOẠI 2] SƠN CHỐNG GỈ ALKYD ĐỎ PCCC (VL_SB.82633_LOT, kg):
   - Link cũ: '04_VLXD_ChiTiet_ToanTinh'!N178 -> Là "Sơn giao thông lót" (sơn kẻ đường giao thông, 94.100 đ/kg).
   - Đánh giá: KHÔNG HỢP LÝ (sơn kẻ đường giao thông không phải sơn chống rỉ kim loại cho ống thép PCCC).
   - Xử lý: Áp GIÁ TẠM TÍNH = 60.000 đ/kg, BÔI MÀU ĐỎ.

3. [LOẠI 3] SƠN PHỦ MÀU ĐỎ PCCC (VL_SB.82633_PHU, kg):
   - Link cũ: '04_VLXD_ChiTiet_ToanTinh'!N198 -> Là "Sơn phủ PEACE 18L" (sơn nước tường nhà, 155.156 đ/lít).
   - Đánh giá: KHÔNG HỢP LÝ (sơn tường nước đóng thùng lít không dùng cho sơn ống kim loại tính theo kg).
   - Xử lý: Áp GIÁ TẠM TÍNH = 80.000 đ/kg, BÔI MÀU ĐỎ.
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
ws = wb['Vật liệu']

# Styling cho ô Giá Tạm Tính - Bôi màu đỏ
font_red_bold = Font(name="Times New Roman", size=10, bold=True, color="FF0000")
font_red_note = Font(name="Times New Roman", size=9, italic=True, bold=True, color="C00000")
font_master_bold = Font(name="Times New Roman", size=10, bold=True, color="1B365D")

# Nền đỏ/hồng nhạt cảnh báo chuyên nghiệp (Light Red / Pink soft highlight)
fill_red_warning = PatternFill(start_color="FFD8D8", end_color="FFD8D8", fill_type="solid")
fill_red_cell = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

border_red = Border(
    left=Side(style='thin', color='FF0000'),
    right=Side(style='thin', color='FF0000'),
    top=Side(style='thin', color='FF0000'),
    bottom=Side(style='thin', color='FF0000')
)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# Cập nhật thông tin tiêu đề giải thích thêm về Giá Tạm Tính bôi đỏ
ws.merge_cells("A4:I4")
ws["A4"] = "(*) LƯU Ý THẨM ĐỊNH: Các đơn giá vật liệu không có trong Công bố giá VLXD Lâm Đồng (hoặc chủng loại công bố không phù hợp) được áp GIÁ TẠM TÍNH THEO THỊ TRƯỜNG VÀ BÔI MÀU ĐỎ (cần thu thập đủ 03 báo giá khi lập dự toán chính thức)."
ws["A4"].font = Font(name="Times New Roman", size=10, italic=True, bold=True, color="C00000")
ws["A4"].alignment = align_left
ws["A4"].fill = PatternFill(start_color="FFF0F0", end_color="FFF0F0", fill_type="solid")
ws.row_dimensions[4].height = 24

# Danh sách vật liệu cần áp giá tạm tính và bôi màu đỏ:
# (row_index, new_temp_price, note_text)
vl_updates = [
    (
        6, # Row 6: Loại 1
        35000,
        "[TẠM TÍNH - BÔI ĐỎ] Không có trong CBG VLXD Lâm Đồng (link cũ N174 là dung môi Jothiner không phù hợp). Áp giá tạm tính thị trường 35.000 đ/kg (nước cất nạp accu & mỡ cọc bình), cần lấy 03 báo giá khi lập dự toán chính thức."
    ),
    (
        12, # Row 12: Loại 2
        60000,
        "[TẠM TÍNH - BÔI ĐỎ] Không có trong CBG VLXD Lâm Đồng (link cũ N178 là sơn kẻ đường giao thông không phù hợp). Áp giá tạm tính thị trường 60.000 đ/kg (sơn lót chống gỉ Alkyd đỏ kim loại), cần lấy 03 báo giá khi lập dự toán chính thức."
    ),
    (
        70, # Row 70: Loại 3
        80000,
        "[TẠM TÍNH - BÔI ĐỎ] Không có trong CBG VLXD Lâm Đồng (link cũ N198 là sơn tường nước PEACE không phù hợp). Áp giá tạm tính thị trường 80.000 đ/kg (sơn phủ Alkyd màu đỏ ống PCCC), cần lấy 03 báo giá khi lập dự toán chính thức."
    ),
]

print(">>> [2/4] Đang áp giá tạm tính và bôi màu đỏ cho các dòng Master trong sheet 'Vật liệu'...")
for r, price_val, note_text in vl_updates:
    # Cập nhật ô Đơn giá (Cột G = 7)
    cell_price = ws.cell(r, 7)
    cell_price.value = price_val
    cell_price.font = font_red_bold
    cell_price.fill = fill_red_cell
    cell_price.border = border_red
    cell_price.alignment = align_right
    cell_price.number_format = '#,##0'
    
    # Cập nhật ô Thành tiền (Cột H = 8)
    cell_total = ws.cell(r, 8)
    cell_total.value = f"=ROUND(F{r}*G{r}, 0)"
    cell_total.font = font_red_bold
    cell_total.fill = fill_red_cell
    cell_total.border = border_red
    cell_total.alignment = align_right
    cell_total.number_format = '#,##0'
    
    # Cập nhật ô Ghi chú (Cột I = 9)
    cell_note = ws.cell(r, 9)
    cell_note.value = note_text
    cell_note.font = font_red_note
    cell_note.fill = fill_red_warning
    cell_note.border = border_red
    cell_note.alignment = align_left
    
    # Highlight cả hàng Master để nhận diện rõ ràng
    for c in range(1, 7):
        c_cell = ws.cell(r, c)
        c_cell.font = font_master_bold
        c_cell.fill = fill_red_warning
        c_cell.border = border_red
        
    print(f"    [OK] Dòng {r:2d} ({ws.cell(r, 2).value}): Đơn giá tạm tính = {price_val:,} đ/kg (Đã bôi đỏ & đóng khung cảnh báo).")

print(">>> [3/4] Đang lưu file chính...")
wb.save(PROJECT_PATH)
print(f"    [OK] Đã lưu: {PROJECT_PATH}")

print(">>> [4/4] Đang sao chép sang Downloads...")
try:
    shutil.copy2(PROJECT_PATH, DOWNLOADS_PATH)
    print(f"    [OK] Đã copy sang: {DOWNLOADS_PATH}")
except PermissionError:
    shutil.copy2(PROJECT_PATH, ALT_DOWNLOADS)
    print(f"    [OK] File chính đang mở trong Excel, đã copy sang: {ALT_DOWNLOADS}")

print("Hoàn tất xử lý Giá tạm tính và Bôi màu đỏ cho sheet Vật liệu!")
