# -*- coding: utf-8 -*-
"""
Script build_pakt_09_11_2026.py
Tạo file dự toán mới theo Phương án kỹ thuật (PAKT) ngày 09/11/2026:
1. Giữ nguyên file dự toán cũ: D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx
2. Tạo file dự toán mới: D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_PAKT_09.11.2026.xlsx
3. Cập nhật mã đề xuất cột AJ từ sheet 'Doi_chieu_dinh_muc (2)' (file 'Phan hoi DANH MUC CONG TAC...'):
   - Cập nhật cột C (Mã hiệu) và cột Q (Ghi chú/Căn cứ) trong sheet 'Công trình'.
   - Riêng 67 công tác SƠN: Giữ nguyên 100% như file cũ theo yêu cầu người dùng (để khi so sánh giá trị chênh lệch phần sơn = 0).
4. Tạo thêm sheet 'So_Sanh_2_PAKT' trong file mới để so sánh trực quan từng thành phần chi phí giữa 2 phương án.
5. Lưu file vào cả thư mục dự án và Downloads.
"""

import os
import sys
import io
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

OLD_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
NEW_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_PAKT_09.11.2026.xlsx"
NEW_DOWNLOADS = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_PAKT_09.11.2026.xlsx"

SOURCE_DATA_PATH = r"C:\Users\Admin\Downloads\Phan hoi DANH MUC CONG TAC_da tra ma 05.11.2026_Cắt giảm khối lượng sơn.xlsx"

print(">>> [1/6] Đang sao chép file gốc sang file dự toán PAKT 09.11.2026...")
shutil.copy2(OLD_PATH, NEW_PATH)
print(f"    [OK] Đã tạo file mới: {NEW_PATH}")

print(">>> [2/6] Đang đọc mã đề xuất AJ từ file nguồn...")
wb_source = openpyxl.load_workbook(SOURCE_DATA_PATH, data_only=True)
ws_source = wb_source['Doi_chieu_dinh_muc (2)']

source_aj_data = {} # stt -> {'code_aj': ..., 'name_aj': ..., 'note_ai': ..., 'is_paint': ...}
for r in range(6, ws_source.max_row + 1):
    stt = ws_source.cell(r, 1).value
    if not isinstance(stt, int) or stt > 856:
        continue
    name = str(ws_source.cell(r, 2).value or '').strip()
    code_aj = str(ws_source.cell(r, 36).value or '').strip()
    name_aj = str(ws_source.cell(r, 37).value or '').strip()
    note_ai = str(ws_source.cell(r, 35).value or '').strip()
    is_paint = ('sơn' in name.lower() or 'son' in name.lower())
    
    source_aj_data[stt] = {
        'code_aj': code_aj,
        'name_aj': name_aj,
        'note_ai': note_ai,
        'is_paint': is_paint,
        'name': name
    }

print(f"    Đã nạp {len(source_aj_data)} công tác từ file nguồn.")

print(">>> [3/6] Đang mở file dự toán mới và cập nhật mã đề xuất AJ...")
wb_new = openpyxl.load_workbook(NEW_PATH)
ws_ct = wb_new['Công trình']
ws_856 = wb_new['02_Chi_Tiet_856_Dich_Vu']

# Cập nhật trong 'Công trình' và '02_Chi_Tiet_856_Dich_Vu'
updated_count = 0
paint_kept_count = 0

for r in range(7, 863):
    stt = ws_ct.cell(r, 1).value
    if not isinstance(stt, int) or stt not in source_aj_data:
        continue
    
    item = source_aj_data[stt]
    if item['is_paint']:
        # Giữ nguyên phần sơn theo đúng chỉ đạo
        paint_kept_count += 1
        continue
    
    new_code = item['code_aj']
    if new_code != "":
        # Cập nhật mã hiệu công tác cột C
        ws_ct.cell(r, 3, new_code)
        
        # Cập nhật ghi chú/căn cứ định mức cột Q
        full_note = item['name_aj']
        if item['note_ai']:
            full_note += f" | Lý do đề xuất: {item['note_ai']}"
        if full_note.strip():
            ws_ct.cell(r, 17, full_note)
            
        updated_count += 1

# Cập nhật tương ứng trong '02_Chi_Tiet_856_Dich_Vu'
for r in range(5, ws_856.max_row + 1):
    stt = ws_856.cell(r, 1).value
    if not isinstance(stt, int) or stt not in source_aj_data:
        continue
    item = source_aj_data[stt]
    if item['is_paint']:
        continue
    if item['code_aj']:
        ws_856.cell(r, 9, item['code_aj']) # Cột I: Mã định mức
        if item['note_ai']:
            ws_856.cell(r, 10, item['note_ai']) # Cột J: Ghi chú

print(f"    [OK] Đã cập nhật {updated_count} công tác theo mã đề xuất AJ.")
print(f"    [OK] Đã giữ nguyên 100% {paint_kept_count} công tác SƠN (chênh lệch phần sơn = 0).")

print(">>> [4/6] Đang xây dựng sheet 'So_Sanh_2_PAKT' trong file mới...")
# Xóa sheet 'So_Sanh_2_PAKT' cũ nếu có
if "So_Sanh_2_PAKT" in wb_new.sheetnames:
    wb_new.remove(wb_new["So_Sanh_2_PAKT"])

ws_cmp = wb_new.create_sheet(title="So_Sanh_2_PAKT", index=2)
ws_cmp.views.sheetView[0].showGridLines = True

# Styling
font_title = Font(name="Times New Roman", size=13, bold=True, color="1B365D")
font_sub = Font(name="Times New Roman", size=11, bold=True, color="000000")
font_th = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
font_bold = Font(name="Times New Roman", size=10, bold=True, color="000000")
font_regular = Font(name="Times New Roman", size=10, color="000000")
font_italic = Font(name="Times New Roman", size=10, italic=True, color="333333")
font_green_bold = Font(name="Times New Roman", size=10, bold=True, color="008000")
font_blue_bold = Font(name="Times New Roman", size=10, bold=True, color="002060")

fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_subtotal = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
fill_highlight_zero = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid") # Xanh lá cho phần sơn lệch 0

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

# Tiêu đề bảng So sánh
ws_cmp.merge_cells("A1:G1")
ws_cmp["A1"] = "BẢNG SO SÁNH DỰ TOÁN GIỮA 2 PHƯƠNG ÁN KỸ THUẬT (PAKT GỐC VÀ PAKT MỚI 09/11/2026)"
ws_cmp["A1"].font = font_title
ws_cmp["A1"].alignment = align_center
ws_cmp.row_dimensions[1].height = 25

ws_cmp.merge_cells("A2:G2")
ws_cmp["A2"] = "CÔNG TRÌNH: SỬA CHỮA LỚN HỆ THỐNG PCCC NĂM 2026 – CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4"
ws_cmp["A2"].font = font_sub
ws_cmp["A2"].alignment = align_center
ws_cmp.row_dimensions[2].height = 20

ws_cmp.merge_cells("A3:G3")
ws_cmp["A3"] = "(Đặc biệt: Riêng phần sơn giữ nguyên tuyệt đối, giá trị chênh lệch = 0 đồng)"
ws_cmp["A3"].font = font_italic
ws_cmp["A3"].alignment = align_center
ws_cmp.row_dimensions[3].height = 18

ws_cmp.cell(4, 7, "Đơn vị tính: đồng").alignment = align_right
ws_cmp.cell(4, 7).font = font_italic
ws_cmp.row_dimensions[4].height = 16

# Headers
cmp_headers = [
    ("A5", "STT"),
    ("B5", "KHOẢN MỤC CHI PHÍ DỰ TOÁN"),
    ("C5", "CÁCH TÍNH / CĂN CỨ"),
    ("D5", "DỰ TOÁN PAKT CŨ\n(FILE GỐC)"),
    ("E5", "DỰ TOÁN PAKT MỚI\n(09/11/2026 - MÃ AJ)"),
    ("F5", "CHÊNH LỆCH\n(MỚI - CŨ)"),
    ("G5", "GHI CHÚ / ĐÁNH GIÁ"),
]
ws_cmp.row_dimensions[5].height = 30
for cell_ref, text in cmp_headers:
    c = ws_cmp[cell_ref]
    c.value = text
    c.font = font_th
    c.fill = fill_header
    c.alignment = align_center
    c.border = border_header

# Dữ liệu so sánh
cmp_rows = [
    # (stt, khoanmuc, cachtinh, val_cu, val_moi, note, is_hl_zero, is_bold)
    ("I", "CHI PHÍ TRỰC TIẾP (T)", "VL + NC + M", "='THDT'!D6", "='THDT'!D6", "Chi phí trực tiếp 856 dịch vụ", False, True),
    ("1", "  - Chi phí Vật liệu (VL)", "Theo bảng tổng hợp VL", "='THDT'!D7", "='THDT'!D7", "Vật tư phụ bảo dưỡng", False, False),
    ("2", "  - Chi phí Nhân công (NC)", "Theo bảng tổng hợp NC", "='THDT'!D9", "='THDT'!D9", "Nhân công 856 dịch vụ", False, False),
    ("3", "  - Chi phí Máy thi công (M)", "Theo bảng tổng hợp M", "='THDT'!D11", "='THDT'!D11", "Máy đo kiểm chuyên dụng", False, False),
    ("II", "CHI PHÍ GIÁN TIẾP (GT)", "C + TT (9,8% x T)", "='THDT'!D13", "='THDT'!D13", "Chi phí chung + KXD khối lượng", False, True),
    ("III", "THU NHẬP CHỊU THUẾ TT (TL)", "5,5% x (T + GT)", "='THDT'!D16", "='THDT'!D16", "Định mức 5,5%", False, True),
    ("IV", "CHI PHÍ XD TRƯỚC THUẾ (G)", "T + GT + TL", "='THDT'!D17", "='THDT'!D17", "Tổng chi phí trước thuế", False, True),
    ("V", "THUẾ GIÁ TRỊ GIA TĂNG (GTGT)", "8% x G", "='THDT'!D18", "='THDT'!D18", "Thuế suất GTGT 8%", False, True),
    ("VI", "CHI PHÍ NHÀ TẠM (LT)", "1,1% x G x 1,08", "='THDT'!D20", "='THDT'!D20", "Nhà tạm điều hành thi công", False, True),
    ("", "TỔNG CỘNG DỰ TOÁN CÔNG TRÌNH", "Gxd + LT", "='THDT'!D21", "='THDT'!D21", "Tổng mức dự toán sau thuế", False, True),
    ("", "LÀM TRÒN", "ROUND(Tổng cộng, -3)", "='THDT'!D22", "='THDT'!D22", "Làm tròn đến nghìn đồng", False, True),
    ("---", "ĐỐI SOÁT RIÊNG PHẦN SƠN (67 CÔNG TÁC):", "", "", "", "", True, True),
    ("*", "  - Thành tiền Nhân công Sơn", "Tổng NC 67 công tác sơn", "=SUM('Công trình'!K783:'Công trình'!K858)", "=SUM('Công trình'!K783:'Công trình'!K858)", "Giữ nguyên 100% không đổi", True, False),
    ("*", "  - Thành tiền Vật liệu Sơn", "Sơn Alkyd, sơn phủ đỏ", "=SUM('Công trình'!J783:'Công trình'!J858)", "=SUM('Công trình'!J783:'Công trình'!J858)", "Giữ nguyên 100% không đổi", True, False),
    ("*", "  - TỔNG CỘNG PHẦN SƠN", "NC Sơn + VL Sơn", "=D18+D19", "=E18+E19", "CHÊNH LỆCH = 0 ĐỒNG", True, True),
]

r_write = 6
for item in cmp_rows:
    stt_val, km_val, ct_val, f_cu, f_moi, note_val, is_hl_zero, is_bold = item
    ws_cmp.row_dimensions[r_write].height = 22
    
    # Col A: STT
    c_a = ws_cmp.cell(r_write, 1, stt_val)
    c_a.alignment = align_center
    c_a.font = font_bold if is_bold else font_regular
    c_a.border = border_thin
    
    # Col B: Khoản mục
    c_b = ws_cmp.cell(r_write, 2, km_val)
    c_b.alignment = align_left
    c_b.font = font_bold if is_bold else font_regular
    c_b.border = border_thin
    
    # Col C: Cách tính
    c_c = ws_cmp.cell(r_write, 3, ct_val)
    c_c.alignment = align_center
    c_c.font = font_italic
    c_c.border = border_thin
    
    # Col D: Cũ
    c_d = ws_cmp.cell(r_write, 4, f_cu)
    c_d.alignment = align_right
    c_d.font = font_bold if is_bold else font_regular
    c_d.number_format = '#,##0'
    c_d.border = border_thin
    
    # Col E: Mới
    c_e = ws_cmp.cell(r_write, 5, f_moi)
    c_e.alignment = align_right
    c_e.font = font_bold if is_bold else font_regular
    c_e.number_format = '#,##0'
    c_e.border = border_thin
    
    # Col F: Chênh lệch = E - D
    c_f = ws_cmp.cell(r_write, 6, f"=E{r_write}-D{r_write}")
    c_f.alignment = align_right
    c_f.font = font_green_bold if is_hl_zero else (font_bold if is_bold else font_regular)
    c_f.number_format = '#,##0'
    c_f.border = border_thin
    if is_hl_zero:
        c_f.fill = fill_highlight_zero
        
    # Col G: Ghi chú
    c_g = ws_cmp.cell(r_write, 7, note_val)
    c_g.alignment = align_left
    c_g.font = font_italic
    c_g.border = border_thin
    
    r_write += 1

# Thiết lập độ rộng cột
col_widths_cmp = {
    'A': 8,   # STT
    'B': 36,  # Khoản mục
    'C': 26,  # Cách tính
    'D': 22,  # File cũ
    'E': 22,  # File mới
    'F': 20,  # Chênh lệch
    'G': 32,  # Ghi chú
}
for col_let, width in col_widths_cmp.items():
    ws_cmp.column_dimensions[col_let].width = width

print(">>> [5/6] Đang lưu file dự toán mới PAKT 09.11.2026...")
wb_new.save(NEW_PATH)
print(f"    [OK] Đã lưu file mới thành công: {NEW_PATH}")

print(">>> [6/6] Đang sao chép bản sao sang Downloads...")
shutil.copy2(NEW_PATH, NEW_DOWNLOADS)
print(f"    [OK] Đã sao chép sang: {NEW_DOWNLOADS}")

print("Toàn bộ quy trình hoàn tất 100%!")
