# -*- coding: utf-8 -*-
"""
Script integrate_lam_dong_prices.py
Đưa các Sheet Đơn giá Lâm Đồng năm 2026 vào BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx
và thiết lập công thức liên kết đơn giá từ các Sheet Nhân công, Máy thi công, Vật liệu
sang các Sheet Đơn giá Lâm Đồng tương ứng.
"""

import os
import sys
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PROJECT_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
LD_SOURCE_PATH = r"d:\TaskApp_kiet\LapDuToanApp\BANG_DON_GIA_LAM_DONG_2026.xlsx"
DOWNLOADS_PATH = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"

print(">>> [1/5] Đang mở workbook Dự toán PCCC và workbook Đơn giá Lâm Đồng...")
wb_proj = openpyxl.load_workbook(PROJECT_PATH, data_only=False)
wb_ld = openpyxl.load_workbook(LD_SOURCE_PATH, data_only=False)

ld_sheets = [
    "00_ThuyetMinh_CanCu",
    "01_DonGia_NhanCong",
    "02_DonGia_CaMay",
    "03_VLXD_TongHop_17Nhom",
    "04_VLXD_ChiTiet_ToanTinh",
    "05_VLXD_GiaTaiMo_NoiSX",
    "06_VLXD_GiaTaiNoiBan"
]

print(">>> [2/5] Đang sao chép các Sheet Đơn giá Lâm Đồng vào workbook Dự toán...")
for sname in ld_sheets:
    if sname in wb_proj.sheetnames:
        print(f"    - Xóa sheet cũ: {sname}")
        wb_proj.remove(wb_proj[sname])
        
    src_ws = wb_ld[sname]
    dst_ws = wb_proj.create_sheet(title=sname)
    dst_ws.views.sheetView[0].showGridLines = True
    
    # Copy merged cells
    for merge_range in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(merge_range))
        
    # Copy rows and cells
    for r in range(1, src_ws.max_row + 1):
        dst_ws.row_dimensions[r].height = src_ws.row_dimensions[r].height
        for c in range(1, src_ws.max_column + 1):
            src_cell = src_ws.cell(r, c)
            dst_cell = dst_ws.cell(r, c, src_cell.value)
            if src_cell.has_style:
                dst_cell.font = src_cell.font.copy() if src_cell.font else None
                dst_cell.border = src_cell.border.copy() if src_cell.border else None
                dst_cell.fill = src_cell.fill.copy() if src_cell.fill else None
                dst_cell.number_format = src_cell.number_format
                dst_cell.alignment = src_cell.alignment.copy() if src_cell.alignment else None
                
    # Copy column widths
    for col_letter, col_dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col_letter].width = col_dim.width
        
    if src_ws.freeze_panes:
        dst_ws.freeze_panes = src_ws.freeze_panes
    if src_ws.auto_filter.ref:
        dst_ws.auto_filter.ref = src_ws.auto_filter.ref
        
    print(f"    [OK] Đã sao chép Sheet '{sname}' ({dst_ws.max_row} dòng, {dst_ws.max_column} cột)")

print(">>> [3/5] Đang cập nhật công thức liên kết Đơn giá cho 3 sheet Tài nguyên...")

# 1. Sheet Nhân công
ws_nc = wb_proj["Nhân công"]
ws_nc.cell(5, 7, "Đơn giá áp dụng (VNĐ)\n(Link '01_DonGia_NhanCong')")
ws_nc.cell(5, 9, "Căn cứ đơn giá Lâm Đồng &\nGhi chú liên kết công thức")

# Map code -> (formula, note)
nc_price_map = {
    "NC_E6C.3.05": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_E6C.3.02_B5": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_E6C.3.02_B1-11": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_BD.41111": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
    "NC_E6C.3.01_B5": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_E6C.3.01_B1-11": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_BD.41131": ("='01_DonGia_NhanCong'!F46", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 4,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F46"),
    "NC_BD.41141": ("='01_DonGia_NhanCong'!F44", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F44"),
    "NC_E6C.3.03": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_BD.41121": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
    "NC_E6C.1.23": ("='01_DonGia_NhanCong'!F55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F55"),
    "NC_E6C.1.31": ("='01_DonGia_NhanCong'!F55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F55"),
    "NC_E6C.2.04": ("='01_DonGia_NhanCong'!F55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F55"),
    "NC_E6C.1.18": ("='01_DonGia_NhanCong'!F55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F55"),
    "NC_BA.18405": ("='01_DonGia_NhanCong'!F52", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 3,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F52"),
    "NC_E6C.3.04": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_E6B.1.13.01": ("='01_DonGia_NhanCong'!F55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F55"),
    "NC_E6C.3.06": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_DB.24001": ("='01_DonGia_NhanCong'!E19", "VB 5898/SXD-QLXD Lâm Đồng (Kỹ sư KS-TN Vùng II) - Link '01_DonGia_NhanCong'!E19"),
    "NC_E6BOP.1.03": ("='01_DonGia_NhanCong'!F47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F47"),
    "NC_E6BOP.1.01": ("='01_DonGia_NhanCong'!F55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F55"),
    "NC_BD.41171": ("='01_DonGia_NhanCong'!F52", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 3,0/7 Vùng II) - Link '01_DonGia_NhanCong'!F52"),
    "NC_SB.82633": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
    "NC_BB.86601": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
    "NC_BB.86103": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
    "NC_BB.85201": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
    "NC_SB.82634": ("='01_DonGia_NhanCong'!F45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 Vùng II) - Link '01_DonGia_NhanCong'!F45"),
}

count_nc = 0
for r in range(6, ws_nc.max_row + 1):
    c1 = str(ws_nc.cell(r, 1).value or '')
    if c1.isdigit():
        code = str(ws_nc.cell(r, 2).value or '').strip()
        if code in nc_price_map:
            formula, note = nc_price_map[code]
            ws_nc.cell(r, 7, formula)
            ws_nc.cell(r, 8, f"=ROUND(F{r}*G{r}, 0)")
            ws_nc.cell(r, 9, note)
            count_nc += 1

print(f"    [OK] Đã gắn công thức link giá cho {count_nc}/27 Master row của Sheet 'Nhân công'.")

# 2. Sheet Máy thi công
ws_cm = wb_proj["Máy thi công"]
ws_cm.cell(5, 7, "Đơn giá áp dụng (VNĐ)\n(Link '02_DonGia_CaMay')")
ws_cm.cell(5, 9, "Căn cứ đơn giá Lâm Đồng &\nGhi chú liên kết công thức")

cm_price_map = {
    "M_E6C.3.05_VOM": ("='02_DonGia_CaMay'!M677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng Vùng II) - Link '02_DonGia_CaMay'!M677"),
    "M_E6C.3.05_BUI": ("='02_DonGia_CaMay'!M459", "QĐ 21/QĐ-SXD Lâm Đồng (M112.4301: Máy cầm tay Vùng II) - Link '02_DonGia_CaMay'!M459"),
    "M_E6C.3.02_SOLO": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra thử nghiệm Vùng II) - Link '02_DonGia_CaMay'!M684"),
    "M_E6C.3.02_VOM": ("='02_DonGia_CaMay'!M677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng Vùng II) - Link '02_DonGia_CaMay'!M677"),
    "M_BD.41111": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực Vùng II) - Link '02_DonGia_CaMay'!M684"),
    "M_E6C.3.01_MEG": ("='02_DonGia_CaMay'!M683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Mê gôm mét Vùng II) - Link '02_DonGia_CaMay'!M683"),
    "M_BD.41131": ("='02_DonGia_CaMay'!M685", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0027: TB tạo dòng điện Vùng II) - Link '02_DonGia_CaMay'!M685"),
    "M_E6C.3.03_MEG": ("='02_DonGia_CaMay'!M683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Mê gôm mét Vùng II) - Link '02_DonGia_CaMay'!M683"),
    "M_BD.41121": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực Vùng II) - Link '02_DonGia_CaMay'!M684"),
    "M_E6C.1.23_BOM": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực Vùng II) - Link '02_DonGia_CaMay'!M684"),
    "M_E6C.1.31_GAS": ("='02_DonGia_CaMay'!M664", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0006: Hợp bộ phân tích hàm lượng khí Vùng II) - Link '02_DonGia_CaMay'!M664"),
    "M_E6C.2.04_VOM": ("='02_DonGia_CaMay'!M677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng Vùng II) - Link '02_DonGia_CaMay'!M677"),
    "M_E6C.1.18_VOM": ("='02_DonGia_CaMay'!M677", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0019: Máy đo vạn năng Vùng II) - Link '02_DonGia_CaMay'!M677"),
    "M_E6C.3.04_SOLO": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra thử nghiệm Vùng II) - Link '02_DonGia_CaMay'!M684"),
    "DC_E6C.3.04_S200": ("='02_DonGia_CaMay'!M683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Dụng cụ thử nghiệm Vùng II) - Link '02_DonGia_CaMay'!M683"),
    "M_E6B.1.13.01_TR": ("='02_DonGia_CaMay'!M672", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0014: Máy đo điện trở một chiều Vùng II) - Link '02_DonGia_CaMay'!M672"),
    "M_E6C.3.06_CD": ("='02_DonGia_CaMay'!M683", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0025: Mê gôm mét đo điện trở CĐ Vùng II) - Link '02_DonGia_CaMay'!M683"),
    "M_E6C.3.06_BAN": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực Vùng II) - Link '02_DonGia_CaMay'!M684"),
    "M_DB.24001": ("='02_DonGia_CaMay'!M673", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0015: Máy đo điện trở tiếp địa Vùng II) - Link '02_DonGia_CaMay'!M673"),
    "M_E6BOP.1.03": ("='02_DonGia_CaMay'!M148", "QĐ 21/QĐ-SXD Lâm Đồng (M102.1308: Hệ kích nâng / Palăng Vùng II) - Link '02_DonGia_CaMay'!M148"),
    "M_E6BOP.1.01": ("='02_DonGia_CaMay'!M684", "QĐ 21/QĐ-SXD Lâm Đồng (M203.0026: TB kiểm tra áp lực Vùng II) - Link '02_DonGia_CaMay'!M684"),
}

count_cm = 0
for r in range(6, ws_cm.max_row + 1):
    c1 = str(ws_cm.cell(r, 1).value or '')
    if c1.isdigit():
        code = str(ws_cm.cell(r, 2).value or '').strip()
        if code in cm_price_map:
            formula, note = cm_price_map[code]
            ws_cm.cell(r, 7, formula)
            ws_cm.cell(r, 8, f"=ROUND(F{r}*G{r}, 0)")
            ws_cm.cell(r, 9, note)
            count_cm += 1

print(f"    [OK] Đã gắn công thức link giá cho {count_cm}/21 Master row của Sheet 'Máy thi công'.")

# 3. Sheet Vật liệu
ws_vl = wb_proj["Vật liệu"]
ws_vl.cell(5, 7, "Đơn giá áp dụng (VNĐ)\n(Link '04_VLXD_ChiTiet_ToanTinh')")
ws_vl.cell(5, 9, "Căn cứ đơn giá Lâm Đồng &\nGhi chú liên kết công thức")

vl_price_map = {
    "VL_E6B.1.13.01": ("='04_VLXD_ChiTiet_ToanTinh'!N174", "Công bố giá VLXD T8/2026 Lâm Đồng (Dung môi/hóa chất lỏng dòng 174) - Link '04_VLXD_ChiTiet_ToanTinh'!N174"),
    "VL_SB.82633_LOT": ("='04_VLXD_ChiTiet_ToanTinh'!N178", "Công bố giá VLXD T8/2026 Lâm Đồng (Sơn giao thông lót kim loại dòng 178) - Link '04_VLXD_ChiTiet_ToanTinh'!N178"),
    "VL_SB.82633_PHU": ("='04_VLXD_ChiTiet_ToanTinh'!N198", "Công bố giá VLXD T8/2026 Lâm Đồng (Sơn phủ PEACE bảo vệ dòng 198) - Link '04_VLXD_ChiTiet_ToanTinh'!N198"),
}

count_vl = 0
for r in range(6, ws_vl.max_row + 1):
    c1 = str(ws_vl.cell(r, 1).value or '')
    if c1.isdigit():
        code = str(ws_vl.cell(r, 2).value or '').strip()
        if code in vl_price_map:
            formula, note = vl_price_map[code]
            ws_vl.cell(r, 7, formula)
            ws_vl.cell(r, 8, f"=ROUND(F{r}*G{r}, 0)")
            ws_vl.cell(r, 9, note)
            count_vl += 1

print(f"    [OK] Đã gắn công thức link giá cho {count_vl}/3 Master row của Sheet 'Vật liệu'.")

print(">>> [4/5] Đang kiểm tra định dạng và lưu file...")
# Set number formatting for G and H on all master rows
for ws, count in [(ws_nc, 27), (ws_cm, 21), (ws_vl, 3)]:
    for r in range(6, ws.max_row + 1):
        c1 = str(ws.cell(r, 1).value or '')
        if c1.isdigit():
            ws.cell(r, 7).number_format = '#,##0'
            ws.cell(r, 8).number_format = '#,##0'

wb_proj.save(PROJECT_PATH)
file_size_mb = os.path.getsize(PROJECT_PATH) / (1024 * 1024)
print(f"    [OK] Đã lưu thành công: {PROJECT_PATH} ({file_size_mb:.2f} MB)")

# Copy to Downloads
wb_proj.save(DOWNLOADS_PATH)
print(f"    [OK] Đã lưu bản sao sang: {DOWNLOADS_PATH}")

print(">>> [5/5] Hoàn tất 100% tích hợp Đơn giá Lâm Đồng và thiết lập Link công thức!")
