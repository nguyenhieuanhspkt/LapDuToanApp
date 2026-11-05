# -*- coding: utf-8 -*-
"""
Script update_nhan_cong_vung_iv.py
Cập nhật Đơn giá Nhân công sang Vùng IV theo yêu cầu của người dùng:
"Đơn giá nhân công thì dùng của Vùng IV nha."
Nguồn: '01_DonGia_NhanCong' Cột H (Bảng B) và Cột G (Bảng A)
Đồng thời cập nhật Ca máy liên kết sang Vùng IV (Cột O trong '02_DonGia_CaMay') để đồng bộ 100% Vùng IV.
"""

import os
import sys
import io
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PROJECT_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
DOWNLOADS_PATH = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"

print(">>> [1/4] Đang mở workbook Dự toán PCCC...")
wb = openpyxl.load_workbook(PROJECT_PATH, data_only=False)

# 1. Update Sheet Nhân công
ws_nc = wb["Nhân công"]
ws_nc.cell(5, 7, "Đơn giá áp dụng (VNĐ)\n(Link '01_DonGia_NhanCong' VÙNG IV)")
ws_nc.cell(5, 9, "Căn cứ đơn giá Lâm Đồng VÙNG IV &\nGhi chú liên kết công thức")

nc_vung4_map = {
    # Nhóm 3 - Bậc 5,0/7 Vùng IV: Row 47, Col H (400.892 đ)
    "NC_E6C.3.05": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.02_B5": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.02_B1-11": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.01_B5": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.01_B1-11": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.03": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.04": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6C.3.06": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),
    "NC_E6BOP.1.03": ("='01_DonGia_NhanCong'!H47", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H47"),

    # Nhóm 4 - Bậc 5,0/7 Vùng IV: Row 55, Col H (420.938 đ)
    "NC_E6C.1.23": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
    "NC_E6C.1.31": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
    "NC_E6C.2.04": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
    "NC_E6C.1.18": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
    "NC_E6B.1.13.01": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),
    "NC_E6BOP.1.01": ("='01_DonGia_NhanCong'!H55", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 5,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H55"),

    # Nhóm 3 - Bậc 3,5/7 Vùng IV: Row 45, Col H (312.490 đ)
    "NC_BD.41111": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    "NC_BD.41121": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    "NC_SB.82633": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    "NC_BB.86601": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    "NC_BB.86103": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    "NC_BB.85201": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),
    "NC_SB.82634": ("='01_DonGia_NhanCong'!H45", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,5/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H45"),

    # Nhóm 3 - Bậc 4,0/7 Vùng IV: Row 46, Col H (339.216 đ)
    "NC_BD.41131": ("='01_DonGia_NhanCong'!H46", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 4,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H46"),

    # Nhóm 3 - Bậc 3,0/7 Vùng IV: Row 44, Col H (285.764 đ)
    "NC_BD.41141": ("='01_DonGia_NhanCong'!H44", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3, Bậc 3,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H44"),

    # Nhóm 4 - Bậc 3,0/7 Vùng IV: Row 52, Col H (300.053 đ)
    "NC_BA.18405": ("='01_DonGia_NhanCong'!H52", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 3,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H52"),
    "NC_BD.41171": ("='01_DonGia_NhanCong'!H52", "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4, Bậc 3,0/7 VÙNG IV) - Link '01_DonGia_NhanCong'!H52"),

    # Kỹ sư KS-TN Vùng IV: Bảng A, Row 19, Col G (313.000 đ)
    "NC_DB.24001": ("='01_DonGia_NhanCong'!G19", "VB 5898/SXD-QLXD Lâm Đồng (Kỹ sư KS-TN VÙNG IV) - Link '01_DonGia_NhanCong'!G19"),
}

count_nc = 0
for r in range(6, ws_nc.max_row + 1):
    c1 = str(ws_nc.cell(r, 1).value or '')
    if c1.isdigit():
        code = str(ws_nc.cell(r, 2).value or '').strip()
        if code in nc_vung4_map:
            formula, note = nc_vung4_map[code]
            ws_nc.cell(r, 7, formula)
            ws_nc.cell(r, 8, f"=ROUND(F{r}*G{r}, 0)")
            ws_nc.cell(r, 9, note)
            count_nc += 1

print(f"    [OK] Đã cập nhật Đơn giá Nhân công VÙNG IV cho {count_nc}/27 Master row.")

# 2. Update Sheet Máy thi công sang Vùng IV (Cột O trong '02_DonGia_CaMay')
ws_cm = wb["Máy thi công"]
ws_cm.cell(5, 7, "Đơn giá áp dụng (VNĐ)\n(Link '02_DonGia_CaMay' VÙNG IV)")
ws_cm.cell(5, 9, "Căn cứ đơn giá Lâm Đồng VÙNG IV &\nGhi chú liên kết công thức")

cm_vung4_map = {
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

count_cm = 0
for r in range(6, ws_cm.max_row + 1):
    c1 = str(ws_cm.cell(r, 1).value or '')
    if c1.isdigit():
        code = str(ws_cm.cell(r, 2).value or '').strip()
        if code in cm_vung4_map:
            formula, note = cm_vung4_map[code]
            ws_cm.cell(r, 7, formula)
            ws_cm.cell(r, 8, f"=ROUND(F{r}*G{r}, 0)")
            ws_cm.cell(r, 9, note)
            count_cm += 1

print(f"    [OK] Đã cập nhật Đơn giá Ca máy VÙNG IV cho {count_cm}/21 Master row.")

print(">>> [3/4] Đang lưu file và kiểm tra...")
wb.save(PROJECT_PATH)
wb.save(DOWNLOADS_PATH)
file_size_mb = os.path.getsize(PROJECT_PATH) / (1024 * 1024)
print(f"    [OK] Đã lưu file: {PROJECT_PATH} ({file_size_mb:.2f} MB)")
print(f"    [OK] Đã lưu bản sao sang: {DOWNLOADS_PATH}")

print(">>> [4/4] Hoàn tất cập nhật Đơn giá Nhân công và Ca máy sang VÙNG IV!")
