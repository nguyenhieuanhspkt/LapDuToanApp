# -*- coding: utf-8 -*-
"""
Script update_nc_group_pricing.py
Cập nhật đơn giá nhân công theo đúng yêu cầu:
"chỉ cần áp dụng theo nhóm nhân công, ví dụ nhóm 3, vùng IV, Lâm Đồng là 312.490 đồng/công, không cần phân theo bậc. tương tự đối với Khánh Hòa"

1. Đối với Vĩnh Tân 4 (Lâm Đồng Vùng IV):
   - Nhóm 3: 312.490 đ/công -> Link ='01_DonGia_NhanCong'!G9
   - Nhóm 4: 328.116 đ/công -> Link ='01_DonGia_NhanCong'!G10
   - Kỹ sư KS-TN: 313.000 đ/công -> Link ='01_DonGia_NhanCong'!G19

2. Đối với Phước Thái 1, 2, 3 (Khánh Hòa KV3 - Ninh Phước):
   - Nhóm 3: 286.151 đ/công -> Link ='KH_01_DonGia_NhanCong'!F9
   - Nhóm 4: 307.450 đ/công -> Link ='KH_01_DonGia_NhanCong'!F10
   - Kỹ sư KS-TN: 296.438 đ/công -> Link ='KH_01_DonGia_NhanCong'!F18
"""

import os
import sys
import io
import shutil
import openpyxl

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

PROJECT_PATH = r"D:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
DOWNLOADS_PATH = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx"
ALT_DOWNLOADS = r"C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_MOI.xlsx"

print(">>> Đang nạp workbook...")
wb = openpyxl.load_workbook(PROJECT_PATH)
ws = wb['Nhân công']

group_map = {
    # Nhóm 3:
    'NC_E6C.3.05': 3,
    'NC_E6C.3.02_B5': 3,
    'NC_E6C.3.02_B1-11': 3,
    'NC_BD.41111': 3,
    'NC_E6C.3.01_B5': 3,
    'NC_E6C.3.01_B1-11': 3,
    'NC_BD.41131': 3,
    'NC_BD.41141': 3,
    'NC_E6C.3.03': 3,
    'NC_BD.41121': 3,
    'NC_E6C.3.04': 3,
    'NC_E6C.3.06': 3,
    'NC_E6BOP.1.03': 3,
    'NC_SB.82633': 3,
    'NC_BB.86601': 3,
    'NC_BB.86103': 3,
    'NC_BB.85201': 3,
    'NC_SB.82634': 3,
    # Nhóm 4:
    'NC_E6C.1.23': 4,
    'NC_E6C.1.31': 4,
    'NC_E6C.2.04': 4,
    'NC_E6C.1.18': 4,
    'NC_BA.18405': 4,
    'NC_E6B.1.13.01': 4,
    'NC_E6BOP.1.01': 4,
    'NC_BD.41171': 4,
    # Kỹ sư:
    'NC_DB.24001': 'KS'
}

updated_count = 0
for r in range(6, ws.max_row + 1):
    c1 = str(ws.cell(r, 1).value or '')
    c2 = str(ws.cell(r, 2).value or '')
    c3 = str(ws.cell(r, 3).value or '')
    c6 = str(ws.cell(r, 6).value or '')
    
    if c6.startswith('=SUM'):
        grp = group_map.get(c2, 3)
        is_pt = 'PHƯỚC THÁI' in c3.upper()
        
        if is_pt:
            # Khánh Hòa KV3 Ninh Phước
            if grp == 3:
                new_formula = "='KH_01_DonGia_NhanCong'!F9"
                note = "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 3 KV3 Ninh Phước - Đơn giá theo nhóm: 286.151 đ) - Link 'KH_01_DonGia_NhanCong'!F9"
            elif grp == 4:
                new_formula = "='KH_01_DonGia_NhanCong'!F10"
                note = "QĐ 646/QĐ-SXD Khánh Hòa (Nhóm 4 KV3 Ninh Phước - Đơn giá theo nhóm: 307.450 đ) - Link 'KH_01_DonGia_NhanCong'!F10"
            else:
                new_formula = "='KH_01_DonGia_NhanCong'!F18"
                note = "QĐ 646/QĐ-SXD Khánh Hòa (Kỹ sư KS-TN KV3 Ninh Phước - Đơn giá theo nhóm: 296.438 đ) - Link 'KH_01_DonGia_NhanCong'!F18"
        else:
            # Lâm Đồng Vùng IV
            if grp == 3:
                new_formula = "='01_DonGia_NhanCong'!G9"
                note = "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 3 VÙNG IV - Đơn giá theo nhóm: 312.490 đ) - Link '01_DonGia_NhanCong'!G9"
            elif grp == 4:
                new_formula = "='01_DonGia_NhanCong'!G10"
                note = "VB 5898/SXD-QLXD Lâm Đồng (Nhóm 4 VÙNG IV - Đơn giá theo nhóm: 328.116 đ) - Link '01_DonGia_NhanCong'!G10"
            else:
                new_formula = "='01_DonGia_NhanCong'!G19"
                note = "VB 5898/SXD-QLXD Lâm Đồng (Kỹ sư KS-TN VÙNG IV - Đơn giá theo nhóm: 313.000 đ) - Link '01_DonGia_NhanCong'!G19"
                
        ws.cell(r, 7, new_formula)
        ws.cell(r, 9, note)
        updated_count += 1
        print(f"Row {r:3d} | STT: {c1:5s} | Code: {c2:18s} | Giá mới: {new_formula:30s} | {note[:60]}...")

print(f">>> Đã cập nhật xong {updated_count} Master rows trong sheet 'Nhân công'.")

print(">>> Đang lưu file chính...")
wb.save(PROJECT_PATH)
print(">>> [OK] Đã lưu:", PROJECT_PATH)

print(">>> Đang sao chép sang Downloads...")
try:
    shutil.copy2(PROJECT_PATH, DOWNLOADS_PATH)
    print(">>> [OK] Đã copy sang:", DOWNLOADS_PATH)
except PermissionError:
    shutil.copy2(PROJECT_PATH, ALT_DOWNLOADS)
    print(">>> [OK] Đã copy sang:", ALT_DOWNLOADS)
