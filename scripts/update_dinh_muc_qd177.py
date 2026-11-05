import openpyxl

wb_path = 'BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
wb = openpyxl.load_workbook(wb_path)
ws_dm = wb['Định mức áp dụng']

print("Updating sheet 'Định mức áp dụng' with QD 177/QD-HDTV...")

# Update Section II Title
ws_dm.cell(45, 1, "MỤC II. CHI TIẾT CÁC ĐỊNH MỨC SỬA CHỮA LỚN EVN (QĐ 45/QĐ-HĐTV & CHUYỂN ĐỔI NHÓM NHÂN CÔNG THEO QĐ 177/QĐ-HĐTV NGÀY 29/10/2026)")

# Update SCL Norm rows with Group mapping
# E6C.3.05 -> Nhóm 3
ws_dm.cell(61, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(61, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 451 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6C.3.01 -> Nhóm 3
ws_dm.cell(79, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(79, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 448 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6C.3.02 -> Nhóm 3
ws_dm.cell(97, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(97, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 449 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6C.3.04 -> Nhóm 3
ws_dm.cell(115, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(115, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 451 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6C.3.03 -> Nhóm 3
ws_dm.cell(134, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(134, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 450 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6C.3.06 -> Nhóm 3
ws_dm.cell(154, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(154, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 453 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6C.1.31 -> Nhóm 4
ws_dm.cell(173, 2, "      • Nhân công kỹ thuật: Nhóm 4 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(173, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 436 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 4)")

# E6B.1.13.01 -> Nhóm 4
ws_dm.cell(191, 2, "      • Nhân công kỹ thuật: Nhóm 4 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(191, 6, "QĐ 45/QĐ-HĐTV Tập 5 trang 383 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 4)")

# E6C.2.04 -> Nhóm 4
ws_dm.cell(197, 2, "      • Nhân công kỹ thuật: Nhóm 4 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(197, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 440 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 4)")

# E6C.1.23 -> Nhóm 4
ws_dm.cell(201, 2, "      • Nhân công kỹ thuật: Nhóm 4 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(201, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 428 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 4)")

# E6C.1.18 -> Nhóm 4
ws_dm.cell(205, 2, "      • Nhân công kỹ thuật: Nhóm 4 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(205, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 425 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 4)")

# E6BOP.1.03 -> Nhóm 3
ws_dm.cell(209, 2, "      • Nhân công kỹ thuật: Nhóm 3 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(209, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 312 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 3)")

# E6BOP.1.01 -> Nhóm 4
ws_dm.cell(213, 2, "      • Nhân công kỹ thuật: Nhóm 4 - Bậc 5,0/7 (Theo QĐ 177/QĐ-HĐTV & QĐ 45)")
ws_dm.cell(213, 6, "QĐ 45/QĐ-HĐTV Tập 6 trang 308 & QĐ 177/QĐ-HĐTV ngày 29/10/2026 (Chuyển đổi sang Nhóm 4)")

wb.save(wb_path)
print("Saved sheet 'Định mức áp dụng' successfully.")
