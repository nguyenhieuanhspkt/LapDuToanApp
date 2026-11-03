# -*- coding: utf-8 -*-
"""
Module: core.constants
Mục đích: Tập trung toàn bộ hằng số cấu hình hệ thống:
- Tên các Sheet trong Workbook
- Định dạng số (Currency, Rate, Quantity)
- Tọa độ và chỉ số cột Excel
- Các mã hiệu đặc thù (Vật tư phụ %, Vật tư thay thế, Tiền tố báo giá)

Tác giả: Tổ Thẩm Định EVN
"""

# ==============================================================================
# 1. TÊN CÁC SHEET CHUẨN TRONG WORKBOOK HỒ SƠ DỰ TOÁN
# ==============================================================================
SHEET_TH_KINH_PHI = "TH_KinhPhi"
SHEET_KHOI_LUONG = "Khối lượng-Định mức-ĐVT"
SHEET_DOI_CHIEU_DINH_MUC = "Doi_chieu_dinh_muc"
SHEET_THONG_KE_DINH_MUC = "00_ThongKe_DinhMuc"
SHEET_CHI_TIET_HAO_PHI = "01_ChiTiet_HaoPhi_DinhMuc"
SHEET_GIAI_TRINH_HE_SO = "02_GiaiTrinh_HeSo_DieuChinh"
SHEET_TH_VAT_LIEU = "TH_VatLieu"
SHEET_TH_NHAN_CONG = "TH_NhanCong"
SHEET_TH_MAY_THI_CONG = "TH_MayThiCong"

ORDERED_STANDARD_SHEET_NAMES = [
    SHEET_TH_KINH_PHI,
    SHEET_KHOI_LUONG,
    SHEET_DOI_CHIEU_DINH_MUC,
    SHEET_THONG_KE_DINH_MUC,
    SHEET_CHI_TIET_HAO_PHI,
    SHEET_GIAI_TRINH_HE_SO,
    SHEET_TH_VAT_LIEU,
    SHEET_TH_NHAN_CONG,
    SHEET_TH_MAY_THI_CONG,
]

# ==============================================================================
# 2. ĐỊNH DẠNG SỐ (NUMBER FORMATS)
# ==============================================================================
FORMAT_CURRENCY = "#,##0"        # Thành tiền, đơn giá (VNĐ)
FORMAT_RATE = "#,##0.000"       # Định mức hao phí thông thường
FORMAT_RATE_4DEC = "0.0000"     # Định mức chi tiết (tỷ lệ phân rã bước công việc)
FORMAT_QUANTITY = "#,##0.00"    # Khối lượng công tác

# ==============================================================================
# 3. TỌA ĐỘ VÀ CỘT EXCEL TRONG SHEET 01_ChiTiet_HaoPhi_DinhMuc
# ==============================================================================
COL_STT = 1         # Cột A: Số thứ tự
COL_CODE = 2        # Cột B: Mã hiệu đơn giá / mã tài nguyên
COL_NAME = 5        # Cột E: Tên công tác / tên tài nguyên hao phí
COL_UNIT = 6        # Cột F: Đơn vị tính
COL_RATE = 7        # Cột G: Định mức hao phí
COL_PRICE = 8       # Cột H: Đơn giá tài nguyên
COL_COEFF = 9       # Cột I: Hệ số điều chỉnh
COL_AMOUNT = 10     # Cột J: Thành tiền (G * H)

LET_COL_RATE = "G"
LET_COL_PRICE = "H"
LET_COL_AMOUNT = "J"

# Cột đơn giá trong các Sheet Tổng hợp giá tài nguyên
LET_COL_PRICE_TH_VAT_LIEU = "G"
LET_COL_PRICE_TH_NHAN_CONG = "F"
LET_COL_PRICE_TH_MAY = "F"

# Cột định mức trong các Sheet Minh chứng (Evidence Sheet)
LET_COL_EVIDENCE_RATE = "R"

# ==============================================================================
# 4. MÃ HIỆU ĐẶC THÙ VÀ QUY TẮC PHÂN LOẠI
# ==============================================================================
# Mã vật tư phụ tính theo tỷ lệ phần trăm (%) trên tổng vật tư chính
PERCENTAGE_MATERIAL_CODE = "V00750"

# Nhóm mã vật tư thay thế (Đơn giá = 0 khi dự án chỉ mua dịch vụ bảo dưỡng)
REPLACEMENT_MATERIAL_CODES = {"V82782", "V82780", "V82781", "V82783"}

# Tiền tố mã định danh công tác dạng Báo giá / Tạm tính
QUOTE_CODE_PREFIXES = ("BG", "TT", "VT")
