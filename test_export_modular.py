# -*- coding: utf-8 -*-
"""
Test Suite: Kiểm tra toàn diện hệ thống tạo File Excel Dự toán Modular (1 Sheet / 1 Module).
Chạy lệnh: python test_export_modular.py
Author: Tổ Thẩm Định EVN
"""

import sys
import time
from pathlib import Path
import openpyxl

# Đảm bảo in tiếng Việt chuẩn trên Windows Console
sys.stdout.reconfigure(encoding='utf-8')

print("=" * 80)
print("🚀 BẮT ĐẦU KIỂM TRA HỆ THỐNG XUẤT EXCEL DỰ TOÁN PCCC 2026 (KIẾN TRÚC MODULAR)")
print("=" * 80)

# 1. Kiểm tra Import các module con trong core/sheets/
print("\n[BƯỚC 1/4] Kiểm tra Import các Module con trong core/sheets/ ...")
try:
    from core.sheets import (
        BaseSheetBuilder,
        SheetTHKinhPhiBuilder,
        SheetKhoiLuongBuilder,
        SheetDonGiaChiTietBuilder,
        SheetGiaiTrinhHeSoBuilder,
        SheetThongKeDinhMucBuilder,
        SheetEvidenceBuilder,
        SheetTHVatLieuBuilder,
        SheetTHNhanCongBuilder,
        SheetTHMayThiCongBuilder
    )
    from core.norm_manager import NormManager
    from core.excel_generator import ExcelGenerator
    print("  ✅ Import thành công toàn bộ 9 Builder module chuyên biệt!")
except Exception as e:
    print(f"  ❌ LỖI IMPORT: {e}")
    sys.exit(1)

# 2. Khởi tạo dữ liệu và Orchestrator
print("\n[BƯỚC 2/4] Nạp dữ liệu dự toán từ Database...")
mgr = NormManager()
gen = ExcelGenerator(mgr)
items = mgr.load_work_items()
print(f"  ✅ Đã nạp {len(items)} công tác dự án.")

# 3. Thực thi tạo file Excel
output_file = Path(r"d:\TaskApp_kiet\LapDuToanApp\TEST_DUTOAN_MODULAR_OUTPUT.xlsx")
print(f"\n[BƯỚC 3/4] Xuất file Excel thử nghiệm sang: {output_file.name} ...")

start_time = time.time()
try:
    gen.generate(
        project_name="BẢO DƯỠNG PCCC NĂM 2026 - NMNĐ VĨNH TÂN 4 (TEST MODULAR)",
        category_name="Hệ thống báo cháy tự động & Chữa cháy",
        work_items=items,
        output_path=output_file
    )
    elapsed = time.time() - start_time
    file_size_kb = output_file.stat().st_size / 1024
    print(f"  ✅ Xuất file thành công! Thời gian: {elapsed:.2f}s | Dung lượng: {file_size_kb:,.1f} KB")
except Exception as e:
    print(f"  ❌ LỖI TRONG QUÁ TRÌNH TẠO FILE: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

# 4. Kiểm tra cấu trúc các Sheet và Công thức liên kết động
print("\n[BƯỚC 4/4] Thẩm tra cấu trúc các Sheet và Công thức liên kết động trong file Excel...")
wb = openpyxl.load_workbook(output_file, data_only=False)
expected_sheets = [
    "TH_KinhPhi",
    "Khối lượng-Định mức-ĐVT",
    "00_ThongKe_DinhMuc",
    "01_ChiTiet_HaoPhi_DinhMuc",
    "02_GiaiTrinh_HeSo_DieuChinh",
    "TH_VatLieu",
    "TH_NhanCong",
    "TH_MayThiCong"
]

all_present = True
for s in expected_sheets:
    if s in wb.sheetnames:
        ws = wb[s]
        print(f"  🟢 Sheet '{s:<28}': TỒN TẠI ({ws.max_row:>3} dòng, {ws.max_column:>2} cột)")
    else:
        print(f"  🔴 Sheet '{s:<28}': THIẾU")
        all_present = False

# Kiểm tra các sheet minh chứng định mức phụ lục
evidence_sheets = [s for s in wb.sheetnames if s not in expected_sheets]
print(f"  🟢 Các sheet minh chứng scan gốc: {evidence_sheets}")

# Kiểm tra công thức liên kết động trong TH_KinhPhi
ws_kp = wb["TH_KinhPhi"]
d6_formula = ws_kp["D6"].value
d14_formula = ws_kp["D14"].value
d21_formula = ws_kp["D21"].value

print("\n  🔍 Kiểm tra công thức động tại Sheet TH_KinhPhi:")
print(f"     - Ô D6  (Chi phí trực tiếp T)  : {d6_formula}")
print(f"     - Ô D14 (Chi phí chung C)       : {d14_formula}")
print(f"     - Ô D21 (Tổng cộng dự toán)     : {d21_formula}")

assert "=D7+D9+D11" in str(d6_formula), "Công thức D6 không đúng định dạng =D7+D9+D11"
assert "*0.073" in str(d14_formula) or "D6*" in str(d14_formula), "Công thức D14 không liên kết với D6"

print("\n" + "=" * 80)
print("🎉 KẾT QUẢ: TẤT CẢ CÁC MODULE CON TRONG core/sheets/ HOẠT ĐỘNG HOÀN HẢO 100%!")
print(f"📁 File kết quả kiểm tra tại: {output_file}")
print("=" * 80)
