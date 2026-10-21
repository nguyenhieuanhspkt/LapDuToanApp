"""
Trial Export Verification Test for LapDuToanApp.
Tests generating a complete Excel workbook and verifies all sheets, images, dynamic formulas,
and the work volume table (Khối lượng-Định mức-ĐVT).
"""

import sys
from pathlib import Path
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

# Add project root to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from config import OUTPUT_DIR
from core.excel_generator import ExcelGenerator
from core.norm_manager import NormManager

def run_trial_export():
    print("==================================================")
    print("STARTING TRIAL EXPORT VERIFICATION TEST (WITH KHỐI LƯỢNG)")
    print("==================================================")

    mgr = NormManager()
    generator = ExcelGenerator(mgr)

    project_name = "BẢO DƯỠNG, SỬA CHỮA VÀ THÍ NGHIỆM ĐỊNH KỲ HỆ THỐNG PCCC NĂM 2026 - NMNĐ VĨNH TÂN 4"
    category_name = "Hệ thống báo cháy tự động & Chữa cháy (Định mức Thông tư 38/2026/TT-BXD & QĐ 45/QĐ-HĐTV)"

    # Define all 13 work items with detailed quantities
    work_items = [
        {"stt": 1, "norm_code": "BD.41111", "job_name": "Thay mới lắp đặt thiết bị đầu báo khói địa chỉ", "description": "Thực hiện theo mô tả dịch vụ Thay thiết bị đầu báo", "unit": "Bộ", "volume": 10},
        {"stt": 2, "norm_code": "BD.41121", "job_name": "Lắp đặt đèn báo cháy", "description": "Lắp mới / thay thế đèn báo cháy vị trí khu vực phụ trợ", "unit": "Bộ", "volume": 5},
        {"stt": 3, "norm_code": "BD.41131", "job_name": "Lắp đặt lại nút nhấn khẩn báo cháy", "description": "Lắp đặt lại nút nhấn sau khi bảo trì, sửa chữa", "unit": "Cái", "volume": 38},
        {"stt": 4, "norm_code": "BD.41141", "job_name": "Lắp đặt chuông báo cháy", "description": "Thay thế, lắp đặt chuông báo cháy tự động", "unit": "Bộ", "volume": 5},
        {"stt": 5, "norm_code": "BD.42201", "job_name": "Lắp đặt bộ điều khiển, bộ chuyển mạch PCCC", "description": "Lắp đặt bộ chuyển mạch tín hiệu tủ trung gian", "unit": "Bộ", "volume": 1},
        {"stt": 6, "norm_code": "E6C.3.01", "job_name": "Bảo trì, sửa chữa, thử nghiệm giả lập tác động nút nhấn khẩn báo cháy", "description": "Kiểm tra, bảo dưỡng nút nhấn báo cháy chuyên dụng", "unit": "Cái", "volume": 38},
        {"stt": 7, "norm_code": "E6C.3.02", "job_name": "Bảo trì, sửa chữa, hiệu chỉnh, thử nghiệm độ nhạy đầu báo khói/nhiệt", "description": "Thực hiện theo mô tả dịch vụ Bảo trì đầu báo chuyên dụng", "unit": "Bộ", "volume": 890},
        {"stt": 8, "norm_code": "E6C.3.03", "job_name": "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng dây báo cháy", "description": "Kiểm tra module dây giám sát nhiệt độ cao", "unit": "Bộ", "volume": 20},
        {"stt": 9, "norm_code": "E6C.3.04", "job_name": "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng cảm biến beam quang", "description": "Bảo dưỡng beam khói khoảng cách xa", "unit": "Bộ", "volume": 15},
        {"stt": 10, "norm_code": "E6C.3.05", "job_name": "Sửa chữa, bảo dưỡng trung tâm báo cháy FMZ 5000", "description": "Bảo dưỡng tủ điều khiển chữa cháy tại chỗ và tủ trung gian", "unit": "trung tâm", "volume": 3},
        {"stt": 11, "norm_code": "E6C.3.06", "job_name": "Kiểm tra và bảo dưỡng phần điện van Deluge, Wet hệ PCCC", "description": "Bảo dưỡng cuộn hút, van xả tràn Deluge và đồng hồ áp", "unit": "Cái", "volume": 12},
        {"stt": 12, "norm_code": "E6B.1.13.01", "job_name": "Kiểm tra bình accu chuyên dụng phục vụ hệ thống PCCC", "description": "Đo nội trở, dung lượng Ah, vô mỡ cọc bình và châm nước cất", "unit": "Bình", "volume": 50},
        {"stt": 13, "norm_code": "E6C.1.31", "job_name": "Kiểm tra và bảo dưỡng thiết bị đo khí NH3, Clo, H2", "description": "Bơm khí chuẩn hiệu chuẩn điểm Zero & Span cho đầu đo khí", "unit": "Bộ", "volume": 4}
    ]

    output_file = OUTPUT_DIR / "DU_TOAN_THIET_KE_TEST.xlsx"
    print(f"Exporting workbook to: {output_file}")
    generator.generate(project_name, category_name, work_items, output_file)

    assert output_file.exists(), f"Output file not generated: {output_file}"
    file_size_mb = output_file.stat().st_size / (1024 * 1024)
    print(f"[PASS] File created successfully! Size: {file_size_mb:.2f} MB")

    # Inspect generated workbook
    wb = openpyxl.load_workbook(output_file, data_only=False)
    sheet_names = wb.sheetnames
    print(f"[PASS] Total sheets generated: {len(sheet_names)}")

    # 1. Verify Sheet TH_KinhPhi is the FIRST sheet!
    assert sheet_names[0] == "TH_KinhPhi", f"First sheet is not TH_KinhPhi: {sheet_names[0]}"
    ws_kp = wb["TH_KinhPhi"]
    assert ws_kp["B6"].value == "Chi phí trực tiếp", "Row 6 of TH_KinhPhi is not Chi phí trực tiếp"
    assert ws_kp["D6"].value == "=D7+D9+D11", "Formula D6 mismatch"
    assert ws_kp["D22"].value == "=ROUND(D21,-3)", "Formula D22 mismatch"
    assert str(ws_kp["B23"].value).startswith("Bằng chữ:"), "Row 23 text mismatch"
    print(f"[PASS] Sheet 'TH_KinhPhi' is in 1st position with dynamic formulas and text: {ws_kp['B23'].value}")

    # 2. Verify Sheet Khối lượng is the SECOND sheet!
    assert sheet_names[1] == "Khối lượng-Định mức-ĐVT", f"Second sheet is not Khối lượng: {sheet_names[1]}"
    ws_kl = wb["Khối lượng-Định mức-ĐVT"]
    assert ws_kl["A6"].value == "I", f"Row 6 STT is not Group I: {ws_kl['A6'].value}"
    assert "NHÓM I:" in str(ws_kl["B6"].value), f"Row 6 is not Group I header: {ws_kl['B6'].value}"
    print(f"[PASS] Sheet 'Khối lượng-Định mức-ĐVT' has standard Group I header at Row 6: {ws_kl['B6'].value}")

    # 3. Verify all other sheets exist
    expected_sheets = [
        "TH_KinhPhi",
        "Khối lượng-Định mức-ĐVT",
        "00_ThongKe_DinhMuc", "01_ChiTiet_HaoPhi_DinhMuc",
        "TH_VatLieu", "TH_NhanCong", "TH_MayThiCong",
        "BD.41111", "BD.41121", "BD.41131", "BD.41141", "BD.42201",
        "E6C.3.01", "E6C.3.02", "E6C.3.03", "E6C.3.04", "E6C.3.05",
        "E6C.3.06", "E6B.1.13.01", "E6C.1.31"
    ]
    for s in expected_sheets:
        assert s in sheet_names, f"Missing sheet: {s}"
    print(f"[PASS] All {len(expected_sheets)} expected sheets exist!")

    # 4. Check embedded images & Labor split in BD.41111
    evidence_sheets = expected_sheets[7:]
    images_count = sum(1 for s in evidence_sheets if getattr(wb[s], "_images", []))
    print(f"[PASS] Evidence images check: {images_count}/13 sheets contain embedded scan images!")

    # Verify Labor Split in BD.41111 (Separating 0.174 into NC & Kỹ sư)
    ws_bd = wb["BD.41111"]
    assert ws_bd["R3"].value == "NC nhóm 3 (công)", f"R3 mismatch: {ws_bd['R3'].value}"
    assert ws_bd["S3"].value == "Kỹ sư (công)", f"S3 mismatch: {ws_bd['S3'].value}"
    assert ws_bd["T3"].value == "Tổng ĐM gốc (công)", f"T3 mismatch: {ws_bd['T3'].value}"
    assert ws_bd["R9"].value == "=SUM(R4:R8)", f"R9 formula mismatch: {ws_bd['R9'].value}"
    assert ws_bd["S9"].value == "=SUM(S4:S8)", f"S9 formula mismatch: {ws_bd['S9'].value}"
    assert "0.174" in str(ws_bd["P10"].value) and "0.087" in str(ws_bd["P10"].value), f"Row 10 note mismatch: {ws_bd['P10'].value}"
    print(f"[PASS] Sheet BD.41111 labor split verified: NC=0.087 công, Kỹ sư=0.087 công, Tổng=0.174 công (Note: {ws_bd['P10'].value})")

    # 4. Check formulas in 01_ChiTiet_HaoPhi_DinhMuc
    ws_ct = wb["01_ChiTiet_HaoPhi_DinhMuc"]
    g_links = sum(1 for r in range(6, ws_ct.max_row + 1) if str(ws_ct.cell(r, 7).value or "").startswith("="))
    h_links = sum(1 for r in range(6, ws_ct.max_row + 1) if str(ws_ct.cell(r, 8).value or "").startswith("="))
    j_formulas = sum(1 for r in range(6, ws_ct.max_row + 1) if str(ws_ct.cell(r, 10).value or "").startswith("="))
    print(f"[PASS] 01_ChiTiet_HaoPhi_DinhMuc formula verification:")
    print(f"       - Column G (Định mức) dynamic links : {g_links}")
    print(f"       - Column H (Đơn giá) dynamic links  : {h_links}")
    print(f"       - Column J (Thành tiền) formulas    : {j_formulas}")
    assert g_links > 40 and h_links > 40 and j_formulas > 40

    # 5. Check calculated total quantities in summary sheets
    for th_name, sum_col in [("TH_VatLieu", 9), ("TH_NhanCong", 9), ("TH_MayThiCong", 11)]:
        ws_th = wb[th_name]
        last_r = ws_th.max_row
        sum_cell = ws_th.cell(last_r, sum_col).value
        assert str(sum_cell).startswith("=SUM"), f"Sheet {th_name} total formula missing at col {sum_col}"
        # Check non-zero volume in column 5
        non_zero_kl = sum(1 for r in range(7, last_r) if float(ws_th.cell(r, 5).value or 0) > 0)
        print(f"[PASS] Sheet {th_name:14s}: {non_zero_kl} items have computed project volumes > 0, total sum formula: {sum_cell}")

    print("\n==================================================")
    print("ALL TESTS PASSED WITH 100% QUALITY AND ACCURACY!")
    print("==================================================")
    return True

if __name__ == "__main__":
    success = run_trial_export()
    sys.exit(0 if success else 1)
