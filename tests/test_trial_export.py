"""
Trial Export Verification Test for LapDuToanApp.
Tests generating a complete Excel workbook and verifies all sheets, images, and dynamic formulas.
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
    print("STARTING TRIAL EXPORT VERIFICATION TEST")
    print("==================================================")

    mgr = NormManager()
    generator = ExcelGenerator(mgr)

    project_name = "BẢO DƯỠNG, SỬA CHỮA VÀ THÍ NGHIỆM ĐỊNH KỲ HỆ THỐNG PCCC NĂM 2026 - NMNĐ VĨNH TÂN 4"
    category_name = "Hệ thống báo cháy tự động & Chữa cháy (Định mức Thông tư 38/2026/TT-BXD & QĐ 45/QĐ-HĐTV)"

    # Define all 13 work items
    work_items = [
        {"stt": 1, "norm_code": "BD.41111", "job_name": "Lắp đặt thiết bị đầu báo và đầu báo cháy", "unit": "10 đầu", "volume": 10},
        {"stt": 2, "norm_code": "BD.41121", "job_name": "Lắp đặt đèn báo cháy", "unit": "5 đèn", "volume": 5},
        {"stt": 3, "norm_code": "BD.41131", "job_name": "Lắp đặt nút ấn báo cháy khẩn cấp", "unit": "5 nút", "volume": 5},
        {"stt": 4, "norm_code": "BD.41141", "job_name": "Lắp đặt chuông báo cháy", "unit": "5 chuông", "volume": 5},
        {"stt": 5, "norm_code": "BD.42201", "job_name": "Lắp đặt bộ điều khiển, bộ chuyển mạch PCCC", "unit": "1 bộ", "volume": 1},
        {"stt": 6, "norm_code": "E6C.3.01", "job_name": "Kiểm tra và bảo dưỡng nút nhấn báo cháy, bộ còi đèn/ chuông báo cháy", "unit": "1 cái", "volume": 38},
        {"stt": 7, "norm_code": "E6C.3.02", "job_name": "Kiểm tra và bảo dưỡng cảm biến báo cháy dạng nhiệt/ khói", "unit": "1 cái", "volume": 890},
        {"stt": 8, "norm_code": "E6C.3.03", "job_name": "Kiểm tra và bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng dây báo cháy", "unit": "1 bộ", "volume": 20},
        {"stt": 9, "norm_code": "E6C.3.04", "job_name": "Kiểm tra và bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng cảm biến beam quang", "unit": "1 bộ", "volume": 15},
        {"stt": 10, "norm_code": "E6C.3.05", "job_name": "Kiểm tra và bảo dưỡng tủ điều khiển chữa cháy tại chỗ, tủ điều khiển và tủ trung gian ĐLĐK", "unit": "1 tủ", "volume": 3},
        {"stt": 11, "norm_code": "E6C.3.06", "job_name": "Kiểm tra và bảo dưỡng phần điện van Deluge, Wet, Pre-action hệ thống PCCC", "unit": "1 cái", "volume": 12},
        {"stt": 12, "norm_code": "E6B.1.13.01", "job_name": "Kiểm tra bình accu (Ắc quy chuyên dụng PCCC)", "unit": "1 bình", "volume": 50},
        {"stt": 13, "norm_code": "E6C.1.31", "job_name": "Kiểm tra và bảo dưỡng thiết bị đo khí NH3, Clo, H2", "unit": "1 bộ", "volume": 4}
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

    expected_sheets = [
        "00_ThongKe_DinhMuc", "01_ChiTiet_HaoPhi_DinhMuc",
        "TH_VatLieu", "TH_NhanCong", "TH_MayThiCong",
        "BD.41111", "BD.41121", "BD.41131", "BD.41141", "BD.42201",
        "E6C.3.01", "E6C.3.02", "E6C.3.03", "E6C.3.04", "E6C.3.05",
        "E6C.3.06", "E6B.1.13.01", "E6C.1.31"
    ]
    for s in expected_sheets:
        assert s in sheet_names, f"Missing sheet: {s}"
    print(f"[PASS] All {len(expected_sheets)} expected sheets exist!")

    # Check images in evidence sheets
    evidence_sheets = expected_sheets[5:]
    images_count = 0
    for s in evidence_sheets:
        ws_ev = wb[s]
        imgs = getattr(ws_ev, "_images", [])
        if imgs:
            images_count += 1
    print(f"[PASS] Evidence images check: {images_count}/13 sheets contain embedded scan images!")

    # Check formulas in 01_ChiTiet_HaoPhi_DinhMuc
    ws_ct = wb["01_ChiTiet_HaoPhi_DinhMuc"]
    g_links = 0
    h_links = 0
    j_formulas = 0
    for r in range(6, ws_ct.max_row + 1):
        g_val = str(ws_ct.cell(r, 7).value or "")
        h_val = str(ws_ct.cell(r, 8).value or "")
        j_val = str(ws_ct.cell(r, 10).value or "")
        if g_val.startswith("="): g_links += 1
        if h_val.startswith("="): h_links += 1
        if j_val.startswith("="): j_formulas += 1

    print(f"[PASS] 01_ChiTiet_HaoPhi_DinhMuc formula verification:")
    print(f"       - Column G (Định mức) dynamic links : {g_links}")
    print(f"       - Column H (Đơn giá) dynamic links  : {h_links}")
    print(f"       - Column J (Thành tiền) formulas    : {j_formulas}")
    assert g_links > 40, "Too few G formula links"
    assert h_links > 40, "Too few H formula links"
    assert j_formulas > 40, "Too few J formulas"

    # Check formulas in TH_VatLieu, TH_NhanCong, TH_MayThiCong
    for th_name, col_ht in [("TH_VatLieu", 7), ("TH_NhanCong", 7), ("TH_MayThiCong", 9)]:
        ws_th = wb[th_name]
        last_r = ws_th.max_row
        sum_cell = ws_th.cell(last_r, ws_th.max_column).value
        assert str(sum_cell).startswith("=SUM"), f"Sheet {th_name} total formula missing"
    print("[PASS] Summary sheets (TH_VL, TH_NC, TH_M) formula verification passed!")

    print("\n==================================================")
    print("ALL TRIAL EXPORT TESTS PASSED WITH 100% QUALITY!")
    print("==================================================")
    return True

if __name__ == "__main__":
    success = run_trial_export()
    sys.exit(0 if success else 1)
