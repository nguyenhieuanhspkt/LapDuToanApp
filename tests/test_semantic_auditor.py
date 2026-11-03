"""
Unit & Integration Test for Semantic Auditor and Excel Exporter
Kiểm tra Module Tiền Thẩm Tra & Bóc Tách Dữ Liệu Thô PAKT Độc Lập
"""

import sys
from pathlib import Path
import openpyxl

# Add project root to sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

from core.semantic_auditor import SemanticAuditor
from core.audit_excel_exporter import AuditExcelExporter

def test_semantic_pipeline():
    print("🚀 Bắt đầu kiểm thử Module Tiền Thẩm Tra Dữ Liệu Thô PAKT...")

    # Dữ liệu kiểm thử mẫu thực tế từ hồ sơ PAKT Vĩnh Tân 4
    sample_raw_items = [
        {
            "stt": 1,
            "job_name": "Thay mới lắp đặt thiết bị đầu báo khói địa chỉ",
            "description": "Đầu báo khói quang điện tử địa chỉ Hochiki ALN-V gian máy",
            "unit": "Bộ",
            "volume": 20.0
        },
        {
            "stt": 2,
            "job_name": "Bảo dưỡng, vệ sinh thổi bụi đầu báo khói",
            "description": "Vệ sinh buồng khói, xịt cồn bảo dưỡng định kỳ khu nhà phụ trợ",
            "unit": "Cái",
            "volume": 50.0
        },
        {
            "stt": 3,
            "job_name": "Kiểm tra, thử nghiệm độ nhạy đầu báo khói địa chỉ",
            "description": "Dùng bình tạo khói chuyên dụng thử kích hoạt tín hiệu loop",
            "unit": "Cái",
            "volume": 35.0
        },
        {
            "stt": 4,
            "job_name": "Bảo dưỡng hệ thống PCCC gian máy chính",
            "description": "Bảo dưỡng toàn bộ thiết bị PCCC khu vực tuabin máy phát",
            "unit": "Hệ",
            "volume": 1.0
        },
        {
            "stt": 5,
            "job_name": "Thí nghiệm tác động cuộn hút van Deluge xả tràn",
            "description": "Đo kiểm cuộn hút solenoid van tràn Deluge DN100 hệ chữa cháy Foam",
            "unit": "Cái",
            "volume": 6.0
        },
        {
            "stt": 6,
            "job_name": "Kiểm tra đo nội trở bình ắc quy dự phòng tủ PCCC",
            "description": "Đo nội trở, dung lượng Ah bình ắc quy 24VDC-100Ah trạm 110kV",
            "unit": "Bình",
            "volume": 12.0
        },
        {
            "stt": 7,
            "job_name": "Bảo dưỡng thiết bị phụ trợ linh tinh",
            "description": "Không có thông số cụ thể",
            "unit": "Gói",
            "volume": 1.0
        }
    ]

    auditor = SemanticAuditor()
    result = auditor.audit_raw_dataset(sample_raw_items)

    rows = result["audited_rows"]
    stats = result["summary_stats"]

    print(f"\n✅ Đã phân tích thành công {len(rows)} dòng dữ liệu thô.")
    print(f"📊 Thống kê chất lượng hồ sơ:")
    print(f" - 🟢 Đạt chuẩn (Rõ ràng): {stats['green_count']} dòng")
    print(f" - 🟡 Cần làm rõ điều kiện / Biến thể A1 / Quy đổi: {stats['yellow_count']} dòng")
    print(f" - 🔴 Mập mờ / Chưa đủ điều kiện (Cấm ép mã): {stats['red_count']} dòng")

    print("\n🔍 Chi tiết phân tích từng dòng:")
    for r in rows:
        print(f"\n[Dòng {r['stt']}] Tên thô: {r['raw_job_name']} | ĐVT: {r['raw_unit']}")
        print(f"   ├─ V (Hành động) : {r['v_action']}")
        print(f"   ├─ O (Đối tượng) : {r['o_object']}")
        print(f"   ├─ S (Quy cách)  : {r['s_spec']}")
        print(f"   ├─ C (Bối cảnh)  : {r['c_context']}")
        print(f"   ├─ Trạng thái    : {r['audit_status']}")
        print(f"   ├─ CẦN LÀM RÕ    : {r['can_lam_ro']}")
        print(f"   ├─ Mã ĐM đề xuất : {r['prop_norm_code']} ({r['prop_norm_name']})")
        print(f"   ├─ ĐVT ĐM gốc    : {r['prop_norm_unit']} | Hệ số quy đổi: {r['conv_factor']} | KL quy đổi: {r['converted_vol']}")
        print(f"   ├─ Hao phí NC    : {r['prop_labor_rate']}")
        print(f"   └─ Cụm lặp lại   : {r['cluster_id']} (Tần suất xuất hiện: {r['cluster_count']} dòng)")

    # Assertions
    # Dòng 4 (Hệ) và Dòng 7 (Gói) phải là RED
    assert rows[3]["audit_level"] == "RED", "Dòng 4 ĐVT 'Hệ' phải bị gắn nhãn RED!"
    assert rows[6]["audit_level"] == "RED", "Dòng 7 ĐVT 'Gói' phải bị gắn nhãn RED!"
    
    # Dòng 2 & Dòng 3 là bảo dưỡng đầu báo -> Phải là YELLOW (cần làm rõ biến thể / quy đổi)
    assert rows[1]["audit_level"] == "YELLOW", "Dòng 2 bảo dưỡng đầu báo phải là YELLOW!"
    
    # Dòng 1 thay mới đầu báo -> Hệ số quy đổi từ Cái/Bộ sang 10 đầu phải là 0.1
    assert rows[0]["conv_factor"] == 0.1, "Dòng 1 quy đổi ĐVT 10 đầu phải là 0.1!"
    assert rows[0]["converted_vol"] == 2.0, "Dòng 1 khối lượng 20 bộ quy đổi phải là 2.0 (10 đầu)!"

    # Dòng 5 (Van Deluge) và Dòng 6 (Bình ắc quy) phải có đầy đủ hao phí nhân công
    assert "công" in rows[4]["prop_labor_rate"], "Dòng 5 phải có hao phí nhân công!"
    assert "công" in rows[5]["prop_labor_rate"], "Dòng 6 phải có hao phí nhân công!"

    print("\n✅ Toàn bộ Assertions logic phân tích ngữ nghĩa: ĐẠT 100%!")

    # Test Xuất Excel
    print("\n📄 Kiểm tra xuất file Excel Báo cáo Thẩm tra Dữ liệu Thô...")
    exporter = AuditExcelExporter()
    excel_bytes = exporter.export_workbook(result, project_title="THỬ NGHIỆM THẨM ĐỊNH VĨNH TÂN 4")

    out_file = Path(__file__).resolve().parent.parent / "tests" / "test_audit_output.xlsx"
    with open(out_file, "wb") as f:
        f.write(excel_bytes)

    # Đọc lại file Excel để kiểm tra tính toàn vẹn
    wb_read = openpyxl.load_workbook(out_file)
    assert "Bang_ThamTra_VOSC" in wb_read.sheetnames, "Thiếu sheet Bang_ThamTra_VOSC!"
    assert "Cum_Mau_LapLai" in wb_read.sheetnames, "Thiếu sheet Cum_Mau_LapLai!"
    assert "TongHop_TrangThai_HoSo" in wb_read.sheetnames, "Thiếu sheet TongHop_TrangThai_HoSo!"

    ws_vosc = wb_read["Bang_ThamTra_VOSC"]
    assert ws_vosc.auto_filter.ref is not None, "Sheet Bang_ThamTra_VOSC chưa bật AutoFilter!"
    print(f"✅ File Excel sinh thành công tại: {out_file} (Kích thước: {len(excel_bytes)} bytes)")
    print(f"✅ AutoFilter range: {ws_vosc.auto_filter.ref}")
    print(f"✅ Số sheets: {wb_read.sheetnames}")
    print("\n🎉 KIỂM THỬ TOÀN DIỆN MODULE ĐỘC LẬP HOÀN TẤT THÀNH CÔNG RỰC RỠ!")

if __name__ == "__main__":
    test_semantic_pipeline()
