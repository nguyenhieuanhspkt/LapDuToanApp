# -*- coding: utf-8 -*-
"""
Script xuất bản Bộ Hồ Sơ Dự Toán PCCC 2026 đầy đủ 876 công tác (19 Sheet, 100% công thức động).
Cách chạy:
    python export_dutoan.py
Author: Tổ Thẩm Định EVN
"""

import sys
import time
from pathlib import Path

# Đảm bảo in tiếng Việt trên console Windows
sys.stdout.reconfigure(encoding='utf-8')

# Import bộ quản lý định mức và bộ điều phối xuất Excel
from core.norm_manager import NormManager
from core.excel_generator import ExcelGenerator

def main():
    print("=" * 80)
    print("🚀 ĐANG KHỞI TẠO XUẤT HỒ SƠ DỰ TOÁN PCCC 2026 (876 CÔNG TÁC)...")
    print("=" * 80)

    # 1. Nạp cơ sở dữ liệu định mức và 876 công tác dự án
    mgr = NormManager()
    items = mgr.load_work_items()
    print(f"✅ Đã nạp thành công {len(items)} công tác từ cơ sở dữ liệu dự án.")

    # 2. Khởi tạo bộ xuất Excel (kiến trúc Modular theo từng sheet)
    gen = ExcelGenerator(mgr)

    # 3. Đặt tên file xuất ra (tự động xử lý fallback nếu file đang mở trong Excel)
    base_name = "DU_TOAN_PCCC_2026_AP_MA_CHUAN"
    output_file = Path(f"{base_name}.xlsx")
    try:
        # Kiểm tra nhanh file có bị khóa không
        if output_file.exists():
            with open(output_file, "a"):
                pass
    except (PermissionError, IOError):
        import datetime
        timestamp = datetime.datetime.now().strftime("%H%M%S")
        output_file = Path(f"{base_name}_{timestamp}.xlsx")
        print(f"⚠️ File '{base_name}.xlsx' đang được mở trong Excel, tự động lưu sang: {output_file.name}")

    print(f"⏳ Đang xử lý tính toán, thiết lập 100% công thức động và tạo 19 sheet...")

    start_time = time.time()
    gen.generate(
        project_name="BẢO DƯỠNG PCCC NĂM 2026 - NMNĐ VĨNH TÂN 4",
        category_name="Hệ thống báo cháy tự động & Chữa cháy",
        work_items=items,
        output_path=output_file
    )
    elapsed = time.time() - start_time

    # 4. Thông báo kết quả
    file_size_kb = output_file.stat().st_size / 1024
    print("\n" + "=" * 80)
    print(f"🎉 XUẤT BẢN THÀNH CÔNG FILE DỰ TOÁN!")
    print(f"📁 Đường dẫn file : {output_file.resolve()}")
    print(f"📦 Dung lượng     : {file_size_kb:,.1f} KB")
    print(f"⏱️ Thời gian chạy : {elapsed:.2f} giây")
    print("=" * 80)

if __name__ == "__main__":
    main()
