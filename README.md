# LapDuToanApp - Ứng dụng Tự động hóa Lập Dự toán & Minh chứng Định mức

Ứng dụng Web App xây dựng bằng **Streamlit (Python)** phục vụ công tác lập và thẩm định dự toán PCCC / SCL EVN.

## Tính năng chính:
- Tự động bóc tách và tra cứu định mức chuẩn (QĐ 45/QĐ-HĐTV và TT 38/2026/TT-BXD).
- Minh chứng số hóa trực quan: Tự động trích xuất và chèn trang PDF scan gốc vào từng mã định mức.
- Tự động sinh file Excel dự toán với **100% công thức liên kết động**:
  - Cột Định mức link từ sheet minh chứng định mức.
  - Cột Đơn giá link từ 3 bảng tổng hợp Vật liệu, Nhân công, Máy thi công.
  - Thành tiền và tổng hợp chi phí tự động cập nhật.
- Dễ dàng quản lý, mở rộng và bổ sung định mức mới qua giao diện web.

## Cấu trúc lưu trữ:
- **Mã nguồn:** `D:\TaskApp_kiet\LapDuToanApp`
- **Dữ liệu & Kho ảnh scan:** `D:\onedrive_hieuna\OneDrive - EVN\Tổ Thẩm định\Hiếu\2026_LAPDUTOAN_APP`
- **GitHub Repository:** `LapDuToanApp`

## Hướng dẫn sử dụng:
1. Cài đặt môi trường: `pip install -r requirements.txt`
2. Khởi chạy ứng dụng: Nhấp đúp vào file `run_app.bat` hoặc gõ lệnh:
   ```bash
   streamlit run app.py
   ```
