# 📋 BIÊN BẢN CHUYỂN GIAO PHIÊN LÀM VIỆC (SESSION HANDOVER DOCUMENT)
**Dự án**: Hệ Thống Quản Lý Định Mức & Lập Dự Toán PCCC - SCL EVN (2026)  
**Đơn vị thực hiện**: Tổ Thẩm định EVN  
**Kho mã nguồn Git**: `https://github.com/nguyenhieuanhspkt/LapDuToanApp.git` (Nhánh: `main`)  
**Commit mới nhất**: `a37a30e` (`chore: save imported project items from PAKT quick intake test`)  
**Thời điểm lập**: 22/10/2026

---

## 1. TỔNG QUAN DỰ ÁN & MỤC TIÊU CỐT LÕI

Hệ thống phần mềm nội bộ chuyên dụng cho công tác lập và thẩm định dự toán PCCC - Sửa chữa lớn (SCL) cho các nhà máy điện trực thuộc EVN (điển hình: NMNĐ Vĩnh Tân 4).

### Mục tiêu kỹ thuật cốt lõi:
1. **Tính toán 100% bằng công thức động khép kín**: Mọi số liệu trong file Excel xuất ra (20-21 sheet) đều liên kết công thức từ sheet chi tiết lên sheet tổng hợp, không hardcode giá trị thành tiền.
2. **Minh chứng định mức scan gốc**: Nhúng trực tiếp ảnh chụp scan định mức gốc (Thông tư 38/2026/TT-BXD, QĐ 45/QĐ-HĐTV) vào từng sheet định mức thành phần.
3. **Chuẩn hóa giao diện F1 / G8**: Bảng tính 2 tầng phẳng (Master - Detail Grid), không sử dụng các expander sổ cụp gây rối mắt; thao tác nhập liệu, tra cứu, gán biến thể thông qua Modal Dialog (`@st.dialog`).
4. **Xử lý linh hoạt hao phí thực tế & biến thể**:
   - Quản lý tập trung biến thể $A0$ (100% gốc), $A1$ (bảo dưỡng không thay đế/vật tư), $A2$...
   - Tự động loại trừ đơn giá thiết bị/vật tư thay thế (đưa về 0 đ) khi chỉ thực hiện bảo dưỡng.
   - Tách nhân công chuyên sâu (ví dụ: `BD.41111` tách 0,174 công thành 0,087 công NC nhóm 3 + 0,087 công Kỹ sư bậc 5,0/8).
   - Tự động quy đổi ĐVT giữa PAKT và Định mức (Bộ/Cái $\leftrightarrow$ 10 đầu $\rightarrow$ hệ số $0,1$).

---

## 2. KIẾN TRÚC MÃ NGUỒN & CẤU TRÚC THƯ MỤC

```text
D:\TaskApp_kiet\LapDuToanApp\
├── app.py                         # Ứng dụng Streamlit chính (5 Tabs, chuẩn F1 2 tầng, @st.dialog)
├── config.py                      # Cấu hình đường dẫn OneDrive, CSDL, thư mục xuất Excel
├── gia_ca_may_lam_dong_2026.xlsx   # Bảng giá ca máy tỉnh Lâm Đồng năm 2026
├── requirements.txt               # Thư viện: streamlit, openpyxl, pandas, pillow...
├── run_app.bat                    # Script khởi chạy nhanh ứng dụng Streamlit
│
├── core/
│   ├── norm_manager.py            # Quản lý CSDL định mức, biến thể A0/A1, tính tỷ lệ bước k, tra giá
│   └── excel_generator.py         # Bộ sinh workbook Excel 20-21 sheet, công thức động, nhúng ảnh scan
│
├── database/
│   ├── norms_pccc_scl.json        # CSDL 13 định mức chuẩn TT38 & QĐ45 (quy trình bước, hao phí VL/NC/Máy)
│   ├── price_database_lam_dong_2026.json # Đơn giá vật tư, nhân công (vùng 2, 3, 4), ca máy & pháp lý
│   ├── project_variants.json      # Thư viện lưu trữ các biến thể định mức của dự án
│   └── current_project_items.json # Bảng danh mục công việc hiện tại của dự án (đang lưu 8 công tác)
│
├── utils/
│   └── num2words_vi.py            # Hàm chuyển đổi số tiền thành chữ tiếng Việt chuẩn ngân hàng/EVN
│
└── tests/
    └── test_trial_export.py       # Bộ kiểm thử tự động toàn diện kiểm tra 21 sheet, công thức và ảnh scan
```

---

## 3. TIẾN ĐỘ & CÁC TÍNH NĂNG ĐÃ HOÀN THÀNH

### ✅ 3.1. Tái cấu trúc toàn diện Tab 1 (Chuẩn F1 / Excel 2 tầng - 0 Expander)
- **Thanh công cụ Toolbar phẳng**:
  - `➕ THÊM CÔNG TÁC MỚI`: Pop-up modal `@st.dialog` hỗ trợ cả định mức CSDL (tự phát hiện ĐVT quy đổi) và báo giá trọn gói.
  - `📥 NHẬP TỪ PAKT / EXCEL`: Pop-up modal `@st.dialog` dán dữ liệu đa cột từ bảng thiết kế/PAKT.
  - `⚙️ QUẢN LÝ BIẾN THỂ (A0/A1...)`: Quản lý cấu hình biến thể và gán biến thể hàng loạt theo từ khóa.
  - `🔄 LÀM MỚI BẢNG`: Làm mới dữ liệu theo thời gian thực.
- **Tầng trên (Master Grid View)**:
  - Xem đơn giá hàng ngang: ĐG Vật liệu, ĐG Nhân công, ĐG Ca máy, ĐG Trực tiếp T, ĐG Trước thuế (sau khi nhân hệ số Chi phí chung $C=7,3\%$, $TT=2,5\%$, Thu nhập chịu thuế tính trước $TL=5,5\%$), Thành tiền trước thuế.
  - Chế độ Chỉnh sửa trực tiếp (F1 Editor): Cho phép sửa trực tiếp STT, Nhóm, Mã ĐM, Tên PAKT, ĐVT, Khối lượng, Đơn giá báo giá.
  - Công cụ Đổi tên nhóm và Xóa nhóm trực quan.
- **Tầng dưới (F1 Detail Inspector)**:
  - Tự động đồng bộ soi chi tiết công việc đang chọn mà không cần mở expander.
  - **Cột trái**: Bảng các bước quy trình & tỷ lệ hao phí $k$, cho phép điều chỉnh bật/tắt bước, đổi biến thể `A0/A1`, lưu cấu hình bước hoặc khôi phục gốc 100%.
  - **Cột phải**: Bảng phân rã chi tiết tài nguyên (Vật liệu, Nhân công, Máy thi công) chuẩn F1/G8, định mức cơ sở vs định mức sau $k$, đơn giá và thành tiền tài nguyên.

### ✅ 3.2. Bộ bóc tách dán dữ liệu thông minh (Smart Multi-Column Parser)
- Đã giải quyết triệt để lỗi thiếu thư viện và hỗ trợ nhận diện các bảng Excel thực tế có thứ tự:
  $$\text{STT} \rightarrow \text{Tên công tác} \rightarrow \text{Mô tả / Quy cách} \rightarrow \text{Khối lượng} \rightarrow \text{Đơn vị tính}$$
- Đã kiểm thử dán thành công thực tế 8 công tác từ PAKT vào `database/current_project_items.json`.

### ✅ 3.3. File Xuất Excel Thẩm Định 20-21 Sheet
- Sheet **`TH_KinhPhi`** đứng ở vị trí Sheet 1 đầu tiên, công thức động liên kết từ sheet khối lượng, chi tiết đơn giá, chi phí báo giá/tạm tính, và có dòng "Bằng chữ" tự sinh bằng `num2words_vi`.
- Sheet **`Khối lượng-Định mức-ĐVT`** phân nhóm La Mã `NHÓM I: ...` chuẩn mực.
- Sheet **`01_ChiTiet_HaoPhi_DinhMuc`** liên kết động đơn giá từ các sheet TH Vật liệu, TH Nhân công, TH Máy.
- Các sheet **`TH_VatLieu`**, **`TH_NhanCong`**, **`TH_MayThiCong`** định dạng theo đúng mẫu bảng giá hiện tại kèm cột Căn cứ tính giá.

---

## 4. TRẠNG THÁI KIỂM THỬ & ĐỒNG BỘ GIT

| Hạng mục kiểm tra | Lệnh thực hiện | Kết quả | Ghi chú |
| :--- | :--- | :---: | :--- |
| **Cú pháp Python** | `python -m py_compile app.py` | **PASS** | Biên dịch sạch 100%, 0 lỗi |
| **Kiểm thử xuất Excel** | `python tests/test_trial_export.py` | **PASS** | Đạt 100% 21 sheet, công thức, ảnh scan |
| **Kiểm tra Expander Tab 1** | Script kiểm tra số lượng | **0** | Xóa bỏ hoàn toàn 100% expander ở Tab 1 |
| **Trạng thái Git** | `git status` | **CLEAN** | Nhánh `main`, up-to-date với remote |
| **Remote Github** | `git push origin main` | **SYNCED** | Đã đẩy tới commit `a37a30e` |

---

## 5. KẾ HOẠCH & CÔNG VIỆC CẦN THỰC HIỆN TRONG PHIÊN TIẾP THEO

Khi bắt đầu phiên làm việc mới, Agent hoặc Chuyên viên có thể tiếp tục ngay các nội dung sau:

1. **Hoàn thiện danh mục khối lượng dự án thực tế**:
   - Sử dụng nút `📥 NHẬP TỪ PAKT / EXCEL` để dán nốt các công tác còn lại của hồ sơ PAKT Vĩnh Tân 4.
   - Kiểm tra mã định mức và biến thể ($A0$ cho thay mới 100%, $A1$ cho bảo dưỡng đo kiểm).
2. **Rà soát đơn giá và vùng áp dụng tại Tab 2**:
   - Kiểm tra đơn giá nhân công và ca máy tại địa bàn thực hiện (Lâm Đồng Vùng IV hoặc Vùng II/III, hoặc Khánh Hòa).
   - Bổ sung văn bản pháp lý/báo giá cho các mã công tác tạm tính `BG.xx` nếu có.
3. **Kiểm tra Tab 3 & Tab 4**:
   - Kiểm tra Bảng tổng hợp kinh phí dự toán, hệ số chi phí chung ($C$), chi phí gián tiếp ($TT$), thu nhập chịu thuế tính trước ($TL$), thuế GTGT ($VAT$).
   - Kiểm tra việc hiển thị ảnh scan minh chứng định mức TT38 & QĐ45 ở Tab 4.
4. **Xuất file Excel thẩm định chính thức**:
   - Nhấn nút xuất file tại Tab 5, kiểm tra mở file bằng Microsoft Excel, đối chiếu khớp số liệu 100% với phương án kinh tế kỹ thuật được duyệt.

---

## 6. LỆNH THAO TÁC NHANH (QUICK COMMANDS)

```bash
# 1. Di chuyển vào thư mục dự án
cd "D:\TaskApp_kiet\LapDuToanApp"

# 2. Khởi chạy ứng dụng Streamlit
streamlit run app.py --server.port 8501

# 3. Chạy kiểm thử tự động xuất Excel
python tests/test_trial_export.py

# 4. Kiểm tra mã nguồn trước khi commit
git status
```
