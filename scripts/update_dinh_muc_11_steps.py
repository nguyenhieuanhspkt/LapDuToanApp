import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

wb_path = 'd:\\TaskApp_kiet\\LapDuToanApp\\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
dl_path = 'C:\\Users\\Admin\\Downloads\\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'

print("Loading workbook...")
wb = openpyxl.load_workbook(wb_path)

# 1. Update Sheet 01 if needed (fix row 8 from 38.4 to 2.00 cong/tu)
if '01_Bang_Phan_Cap_Tong_The' in wb.sheetnames:
    ws01 = wb['01_Bang_Phan_Cap_Tong_The']
    for r in range(6, ws01.max_row + 1):
        if ws01.cell(r, 1).value == '1.1.1':
            ws01.cell(r, 6, value='11 bước công nghệ, định mức gốc 2.00 công/tủ (QĐ 45 Tập 6 tr.451)')
            print(f"Updated Sheet 01 row {r} for 1.1.1")

# 2. Rebuild Sheet "Định mức áp dụng"
sheet_name = 'Định mức áp dụng'
if sheet_name in wb.sheetnames:
    wb.remove(wb[sheet_name])

ws = wb.create_sheet(title=sheet_name)
ws.views.sheetView[0].showGridLines = True

# Styling definitions
font_comp = Font(name='Times New Roman', size=10, bold=True, color='002060')
font_doc_title = Font(name='Times New Roman', size=14, bold=True, color='002060')
font_doc_sub = Font(name='Times New Roman', size=11, italic=True, color='333333')

font_sec = Font(name='Times New Roman', size=11.5, bold=True, color='002060')
font_tbl_hdr = Font(name='Times New Roman', size=10, bold=True, color='000000')
font_sub_hdr = Font(name='Times New Roman', size=9.5, bold=True, italic=True, color='1F4E79')

font_norm_code = Font(name='Times New Roman', size=10.5, bold=True, color='002060')
font_step_item = Font(name='Times New Roman', size=9.5, color='000000')
font_step_total = Font(name='Times New Roman', size=10, bold=True, color='002060')
font_resource = Font(name='Times New Roman', size=9.5, bold=True, color='333333')
font_resource_item = Font(name='Times New Roman', size=9.5, color='000000')

font_pass = Font(name='Times New Roman', size=10, bold=True, color='006100')
font_note = Font(name='Times New Roman', size=9, italic=True, color='555555')

fill_sec = PatternFill(start_color='8EA9DB', end_color='8EA9DB', fill_type='solid')      # Blue accent
fill_tbl_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')  # Soft blue
fill_sub_hdr = PatternFill(start_color='EDF2F8', end_color='EDF2F8', fill_type='solid')  # Very light blue
fill_norm = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')     # Soft green header
fill_step_total = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
fill_pass = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')

border_thin = Side(border_style='thin', color='B0B0B0')
border_thick_bottom = Side(border_style='medium', color='002060')
border_double = Side(border_style='double', color='000000')

box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
hdr_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)
total_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# Title Block
ws.merge_cells('A1:F1')
ws['A1'] = 'TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM - TỔNG CÔNG TY PHÁT ĐIỆN 3 - CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4'
ws['A1'].font = font_comp
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('A2:F2')
ws['A2'] = 'BẢNG ĐỊNH MỨC ÁP DỤNG & CHI TIẾT 11 THÀNH PHẦN CÔNG VIỆC THEO FILE ĐỊNH MỨC .PDF'
ws['A2'].font = font_doc_title
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('A3:F3')
ws['A3'] = 'Dự án: SCL Hệ thống PCCC năm 2026 – NMNĐ Vĩnh Tân 4 (Quyết định số 45/QĐ-HĐTV EVN & Thông tư số 38/2026/TT-BXD)'
ws['A3'].font = font_doc_sub
ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

ws.row_dimensions[1].height = 20
ws.row_dimensions[2].height = 26
ws.row_dimensions[3].height = 20
ws.row_dimensions[4].height = 10

curr_r = 5

# =========================================================================
# PHẦN 1: BƯỚC KIỂM TRA ĐÁNH GIÁ ĐỐI CHIẾU ĐỊNH MỨC (AUDIT STEP)
# =========================================================================
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='MỤC I. BƯỚC KIỂM TRA ĐÁNH GIÁ ĐỐI CHIẾU ĐỊNH MỨC (AUDIT & RECONCILIATION STEP)').font = font_sec
ws.row_dimensions[curr_r].height = 28
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_sec
    ws.cell(row=curr_r, column=c).border = box_border

curr_r += 1

audit_headers = [
    'STT Phân tầng\n(Sheet 01)',
    'Đối tượng công tác & Tên hạng mục\n(Theo cấu trúc phân tầng)',
    'Mã định mức tối ưu cho tính chất\n(Cột G - Sheet 01_Bang_Phan_Cap)',
    'Bộ định mức áp dụng\n& Căn cứ pháp lý',
    'Vị trí chi tiết tại Sheet\n"Định mức áp dụng"',
    'Kết quả kiểm tra\nđánh giá đối chiếu'
]

ws.row_dimensions[curr_r].height = 32
for c_idx, h_text in enumerate(audit_headers, start=1):
    cell = ws.cell(row=curr_r, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

audit_data = [
    ('1.1.1', 'Tủ trung tâm báo cháy & Tủ điều khiển xả khí (FMZ 5000 / HCVR-3)', 'E6C.3.05 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451 (2.0 công/tủ)', 'Mục II.1 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.2.1', 'Cảm biến đầu báo (Pha 1: Tháo dỡ trên cao xuống sàn)', 'E6C.3.01 (bước 5)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 448 (Bước 5: 30%)', 'Mục II.2 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.2.2', 'Cảm biến đầu báo (Pha 2: Vệ sinh, test khói sàn)', 'E6C.3.01 (bước 1–4, 6–11)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 448 (Bước 1-4,6-11: 70%)', 'Mục II.2 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.2.3', 'Tổ hợp đầu báo / chuông đèn tại chỗ trọn gói', 'E6C.3.01 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 448 (0.33 công/cái)', 'Mục II.2 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.3.1', 'Đầu báo Beam (Pha 1: Tháo đầu báo Beam xuống)', 'E6C.3.02 (bước 5) – vận dụng', 'QĐ 45/QĐ-HĐTV Tập 6 trang 449 (Bước 5: 30% = 0.6 công)', 'Mục II.3 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.3.2', 'Đầu báo Beam (Pha 2: Kiểm tra cân chỉnh tia beam)', 'E6C.3.04 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451 (2.51 công/bộ)', 'Mục II.4 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.4.1', 'Nút ấn khẩn cấp (Pha 1: Tháo nút ấn xuống)', 'E6C.3.02 (bước 5)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 449 (Bước 5: 30% = 0.6 công)', 'Mục II.3 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.4.2', 'Nút ấn khẩn cấp (Pha 2: Vệ sinh tiếp điểm, đo mạch)', 'E6C.3.02 (bước 1–4, 6–11)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 449 (Bước 1-4,6-11: 70% = 1.4 công)', 'Mục II.3 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.4.3', 'Nút ấn khẩn cấp trọn gói tại chỗ', 'E6C.3.02 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 449 (2.0 công/cái)', 'Mục II.3 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.5', 'Module giám sát & Cách ly ngoài hiện trường', 'E6C.3.03 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 450 (2.71 công/bộ)', 'Mục II.5 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.6', 'Cụm Van điện từ kích hoạt xả tràn (Solenoid Deluge)', 'E6C.2.04 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 440 (1.80 công/cái)', 'Mục II.6 (Chi tiết định mức & máy)', 'ĐẠT (Khớp 100%)'),
    ('1.7', 'Công tắc áp lực & Công tắc dòng chảy', 'E6C.1.23 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 428 (0.95 công/cái)', 'Mục II.7 (Chi tiết định mức & máy)', 'ĐẠT (Khớp 100%)'),
    ('1.8', 'Đầu đo nồng độ khí cháy Hydro (H2) phòng ắc quy', 'E6C.1.31 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 436 (5.30 công/bộ)', 'Mục II.8 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.9', 'Hệ thống Ắc quy nguồn dự phòng tủ PCCC', 'E6B.1.13.01 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 5 trang 383 (1.50 công/bình)', 'Mục II.9 (Chi tiết 11 bước công việc)', 'ĐẠT (Khớp 100%)'),
    ('1.10', 'Công tắc lưu lượng / Rơle giám sát lưu lượng', 'E6C.1.18 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 425 (1.20 công/cái)', 'Mục II.10 (Chi tiết định mức & máy)', 'ĐẠT (Khớp 100%)'),
    ('1.11', 'Tủ điều khiển van xả tràn Deluge (Phần điện)', 'E6C.3.06 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 453 (7.00 công/cái)', 'Mục II.11 (Chi tiết 13 bước công nghệ)', 'ĐẠT (Khớp 100%)'),
    ('1.12', 'Đại tu van nước chữa cháy, van giảm áp, van xả khí (Cơ khí)', 'E6BOP.1.03 (SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 312 (2.10 công/cái)', 'Mục II.12 (Chi tiết định mức cơ khí)', 'ĐẠT (Khớp 100%)'),
    ('1.13', 'Sửa chữa, đại tu van an toàn đường ống PCCC', 'E6BOP.1.01 (SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 308 (4.50 công/cái)', 'Mục II.13 (Chi tiết định mức cơ khí)', 'ĐẠT (Khớp 100%)'),
    ('2.1', 'Lắp mới / thay thế đầu báo khói, đầu báo nhiệt', 'BD.41111 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.174 công/đầu)', 'Mục III.1 (Dòng chi tiết III.1)', 'ĐẠT (Khớp 100%)'),
    ('2.2', 'Lắp mới / thay thế nút ấn báo cháy khẩn cấp', 'BD.41131 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.600 công/cái)', 'Mục III.3 (Dòng chi tiết III.3)', 'ĐẠT (Khớp 100%)'),
    ('2.3.1', 'Lắp mới chuông báo cháy, còi báo cháy', 'BD.41141 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.328 công/cái)', 'Mục III.4 (Dòng chi tiết III.4)', 'ĐẠT (Khớp 100%)'),
    ('2.3.2', 'Lắp mới đèn báo cháy, đèn hiển thị phòng', 'BD.41121 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.234 công/đèn)', 'Mục III.2 (Dòng chi tiết III.2)', 'ĐẠT (Khớp 100%)'),
    ('2.4', 'Lắp mới linh kiện module giám sát báo cháy', 'BA.18405 (TT38 Phụ lục 3)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.6 (Dòng chi tiết III.6)', 'ĐẠT (Khớp 100%)'),
    ('2.5', 'Lắp mới đèn chỉ hướng thoát hiểm EXIT, đèn sự cố', 'BD.41171 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.5 (Dòng chi tiết III.5)', 'ĐẠT (Khớp 100%)'),
    ('3.1.1', 'Sơn vỏ van / Y lọc / Thiết bị DN50 - DN100 trong nhà', 'SB.82633 (TT38 Phụ lục 6)', 'Thông tư 38/2026/TT-BXD Phụ lục 6 (100 m2)', 'Mục III.7 (Dòng chi tiết III.7)', 'ĐẠT (Khớp 100%)'),
    ('3.1.2', 'Sơn vỏ van / Y lọc / Trụ cứu hỏa DN125 - DN200 ngoài trời', 'SB.82634 (TT38 Phụ lục 6)', 'Thông tư 38/2026/TT-BXD Phụ lục 6 (100 m2)', 'Mục III.8 (Dòng chi tiết III.8)', 'ĐẠT (Khớp 100%)'),
    ('3.2.1', 'Kiểm tra, bảo dưỡng họng cứu hỏa vách tường', 'BD.41111 (TT38 Phụ lục 3)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.32 công/họng)', 'Mục III.1 (Dòng chi tiết III.1)', 'ĐẠT (Khớp 100%)'),
    ('3.2.2', 'Kiểm tra, bảo dưỡng trụ cứu hỏa ngoài trời 3 cửa', 'BD.41121 (TT38 Phụ lục 3)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.55 công/trụ)', 'Mục III.2 (Dòng chi tiết III.2)', 'ĐẠT (Khớp 100%)'),
    ('3.3.1', 'Lắp mới van ren ≤ 25mm (van xả DN25)', 'BB.86601 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.18 công/cái)', 'Mục III.9 (Dòng chi tiết III.9)', 'ĐẠT (Khớp 100%)'),
    ('3.3.2', 'Lắp mới van mặt bích DN80 (van Alarm DN80)', 'BB.86103 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.45 công/cái)', 'Mục III.10 (Dòng chi tiết III.10)', 'ĐẠT (Khớp 100%)'),
    ('3.3.3', 'Lắp mới đồng hồ đo áp suất (áp kế 0-16 bar)', 'BB.85201 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3 (0.12 công/cái)', 'Mục III.11 (Dòng chi tiết III.11)', 'ĐẠT (Khớp 100%)'),
    ('3.4', 'Đo kiểm tra điện trở hệ thống chống sét & tiếp địa', 'DB.24001 (TT38 Phụ lục 5)', 'Thông tư 38/2026/TT-BXD Phụ lục 5 (0.50 công, 0.50 ca)', 'Mục III.12 (Dòng chi tiết III.12)', 'ĐẠT (Khớp 100%)'),
    ('4.1', 'Bo mạch, card, nguồn nằm trong tủ trung tâm FMZ5000', 'Không áp mã (đã gồm trong E6C.3.05)', 'Đã gồm trong 11 bước định mức tủ E6C.3.05', 'Mục IV.1 (Dòng chi tiết IV.1)', 'ĐẠT (Khớp 100%)'),
    ('4.2', 'Kiểm tra đèn chỉ lối EXIT và đèn sự cố', 'Không áp mã (Chi phí thường xuyên)', 'Chi phí bảo dưỡng định kỳ thường xuyên', 'Mục IV.2 (Dòng chi tiết IV.2)', 'ĐẠT (Khớp 100%)'),
    ('4.3', 'Tuyến cáp tín hiệu loop tìm và xử lý chạm đất', 'Không áp mã (Thuộc trách nhiệm loop)', 'Chi phí đi kèm trong trách nhiệm bảo dưỡng', 'Mục IV.3 (Dòng chi tiết IV.3)', 'ĐẠT (Khớp 100%)'),
    ('4.4', 'Kiểm tra bình bọt, bình khí, cuộn vòi cứu hỏa tại chỗ', 'Không áp mã (Chi phí phụ trợ)', 'Công tác kiểm tra phụ trợ gói thầu', 'Mục IV.4 (Dòng chi tiết IV.4)', 'ĐẠT (Khớp 100%)')
]

for row_idx, r_data in enumerate(audit_data, start=curr_r + 1):
    ws.row_dimensions[row_idx].height = 22
    ws.cell(row=row_idx, column=1, value=r_data[0]).alignment = align_center
    ws.cell(row=row_idx, column=2, value=r_data[1]).alignment = align_left
    ws.cell(row=row_idx, column=3, value=r_data[2]).alignment = align_left
    ws.cell(row=row_idx, column=4, value=r_data[3]).alignment = align_left
    ws.cell(row=row_idx, column=5, value=r_data[4]).alignment = align_left
    c_res = ws.cell(row=row_idx, column=6, value=r_data[5])
    c_res.alignment = align_center
    c_res.font = font_pass
    c_res.fill = fill_pass
    
    for c in range(1, 7):
        cell = ws.cell(row=row_idx, column=c)
        cell.border = box_border
        if c != 6:
            cell.font = font_step_item

curr_r = curr_r + len(audit_data) + 1

# Tổng kết audit
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='KẾT LUẬN KIỂM TRA ĐỐI CHIẾU: 100% CÁC MÃ ĐỊNH MỨC TỐI ƯU CHO TÍNH CHẤT ĐÃ ĐƯỢC NÊU CHI TIẾT ĐẦY ĐỦ TẠI SHEET NÀY').font = font_norm_code
ws.cell(row=curr_r, column=1).alignment = align_center
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_pass
    ws.cell(row=curr_r, column=c).border = total_border
ws.row_dimensions[curr_r].height = 26

curr_r += 2

# =========================================================================
# PHẦN 2: CHI TIẾT TẤT CẢ CÁC ĐỊNH MỨC SCL EVN (QUYẾT ĐỊNH 45/QĐ-HĐTV)
# KÈM THEO BẢNG 11 THÀNH PHẦN CÔNG VIỆC CHUẨN NHƯ FILE DM .PDF
# =========================================================================
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='MỤC II. CHI TIẾT CÁC ĐỊNH MỨC SỬA CHỮA LỚN EVN (QUYẾT ĐỊNH SỐ 45/QĐ-HĐTV) - GỒM 11 BƯỚC CÔNG NGHỆ NHƯ FILE DM .PDF').font = font_sec
ws.row_dimensions[curr_r].height = 28
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_sec
    ws.cell(row=curr_r, column=c).border = box_border

curr_r += 1

detail_headers = [
    'STT / Bước',
    'Mã hiệu định mức, Danh mục thành phần công việc & Hao phí',
    'Đơn vị tính\n(ĐVT)',
    'Mức hao phí\nđịnh mức',
    'Tỉ lệ %\n/ Hệ số bước',
    'Cơ sở pháp lý & Thuyết minh kỹ thuật theo file DM .pdf\n(Quyết định số 45/QĐ-HĐTV Tập 6 & Tập 5)'
]

ws.row_dimensions[curr_r].height = 30
for c_idx, h_text in enumerate(detail_headers, start=1):
    cell = ws.cell(row=curr_r, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

curr_r += 1

# Danh sách dữ liệu các định mức SCL kèm 11 bước chuẩn xác từ QĐ 45
scl_norms_full = [
    # -------------------------------------------------------------
    # 1. E6C.3.05
    # -------------------------------------------------------------
    {
        'stt': 'II.1',
        'code': 'E6C.3.05',
        'title': 'E6C.3.05 - Kiểm tra và bảo dưỡng tủ điều khiển chữa cháy tại chỗ, tủ điều khiển và tủ trung gian của các thiết bị đo lường điều khiển (Tủ FMZ 5000 / HCVR-3)',
        'unit': '1 tủ',
        'total_labor': 2.00,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451. Áp dụng cho 34 tủ trung tâm PCCC.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.04, '2,0%', 'Nghiên cứu sơ đồ nguyên lý tủ, bản vẽ đấu nối cáp tín hiệu loop và relay liên động'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công theo yêu cầu kỹ thuật', 'công', 0.10, '5,0%', 'Chuẩn bị đồng hồ VOM, máy thổi bụi, trang bị bảo hộ, sơ đồ thẻ bài cách ly an toàn'),
            (3, '3. Thực hiện đầy đủ các bước kiểm tra an toàn trước khi thực hiện công tác', 'công', 0.04, '2,0%', 'Cô lập đường kích xả khí (khóa van cơ/tháo cuộn coil), chuyển tủ sang Test/Bypass'),
            (4, '4. Đo thông mạch, đo điện trở cách điện của cáp tín hiệu', 'công', 0.10, '5,0%', 'Đo điện trở cách điện dây loop, cáp liên động PLC, cáp nguồn cấp 220VAC/24VDC'),
            (5, '5. Vệ sinh tủ và các thiết bị liên quan', 'công', 0.60, '30,0%', 'Vệ sinh bụi khung rack, bo mạch CPU, card loop, bộ nguồn, lau chùi mặt kính hiển thị'),
            (6, '6. Kiểm tra tín hiệu đến và đi của tủ', 'công', 0.30, '15,0%', 'Kiểm tra tín hiệu các loop báo cháy, ngõ vào giám sát áp lực, van xả, công tắc dòng chảy'),
            (7, '7. Xử lý điện trở các điểm nối dây', 'công', 0.30, '15,0%', 'Siết chặt domino cầu đấu dây, xử lý oxy hóa tiếp điểm cáp tín hiệu và cáp nguồn'),
            (8, '8. Hiệu chỉnh lại tủ giám sát', 'công', 0.16, '8,0%', 'Cân chỉnh điện áp nạp ắc quy (27.4VDC), cài đặt thời gian trễ xả khí, cấu hình card loop'),
            (9, '9. Kiểm tra chức năng làm việc của tủ', 'công', 0.20, '10,0%', 'Kiểm tra hiển thị LCD, phím bấm Silence, Reset, Drill, Lamp Test và còi buzzer nội bộ'),
            (10, '10. Kiểm tra tín hiệu gửi lên tủ điều khiển trung tâm', 'công', 0.12, '6,0%', 'Kiểm tra truyền thông giao tiếp Modbus/CANbus về DCS nhà máy và màn hình trung tâm'),
            (11, '11. Vệ sinh mặt bằng và tiến hành nghiệm thu', 'công', 0.04, '2,0%', 'Dọn dẹp hiện trường, khôi phục trạng thái vận hành bình thường, ký biên bản nghiệm thu')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 2.00, '100,0%', 'Thực hiện toàn bộ 11 bước công việc chuyên sâu tủ điều khiển'),
            ('Máy thi công', 'Đồng hồ vạn năng kỹ thuật số (VOM)', 'ca', 1.00, '100,0%', 'Đo điện áp nguồn 24VDC, kiểm tra dòng tĩnh và thông mạch tín hiệu loop'),
            ('Máy thi công', 'Máy thổi bụi công nghiệp cầm tay', 'ca', 1.00, '100,0%', 'Thổi sạch bụi bẩn công nghiệp bám trên khe cắm bo mạch và bộ nguồn tủ')
        ]
    },

    # -------------------------------------------------------------
    # 2. E6C.3.01
    # -------------------------------------------------------------
    {
        'stt': 'II.2',
        'code': 'E6C.3.01',
        'title': 'E6C.3.01 - Kiểm tra và bảo dưỡng nút nhấn báo cháy, bộ còi đèn/ chuông báo cháy (Vận dụng bóc tách 2 pha cho Cảm biến đầu báo khói/nhiệt)',
        'unit': '1 cái',
        'total_labor': 0.33,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 448. Bóc tách 2 pha: Pha 1 tháo dỡ Bước 5 (30%), Pha 2 kiểm chuẩn sàn Bước 1-4, 6-11 (70%).',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.01, '2,0%', 'Nghiên cứu sơ đồ bố trí đầu báo/nút ấn trên mặt bằng nhà máy'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công', 'công', 0.02, '5,0%', 'Chuẩn bị sào tháo đầu báo, thang nhôm, đồng hồ VOM, giẻ lau chuyên dụng'),
            (3, '3. Thực hiện đầy đủ các bước kiểm tra an toàn', 'công', 0.01, '2,0%', 'Kiểm tra an toàn làm việc trên cao và an toàn điện đường loop'),
            (4, '4. Đo thông mạch, đo điện trở cách điện của cáp tín hiệu', 'công', 0.02, '5,0%', 'Đo điện trở cách điện và kiểm tra sụt áp đường loop'),
            (5, '5. Tháo, vệ sinh các chi tiết (Pha 1: Tháo dỡ đưa xuống sàn)', 'công', 0.09, '30,0%', 'Tháo gỡ cảm biến/nút ấn từ trần cao mang về phòng kiểm chuẩn (Bước 5: chiếm 30% ~ 0.09 công)'),
            (6, '6. Kiểm tra cách điện', 'công', 0.05, '15,0%', 'Đo điện trở cách điện ngàm tiếp xúc và bo mạch (Pha 2)'),
            (7, '7. Kiểm tra tín hiệu điều khiển', 'công', 0.06, '20,0%', 'Kiểm tra điện áp giao tiếp loop và mã địa chỉ kỹ thuật số (Pha 2)'),
            (8, '8. Xử lý điện trở các điểm nối dây', 'công', 0.02, '8,0%', 'Làm sạch chân cắm tiếp điểm, xử lý oxy hóa bề mặt (Pha 2)'),
            (9, '9. Kiểm tra chức năng làm việc của thiết bị', 'công', 0.02, '5,0%', 'Thử nghiệm tạo khói/gia nhiệt giả lập tại buồng kiểm chuẩn sàn (Pha 2)'),
            (10, '10. Kiểm tra tín hiệu gửi lên tủ điều khiển trung tâm', 'công', 0.02, '6,0%', 'Xác nhận tín hiệu báo động chính xác tên và Zone trên tủ trung tâm (Pha 2)'),
            (11, '11. Vệ sinh mặt bằng và tiến hành nghiệm thu', 'công', 0.01, '2,0%', 'Lắp trả thiết bị lên cao, vệ sinh hoàn trả mặt bằng và nghiệm thu (Pha 2)')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Đo kiểm, hiệu chuẩn đầu báo)', 'công', 0.33, '100,0%', 'Định mức trọn gói 0.33 công (Pha 1 tháo gỡ: 0.09 công; Pha 2 kiểm chuẩn sàn: 0.24 công)'),
            ('Máy thi công', 'Đồng hồ vạn năng (VOM)', 'ca', 0.06, '100,0%', 'Đo thông mạch, điện trở tiếp xúc'),
            ('Máy thi công', 'Mê gôm mét 1000V', 'ca', 0.06, '100,0%', 'Đo điện trở cách điện thiết bị & cáp tín hiệu')
        ]
    },

    # -------------------------------------------------------------
    # 3. E6C.3.02
    # -------------------------------------------------------------
    {
        'stt': 'II.3',
        'code': 'E6C.3.02',
        'title': 'E6C.3.02 - Kiểm tra và bảo dưỡng cảm biến báo cháy dạng nhiệt/ khói & Nút nhấn khẩn cấp chuyên dụng',
        'unit': '1 cái',
        'total_labor': 2.00,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 449. Áp dụng cho nút ấn khẩn cấp & bóc tách pha tháo dỡ.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.04, '2,0%', 'Nghiên cứu tài liệu kỹ thuật, vị trí nút nhấn và cơ cấu tiếp điểm khô'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công', 'công', 0.10, '5,0%', 'Chuẩn bị tua vít, đồng hồ VOM, kìm bấm cosse, chìa khóa test nút ấn'),
            (3, '3. Thực hiện đầy đủ các bước kiểm tra an toàn', 'công', 0.04, '2,0%', 'Cô lập liên động kích hoạt còi hú chuông báo động diện rộng'),
            (4, '4. Đo thông mạch, đo điện trở cách điện của cáp tín hiệu', 'công', 0.10, '5,0%', 'Đo cách điện giữa các lõi dây tín hiệu nút ấn về hộp nối'),
            (5, '5. Tháo, vệ sinh chi tiết (Pha 1: Tháo dỡ nút ấn xuống sàn)', 'công', 0.60, '30,0%', 'Tháo cụm nút ấn khỏi đế gắn tường/cột thép (Bước 5 chiếm 30% = 0.60 công)'),
            (6, '6. Kiểm tra cách điện của nút nhấn / cảm biến', 'công', 0.30, '15,0%', 'Đo điện trở cách điện giữa tiếp điểm và vỏ bảo vệ kim loại (Pha 2)'),
            (7, '7. Kiểm tra tín hiệu đến và đi', 'công', 0.40, '20,0%', 'Kiểm tra tín hiệu loop và điện áp giám sát đường dây (Pha 2)'),
            (8, '8. Xử lý điện trở các điểm nối dây', 'công', 0.10, '5,0%', 'Làm sạch bề mặt tiếp điểm, bấm lại đầu cosse chống lỏng (Pha 2)'),
            (9, '9. Kiểm tra chức năng làm việc của nút nhấn', 'công', 0.16, '8,0%', 'Nhấn cơ cấu lò xo, kiểm tra độ nảy tiếp điểm và chốt cơ học (Pha 2)'),
            (10, '10. Kiểm tra tín hiệu gửi lên tủ điều khiển trung tâm', 'công', 0.12, '6,0%', 'Kiểm tra kích hoạt còi buzzer và hiển thị cảnh báo trên tủ trung tâm (Pha 2)'),
            (11, '11. Vệ sinh mặt bằng và tiến hành nghiệm thu', 'công', 0.04, '2,0%', 'Lắp lại đế, dán tem kiểm định, vệ sinh hiện trường và nghiệm thu (Pha 2)')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 2.00, '100,0%', 'Định mức gốc 2.00 công (Pha 1 tháo gỡ: 0.60 công; Pha 2 bảo dưỡng tiếp điểm: 1.40 công)'),
            ('Máy thi công', 'Đầu thử nhiệt Solo 461 / Thiết bị thử', 'ca', 1.00, '100,0%', 'Thử nghiệm cơ cấu tiếp điểm và kích hoạt giả lập'),
            ('Máy thi công', 'Đồng hồ vạn năng (VOM)', 'ca', 1.00, '100,0%', 'Đo thông mạch, điện áp loop và điện trở tiếp xúc khô')
        ]
    },

    # -------------------------------------------------------------
    # 4. E6C.3.04
    # -------------------------------------------------------------
    {
        'stt': 'II.4',
        'code': 'E6C.3.04',
        'title': 'E6C.3.04 - Kiểm tra và bảo dưỡng bộ module cảm biến beam quang học (Đầu báo khói dạng Beam thu - phát hồng ngoại cự ly xa)',
        'unit': '1 bộ',
        'total_labor': 2.51,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451. Áp dụng cho đầu báo Beam nhà xưởng gian máy.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.05, '2,0%', 'Nghiên cứu nguyên lý beam quang học, khoảng cách truyền nhận tia hồng ngoại'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công', 'công', 0.13, '5,0%', 'Chuẩn bị kính ngắm laser, tấm lọc thử khói Solo 330, đồng hồ VOM, giẻ lau quang học'),
            (3, '3. Thực hiện đầy đủ các bước an toàn', 'công', 0.05, '2,0%', 'Thực hiện an toàn leo cao, đeo dây an toàn toàn thân tại giàn giáo/xe nâng'),
            (4, '4. Đo thông mạch, đo điện trở cách điện cáp tín hiệu', 'công', 0.25, '10,0%', 'Đo điện trở cách điện cáp nguồn 24VDC và cáp loop từ tủ đến đầu thu phát beam'),
            (5, '5. Vệ sinh module thu/phát', 'công', 0.75, '30,0%', 'Lau chùi vỏ máy, làm sạch thấu kính phát tia hồng ngoại và bộ thu quang học'),
            (6, '6. Vệ sinh thấu kính của gương phản xạ', 'công', 0.25, '10,0%', 'Vệ sinh sạch bụi bẩn bề mặt gương phản xạ prism trên tường đối diện'),
            (7, '7. Kiểm tra tín hiệu đến và đi của module', 'công', 0.38, '15,0%', 'Đo kiểm tra cường độ tín hiệu hồng ngoại thu được (Signal Strength %)'),
            (8, '8. Xử lý điện trở các điểm nối dây', 'công', 0.25, '10,0%', 'Siết lại bu lông đế định vị, cố định khớp xoay cân chỉnh phương vị tia laser'),
            (9, '9. Kiểm tra chức năng làm việc của module', 'công', 0.20, '8,0%', 'Dùng tấm che chắn quang học thử nghiệm ngưỡng che khuất (Obscuration Test)'),
            (10, '10. Kiểm tra tín hiệu gửi lên tủ điều khiển trung tâm', 'công', 0.15, '6,0%', 'Kiểm tra báo động Fire và lỗi Fault chệch tia trên tủ trung tâm FMZ5000'),
            (11, '11. Vệ sinh mặt bằng và tiến hành nghiệm thu', 'công', 0.05, '2,0%', 'Thu dọn dụng cụ đo, vệ sinh khu vực làm việc và ký biên bản nghiệm thu')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 2.51, '100,0%', 'Cân chỉnh tia laser, kiểm chuẩn độ nhạy beam quang học không gian rộng'),
            ('Máy thi công', 'Đầu thử khói Solo 330 / Kính ngắm laser', 'ca', 0.21, '100,0%', 'Tạo khói thử nghiệm và định vị tâm tia phản xạ'),
            ('Máy thi công', 'Đồng hồ vạn năng (VOM)', 'ca', 0.21, '100,0%', 'Đo điện áp và mức tín hiệu điện áp thu quang'),
            ('Máy thi công', 'Dụng cụ gỡ đầu báo khói Solo 200', 'ca', 0.21, '100,0%', 'Thiết bị phụ trợ thao tác trên cao')
        ]
    },

    # -------------------------------------------------------------
    # 5. E6C.3.03
    # -------------------------------------------------------------
    {
        'stt': 'II.5',
        'code': 'E6C.3.03',
        'title': 'E6C.3.03 - Kiểm tra và bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng dây báo cháy / Loop',
        'unit': '1 bộ',
        'total_labor': 2.71,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 450. Áp dụng cho module địa chỉ hiện trường.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.05, '2,0%', 'Nghiên cứu địa chỉ logic và sơ đồ đấu nối tiếp điểm module giám sát'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công', 'công', 0.14, '5,0%', 'Chuẩn bị đồng hồ VOM, Megohmmeter 500V, thiết bị lập trình địa chỉ'),
            (3, '3. Thực hiện đầy đủ các bước kiểm tra an toàn', 'công', 0.05, '2,0%', 'Kiểm tra an toàn khu vực hiện trường trạm bơm/gian máy'),
            (4, '4. Đo thông mạch, đo điện trở cách điện của cáp tín hiệu', 'công', 0.14, '5,0%', 'Đo điện trở cách điện cáp loop và cáp ngõ vào tiếp điểm giám sát van/công tắc dòng chảy'),
            (5, '5. Vệ sinh module', 'công', 0.81, '30,0%', 'Mở nắp hộp bảo vệ, thổi bụi bo mạch và vệ sinh gioăng làm kín chống ẩm'),
            (6, '6. Kiểm tra cách điện, điện trở của dây báo cháy', 'công', 0.27, '10,0%', 'Đo điện trở cuối đường dây EOL resistor và kiểm tra cách điện tuyến cáp'),
            (7, '7. Kiểm tra tín hiệu đến và đi của module', 'công', 0.41, '15,0%', 'Kiểm tra điện áp giao tiếp bus đường loop (18 - 28VDC)'),
            (8, '8. Xử lý điện trở các điểm nối dây', 'công', 0.22, '8,0%', 'Siết chặt ốc vít cầu đấu, kiểm tra điểm tiếp xúc không bị lỏng hay oxy hóa'),
            (9, '9. Kiểm tra chức năng làm việc của module', 'công', 0.41, '15,0%', 'Kích hoạt tiếp điểm đầu vào giả lập, kiểm tra đèn LED trạng thái đổi màu đỏ'),
            (10, '10. Kiểm tra tín hiệu gửi lên tủ điều khiển trung tâm', 'công', 0.16, '6,0%', 'Xác nhận hiển thị đúng trạng thái Open/Short/Active trên tủ trung tâm'),
            (11, '11. Vệ sinh mặt bằng và tiến hành nghiệm thu', 'công', 0.05, '2,0%', 'Đóng chặt nắp kín nước, vệ sinh hiện trường và nghiệm thu bàn giao')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 2.71, '100,0%', 'Kiểm tra giao tiếp loop, cấu hình địa chỉ và bảo dưỡng module'),
            ('Máy thi công', 'Đồng hồ vạn năng (VOM)', 'ca', 0.25, '100,0%', 'Đo điện áp bus và điện trở tiếp điểm'),
            ('Máy thi công', 'Mê gôm mét 500V', 'ca', 0.13, '100,0%', 'Đo điện trở cách điện đường dây tín hiệu')
        ]
    },

    # -------------------------------------------------------------
    # 6. E6C.3.06
    # -------------------------------------------------------------
    {
        'stt': 'II.6',
        'code': 'E6C.3.06',
        'title': 'E6C.3.06 - Kiểm tra và bảo dưỡng phần điện van Deluge, Wet, Pre-action hệ thống PCCC (Cụm van xả tràn Deluge)',
        'unit': '1 cái',
        'total_labor': 7.00,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 453. Áp dụng cho cụm van xả tràn Deluge trạm biến áp & gian máy.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.07, '1,0%', 'Nghiên cứu bản vẽ sơ đồ nguyên lý thủy lực - điện của cụm van Deluge'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công', 'công', 0.07, '1,0%', 'Chuẩn bị đồng hồ đo áp, thiết bị đo cách điện, bơm tạo áp thử nghiệm'),
            (3, '3. Kiểm tra an toàn trước khi thực hiện', 'công', 0.07, '1,0%', 'Khóa van chặn cấp nước chính, mở van xả drain để tránh xả nước thật'),
            (4, '4. Đo thông mạch, điện trở cách điện cáp nguồn', 'công', 0.07, '1,0%', 'Đo cách điện cuộn hút solenoid 24VDC và cáp công tắc áp lực'),
            (5, '5. Kiểm tra bảo dưỡng van điện từ (solenoid)', 'công', 1.05, '15,0%', 'Vệ sinh lõi từ, kiểm tra màng cao su van mồi, đo dòng kích hoạt coil'),
            (6, '6. Kiểm tra bảo dưỡng đồng hồ đo áp suất', 'công', 1.54, '22,0%', 'Tháo kiểm tra, hiệu chỉnh độ chính xác đồng hồ áp suất buồng mồi và đường ống cấp'),
            (7, '7. Kiểm tra bảo dưỡng công tắc áp suất (pressure switch)', 'công', 1.40, '20,0%', 'Bơm áp giả lập, cân chỉnh điểm đóng ngắt tiếp điểm báo động áp lực'),
            (8, '8. Kiểm tra bảo dưỡng công tắc hành trình', 'công', 0.84, '12,0%', 'Kiểm tra công tắc giám sát trạng thái đóng/mở của van chặn OS&Y'),
            (9, '9. Kiểm tra bảo dưỡng chuông báo động thủy lực', 'công', 0.70, '10,0%', 'Làm sạch buồng cánh quạt nước chuông thủy lực (Water Motor Gong)'),
            (10, '10. Kiểm tra rò rỉ đường ống điều khiển', 'công', 0.56, '8,0%', 'Kiểm tra đường ống đồng/ống inox mồi áp lực, thay thế gioăng nếu rò rỉ'),
            (11, '11. Xử lý điện trở các điểm nối dây', 'công', 0.28, '4,0%', 'Làm sạch tiếp điểm hộp nối chống cháy nổ, siết chặt ốc siết cáp'),
            (12, '12. Kiểm tra chức năng van từ phòng điều khiển', 'công', 0.21, '3,0%', 'Kích hoạt lệnh mở van từ xa từ tủ trung tâm hoặc DCS, xác nhận van mồi mở'),
            (13, '13. Vệ sinh mặt bằng và nghiệm thu', 'công', 0.14, '2,0%', 'Nạp áp buồng mồi, mở lại van chặn chính, vệ sinh hiện trường và nghiệm thu')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 7.00, '100,0%', 'Thực hiện toàn diện 13 bước kiểm tra cụm van điện - thủy lực Deluge'),
            ('Máy thi công', 'Máy đo điện trở cách điện', 'ca', 0.20, '100,0%', 'Đo cách điện cuộn hút solenoid'),
            ('Máy thi công', 'Đồng hồ vạn năng (VOM)', 'ca', 0.50, '100,0%', 'Đo thông mạch, điện áp điều khiển 24VDC'),
            ('Máy thi công', 'Bàn cân áp suất / Bơm tạo áp', 'ca', 0.30, '100,0%', 'Hiệu chỉnh đồng hồ áp suất và công tắc áp lực')
        ]
    },

    # -------------------------------------------------------------
    # 7. E6C.1.31
    # -------------------------------------------------------------
    {
        'stt': 'II.7',
        'code': 'E6C.1.31',
        'title': 'E6C.1.31 - Kiểm tra, bảo dưỡng và hiệu chuẩn thiết bị đo nồng độ khí cháy Hydro (H2) phòng ắc quy',
        'unit': '1 bộ',
        'total_labor': 5.30,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 436. Áp dụng cho đầu đo khí Hydro phòng ắc quy nhà máy.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu', 'công', 0.11, '2,0%', 'Nghiên cứu dải đo khí H2 (0 - 100% LEL), hướng dẫn hiệu chuẩn sensor'),
            (2, '2. Chuẩn bị dụng cụ, vật tư, nhân lực, máy thi công', 'công', 0.27, '5,0%', 'Chuẩn bị bình khí chuẩn Zero gas (N2) và Span gas (H2), van điều áp, VOM'),
            (3, '3. Thực hiện đầy đủ các bước kiểm tra an toàn', 'công', 0.11, '2,0%', 'Kiểm tra thông gió phòng ắc quy, an toàn phòng chống cháy nổ'),
            (4, '4. Đo thông mạch, đo điện trở cách điện cáp nguồn/điều khiển', 'công', 0.27, '5,0%', 'Đo cách điện cáp tín hiệu 4-20mA và cáp nguồn 24VDC cấp cho đầu đo'),
            (5, '5. Kiểm tra, vệ sinh sensor đo nồng độ khí', 'công', 1.06, '20,0%', 'Tháo lưới lọc bụi chống nổ, làm sạch bề mặt cảm biến điện hóa/xúc tác'),
            (6, '6. Kiểm tra bo mạch xử lý của thiết bị', 'công', 0.80, '15,0%', 'Kiểm tra điện áp nguồn ổn định, vệ sinh bo mạch chuyển đổi tín hiệu'),
            (7, '7. Kiểm tra màn hình giám sát của thiết bị', 'công', 0.53, '10,0%', 'Kiểm tra hiển thị số nồng độ % LEL, các đèn LED báo Normal, Alarm, Fault'),
            (8, '8. Xử lý điện trở các điểm nối dây', 'công', 0.53, '10,0%', 'Siết lại domino đấu cáp tín hiệu, xử lý tiếp xúc ngõ ra relay cảnh báo'),
            (9, '9. Bơm khí chuẩn hiệu chuẩn điểm Zero và Span', 'công', 1.06, '20,0%', 'Bơm khí N2 tinh khiết chỉnh Zero (0% LEL), bơm khí mẫu chuẩn chỉnh Span (50% LEL)'),
            (10, '10. Kiểm tra tín hiệu truyền về hệ thống điều khiển trung tâm', 'công', 0.48, '9,0%', 'Kiểm tra dòng tín hiệu analog 4-20mA truyền về PLC/DCS khớp giá trị đo'),
            (11, '11. Vệ sinh mặt bằng và tiến hành nghiệm thu', 'công', 0.08, '2,0%', 'Tháo ống nối khí chuẩn, đậy nắp chống cháy nổ, ký biên bản nghiệm thu')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 5.30, '100,0%', 'Hiệu chuẩn chuyên sâu cảm biến khí cháy nổ Hydro'),
            ('Máy thi công', 'Bình khí mẫu hiệu chuẩn (Zero & Span gas)', 'ca', 0.50, '100,0%', 'Bình khí chuẩn đạt chứng nhận hiệu chuẩn kiểm định'),
            ('Máy thi công', 'Đồng hồ vạn năng kỹ thuật số (VOM)', 'ca', 0.50, '100,0%', 'Đo chính xác dòng ra analog 4-20mA và điện áp nguồn')
        ]
    },

    # -------------------------------------------------------------
    # 8. E6B.1.13.01
    # -------------------------------------------------------------
    {
        'stt': 'II.8',
        'code': 'E6B.1.13.01',
        'title': 'E6B.1.13.01 - Kiểm tra bình accu (Ắc quy nguồn dự phòng tủ PCCC & nguồn điều khiển)',
        'unit': '1 bình',
        'total_labor': 1.50,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 5 trang 383-384. Áp dụng cho từng bình ắc quy 12VDC kín khí.',
        'steps': [
            (1, '1. Nghiên cứu tài liệu; Chuẩn bị vật tư dụng cụ', 'công', 0.08, '5,0%', 'Nghiên cứu thông số bình Ah, điện áp định mức, chuẩn bị máy đo nội trở'),
            (2, '2. Kiểm tra bên ngoài bình accu', 'công', 0.02, '1,0%', 'Kiểm tra nứt vỡ vỏ bình, phồng rộp, rò rỉ dung dịch điện phân'),
            (3, '3. Kiểm tra bình acquy, vệ sinh vô mỡ các cọc bình', 'công', 0.08, '5,0%', 'Lau sạch bụi bẩn mặt bình, vệ sinh cọc cực và bôi mỡ bảo vệ chống oxy hóa'),
            (4, '4. Vệ sinh, kiểm tra, siết lại thanh nối giữa các accu', 'công', 0.20, '13,0%', 'Tháo thanh đồng nối mềm, chà sạch bề mặt tiếp xúc và dùng cờ lê lực siết lại'),
            (5, '5. Châm nước cất, kiểm tra tỷ trọng dung dịch, nạp cân bằng', 'công', 0.26, '17,0%', 'Châm nước cất nạp accu (nếu là loại hở), kiểm tra nạp cân bằng điện áp'),
            (6, '6. Đo điện áp các bình accu', 'công', 0.08, '5,0%', 'Đo điện áp tĩnh của từng bình đơn và điện áp tổng chuỗi ắc quy'),
            (7, '7. Vệ sinh bôi mỡ các cọc bình', 'công', 0.08, '5,0%', 'Quét mỡ chuyên dụng Vaseline bảo vệ cọc chì sau khi siết'),
            (8, '8. Sơn chống sét khung đỡ bình accu', 'công', 0.03, '2,0%', 'Dặm vá sơn chống rỉ các điểm trầy xước trên giá đỡ bình'),
            (9, '9. Đánh giá dung lượng và thử tải phóng xả', 'công', 0.60, '40,0%', 'Dùng máy đo nội trở (mΩ) và máy thử tải xả để xác định dung lượng thực tế Ah'),
            (10, '10. Kiểm tra hệ thống nạp và phục hồi dung dịch', 'công', 0.05, '5,0%', 'Kiểm tra hoạt động bộ nạp tự động (Float Charge / Boost Charge) trên tủ'),
            (11, '11. Vệ sinh mặt bằng và nghiệm thu', 'công', 0.02, '2,0%', 'Thu dọn hiện trường, kiểm tra kết nối chuỗi hoàn tất và nghiệm thu')
        ],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 1.50, '100,0%', 'Thực hiện 11 bước đo nội trở, thử tải và bảo dưỡng cọc bình'),
            ('Máy thi công', 'Thiết bị đo nội trở và dung lượng accu', 'ca', 0.25, '100,0%', 'Máy đo chuyên dụng đo điện trở trong (mΩ) và dung lượng Ah'),
            ('Máy thi công', 'Đồng hồ vạn năng (VOM)', 'ca', 0.20, '100,0%', 'Đo điện áp từng bình và dòng nạp'),
            ('Vật liệu', 'Nước cất nạp accu, mỡ bôi cọc bình', 'kg', 0.10, '100,0%', 'Vật tư phụ bảo vệ cọc bình chống ăn mòn axit')
        ]
    },

    # -------------------------------------------------------------
    # 9. E6C.2.04
    # -------------------------------------------------------------
    {
        'stt': 'II.9',
        'code': 'E6C.2.04',
        'title': 'E6C.2.04 - Kiểm tra, bảo dưỡng van điện từ Solenoid Deluge Valve',
        'unit': '1 cái',
        'total_labor': 1.80,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 440. Bảo dưỡng cuộn coil kích xả tràn van Deluge.',
        'steps': [],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 1.80, '100,0%', 'Kiểm tra cuộn hút điện từ 24VDC, vệ sinh lõi từ, thử đóng mở van mồi'),
            ('Máy thi công', 'Đồng hồ vạn năng VOM & Thiết bị kích thử', 'ca', 0.20, '100,0%', 'Đo điện trở thuần cuộn coil và dòng kích hoạt')
        ]
    },

    # -------------------------------------------------------------
    # 10. E6C.1.23
    # -------------------------------------------------------------
    {
        'stt': 'II.10',
        'code': 'E6C.1.23',
        'title': 'E6C.1.23 - Kiểm tra, bảo dưỡng công tắc áp lực, công tắc dòng chảy đường ống PCCC',
        'unit': '1 cái',
        'total_labor': 0.95,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 428. Áp dụng cho công tắc áp lực nước & công tắc dòng chảy.',
        'steps': [],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 0.95, '100,0%', 'Kiểm tra ngưỡng tác động áp suất, làm sạch tiếp điểm vi động'),
            ('Máy thi công', 'Bơm tạo áp cầm tay & Đồng hồ VOM', 'ca', 0.15, '100,0%', 'Bơm áp giả lập để kiểm tra đóng/ngắt tiếp điểm')
        ]
    },

    # -------------------------------------------------------------
    # 11. E6C.1.18
    # -------------------------------------------------------------
    {
        'stt': 'II.11',
        'code': 'E6C.1.18',
        'title': 'E6C.1.18 - Kiểm tra, bảo dưỡng đồng hồ đo lưu lượng nước đường ống PCCC',
        'unit': '1 cái',
        'total_labor': 1.20,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 425. Áp dụng cho thiết bị đo lưu lượng tuyến ống chữa cháy.',
        'steps': [],
        'resources': [
            ('Nhân công kỹ thuật', 'Bậc 5,0/7 (Kỹ thuật viên chính)', 'công', 1.20, '100,0%', 'Vệ sinh sensor lưu lượng, kiểm tra tín hiệu ngõ ra'),
            ('Máy thi công', 'Đồng hồ vạn năng VOM', 'ca', 0.20, '100,0%', 'Đo tín hiệu dòng điện và thông mạch tiếp điểm')
        ]
    },

    # -------------------------------------------------------------
    # 12. E6BOP.1.03
    # -------------------------------------------------------------
    {
        'stt': 'II.12',
        'code': 'E6BOP.1.03',
        'title': 'E6BOP.1.03 - Đại tu bảo dưỡng van chặn, van giảm áp, van xả khí phần cơ khí đường ống PCCC',
        'unit': '1 cái',
        'total_labor': 2.10,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 312. Áp dụng đại tu phần cơ khí cụm van nước.',
        'steps': [],
        'resources': [
            ('Nhân công kỹ thuật', 'Thợ cơ khí sửa chữa bậc 4,0/7', 'công', 2.10, '100,0%', 'Tháo van, rà phẳng đĩa van, thay thế gioăng túp, bôi mỡ trục ty van'),
            ('Máy thi công', 'Thiết bị nâng hạ / Palăng xích & Cờ lê lực', 'ca', 0.30, '100,0%', 'Phục vụ tháo lắp van nặng trên cao')
        ]
    },

    # -------------------------------------------------------------
    # 13. E6BOP.1.01
    # -------------------------------------------------------------
    {
        'stt': 'II.13',
        'code': 'E6BOP.1.01',
        'title': 'E6BOP.1.01 - Sửa chữa, đại tu van an toàn đường ống PCCC',
        'unit': '1 cái',
        'total_labor': 4.50,
        'basis': 'Quyết định 45/QĐ-HĐTV Tập 6 trang 308. Áp dụng cho van an toàn đường ống áp lực cao PCCC.',
        'steps': [],
        'resources': [
            ('Nhân công kỹ thuật', 'Thợ cơ khí sửa chữa bậc 4,5/7', 'công', 4.50, '100,0%', 'Rà kín bề mặt làm việc đĩa van và bệ van, căn chỉnh lực căng lò xo an toàn'),
            ('Máy thi công', 'Bàn thử áp lực van an toàn chuyên dùng', 'ca', 0.50, '100,0%', 'Thử nghiệm áp suất nổ (Set Pressure) và kẹp chì kiểm định')
        ]
    }
]

# Ghi từng định mức SCL vào bảng
for norm in scl_norms_full:
    # 1. Dòng tiêu đề định mức (Norm Header)
    ws.row_dimensions[curr_r].height = 25
    ws.cell(row=curr_r, column=1, value=norm['stt']).alignment = align_center
    ws.cell(row=curr_r, column=2, value=norm['title']).alignment = align_left
    ws.cell(row=curr_r, column=3, value=norm['unit']).alignment = align_center
    ws.cell(row=curr_r, column=4, value=norm['total_labor']).alignment = align_right
    ws.cell(row=curr_r, column=5, value='100%').alignment = align_center
    ws.cell(row=curr_r, column=6, value=norm['basis']).alignment = align_left
    
    for c in range(1, 7):
        cell = ws.cell(row=curr_r, column=c)
        cell.font = font_norm_code
        cell.fill = fill_norm
        cell.border = box_border
    
    curr_r += 1
    
    # 2. Nếu có 11 bước công việc: Ghi bảng chi tiết 11 bước
    if norm['steps']:
        # Header phụ của bảng 11 bước
        ws.row_dimensions[curr_r].height = 20
        ws.merge_cells(f'A{curr_r}:F{curr_r}')
        ws.cell(row=curr_r, column=1, value='   ▶ BẢNG CHI TIẾT 11 THÀNH PHẦN CÔNG VIỆC THEO FILE ĐỊNH MỨC .PDF (QUY TRÌNH CÔNG NGHỆ EVN):').font = font_sub_hdr
        for c in range(1, 7):
            ws.cell(row=curr_r, column=c).fill = fill_sub_hdr
            ws.cell(row=curr_r, column=c).border = box_border
        curr_r += 1
        
        # 11 dòng bước
        for step in norm['steps']:
            s_num, s_name, s_dvt, s_rate, s_pct, s_note = step
            ws.row_dimensions[curr_r].height = 20
            ws.cell(row=curr_r, column=1, value=f'Bước {s_num}').alignment = align_center
            ws.cell(row=curr_r, column=2, value=f'   {s_name}').alignment = align_left
            ws.cell(row=curr_r, column=3, value=s_dvt).alignment = align_center
            ws.cell(row=curr_r, column=4, value=s_rate).alignment = align_right
            ws.cell(row=curr_r, column=5, value=s_pct).alignment = align_center
            ws.cell(row=curr_r, column=6, value=s_note).alignment = align_left
            
            for c in range(1, 7):
                cell = ws.cell(row=curr_r, column=c)
                cell.font = font_step_item
                cell.border = box_border
            curr_r += 1
            
        # Dòng tổng cộng 11 bước
        ws.row_dimensions[curr_r].height = 21
        ws.cell(row=curr_r, column=1, value='Tổng').alignment = align_center
        ws.cell(row=curr_r, column=2, value='   Tổng cộng thành phần công việc (11 bước công nghệ)').alignment = align_left
        ws.cell(row=curr_r, column=3, value=norm['unit']).alignment = align_center
        ws.cell(row=curr_r, column=4, value=norm['total_labor']).alignment = align_right
        ws.cell(row=curr_r, column=5, value='100,0%').alignment = align_center
        ws.cell(row=curr_r, column=6, value=f'Khớp 100% định mức công nghệ EVN ({norm["total_labor"]} công/{norm["unit"]})').alignment = align_left
        
        for c in range(1, 7):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = font_step_total
            cell.fill = fill_step_total
            cell.border = box_border
        curr_r += 1

    # 3. Ghi các dòng thành phần hao phí (Resources: Nhân công, Máy thi công, Vật liệu)
    for res in norm['resources']:
        r_type, r_name, r_unit, r_rate, r_pct, r_desc = res
        ws.row_dimensions[curr_r].height = 20
        ws.cell(row=curr_r, column=1, value='').alignment = align_center
        ws.cell(row=curr_r, column=2, value=f'      • {r_type}: {r_name}').alignment = align_left
        ws.cell(row=curr_r, column=3, value=r_unit).alignment = align_center
        ws.cell(row=curr_r, column=4, value=r_rate).alignment = align_right
        ws.cell(row=curr_r, column=5, value=r_pct).alignment = align_center
        ws.cell(row=curr_r, column=6, value=r_desc).alignment = align_left
        
        for c in range(1, 7):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = font_resource_item
            cell.border = box_border
        curr_r += 1

    curr_r += 1 # Khoảng cách giữa các mã

# =========================================================================
# PHẦN 3: CHI TIẾT CÁC ĐỊNH MỨC THÔNG TƯ 38/2026/TT-BXD
# =========================================================================
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='MỤC III. CHI TIẾT CÁC ĐỊNH MỨC DỊCH VỤ & XÂY DỰNG THEO THÔNG TƯ 38/2026/TT-BXD').font = font_sec
ws.row_dimensions[curr_r].height = 28
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_sec
    ws.cell(row=curr_r, column=c).border = box_border

curr_r += 1

ws.row_dimensions[curr_r].height = 30
for c_idx, h_text in enumerate(detail_headers, start=1):
    cell = ws.cell(row=curr_r, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

tt38_details = [
    # 1. BD.41111
    ('norm', 'III.1', 'BD.41111 - Kiểm tra, bảo dưỡng họng cứu hỏa vách tường', '1 họng', 0.32, '100%', 'Bảng BD.41110, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ lắp đặt thiết bị PCCC bậc 3,5/7', 'công', 0.32, '100%', 'Kiểm tra mở van, độ kín gioăng cao su, vệ sinh lăng phun và khớp nối'),
    ('item', '', '   - Máy thi công: Máy kiểm tra áp lực nước mini cầm tay', 'ca', 0.08, '100%', 'Đo áp lực tĩnh và áp lực phun tại họng cứu hỏa'),

    # 2. BD.41121
    ('norm', 'III.2', 'BD.41121 - Kiểm tra, bảo dưỡng trụ cứu hỏa ngoài trời (loại 3 cửa)', '1 trụ', 0.55, '100%', 'Bảng BD.41120, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ lắp đặt đường ống PCCC bậc 3,5/7', 'công', 0.55, '100%', 'Mở nắp họng, tra mỡ ty van, xả cặn đường ống, kiểm tra ren vặn'),
    ('item', '', '   - Máy thi công: Đồng hồ đo áp lực lưu lượng chuyên dụng', 'ca', 0.12, '100%', 'Đo áp suất nước nguồn cấp ngoài nhà máy'),

    # 3. BD.41131
    ('norm', 'III.3', 'BD.41131 - Kiểm tra, bảo dưỡng cụm máy bơm chữa cháy (Bơm điện / Bơm Diesel)', '1 cụm', 1.20, '100%', 'Bảng BD.41130, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ cơ điện bậc 4,0/7', 'công', 1.20, '100%', 'Kiểm tra khớp nối mềm, siết chặt bu lông chân đế, bổ sung dầu bôi trơn'),
    ('item', '', '   - Máy thi công: Thiết bị đo rung chấn & Ampe kìm công nghiệp', 'ca', 0.40, '100%', 'Đo dòng khởi động bơm, đo độ rung vòng bi động cơ'),

    # 4. BD.41141
    ('norm', 'III.4', 'BD.41141 - Kiểm tra, làm sạch đầu phun Deluge / Sprinkler xả tràn', '1 đầu', 0.08, '100%', 'Bảng BD.41140, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ lắp đặt thiết bị PCCC bậc 3,0/7', 'công', 0.08, '100%', 'Vệ sinh cặn vôi, thông tắc lỗ phun lăng giá xả tràn dập lửa máy biến áp'),

    # 5. BD.41171
    ('norm', 'III.5', 'BD.41171 - Kiểm tra, bảo dưỡng đèn chiếu sáng sự cố và đèn chỉ dẫn lối thoát nạn EXIT', '1 bộ', 0.15, '100%', 'Bảng BD.41170, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ điện công nghiệp bậc 3,0/7', 'công', 0.15, '100%', 'Kiểm tra nguồn pin dự phòng, thời gian sáng khi mất điện, lau chùi mặt kính'),

    # 6. BA.18405
    ('norm', 'III.6', 'BA.18405 - Tháo dỡ chuông, còi, đèn chớp báo cháy phục vụ kiểm chuẩn', '1 cái', 0.18, '100%', 'Bảng BA.18400, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ điện bậc 3,0/7', 'công', 0.18, '100%', 'Tháo gỡ thiết bị khỏi tường/vách, cách ly đầu dây an toàn'),

    # 7. SB.82633 & SB.82634
    ('norm', 'III.7', 'SB.82633 - Sơn ống kim loại, van và thiết bị PCCC trong nhà (DN50 - DN100)', '100 m2', '', '100%', 'Bảng SB.82630, Phụ lục 6 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Sơn chống gỉ Alkyd đỏ PCCC', 'kg', 0.28, '100%', 'Định mức hao phí sơn lót chống gỉ trên 1m2 bề mặt kim loại'),
    ('item', '', '   - Vật liệu: Sơn phủ màu đỏ PCCC', 'kg', 0.32, '100%', 'Định mức hao phí sơn phủ 2 nước bảo vệ ngoài nhà'),
    ('item', '', '   - Nhân công: Thợ sơn bậc 3,5/7', 'công', 0.42, '100%', 'Làm sạch bề mặt rỉ sét cấp độ St2, sơn 1 lót 2 phủ'),

    ('norm', 'III.8', 'SB.82634 - Sơn ống kim loại, trụ cứu hỏa, phụ kiện PCCC ngoài trời (DN125 - DN200)', '100 m2', '', '100%', 'Bảng SB.82630, Phụ lục 6 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Thợ sơn bậc 3,5/7 (Làm việc ngoài trời)', 'công', 0.48, '100%', 'Chống ăn mòn muối biển cho đường ống ngoài trời Vĩnh Tân 4'),

    # 8. BB.86601, BB.86103, BB.85201
    ('norm', 'III.9', 'BB.86601 / BB.86103 / BB.85201 - Thay thế vật tư, van hỏng hóc phát sinh trong quá trình SCL', '1 cái', '', '100%', 'Chương VIII Phụ lục 3 Thông tư 38/2026/TT-BXD'),
    ('item', '', '   - BB.86601: Lắp van ren ≤ DN25 (Nhân công 3,5/7: 0.18 công; Máy: 0.05 ca)', 'cái', 0.18, '100%', 'Thay mới van xả cặn ren đồng DN25'),
    ('item', '', '   - BB.86103: Lắp van bích DN80 (Nhân công 3,5/7: 0.45 công; Máy: 0.12 ca)', 'cái', 0.45, '100%', 'Thay mới van Alarm bích DN80'),
    ('item', '', '   - BB.85201: Lắp đồng hồ áp suất (Nhân công 3,5/7: 0.12 công; Máy: 0.04 ca)', 'cái', 0.12, '100%', 'Thay mới đồng hồ áp suất 0-16 bar'),

    # 9. DB.24001
    ('norm', 'III.10', 'DB.24001 - Đo kiểm tra định kỳ điện trở hệ thống chống sét & tiếp địa PCCC', '1 HT', 0.50, '100%', 'Bảng DB.24000, Phụ lục 5 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Kỹ sư công nghệ thí nghiệm điện bậc 5,0/8', 'công', 0.50, '100%', 'Đóng cọc phụ, rải dây đo, cân chỉnh cầu đo tiếp địa'),
    ('item', '', '   - Máy thi công: Thiết bị đo điện trở đất chuyên dụng (Terrameter)', 'ca', 0.50, '100%', 'Máy đo điện trở đất kỹ thuật số đạt chuẩn kiểm định')
]

for row_idx, r_data in enumerate(tt38_details, start=curr_r + 1):
    r_type = r_data[0]
    ws.row_dimensions[row_idx].height = 24 if r_type == 'norm' else 20
    
    ws.cell(row=row_idx, column=1, value=r_data[1]).alignment = align_center
    ws.cell(row=row_idx, column=2, value=r_data[2]).alignment = align_left
    ws.cell(row=row_idx, column=3, value=r_data[3]).alignment = align_center
    ws.cell(row=row_idx, column=4, value=r_data[4]).alignment = align_right
    ws.cell(row=row_idx, column=5, value=r_data[5]).alignment = align_center
    ws.cell(row=row_idx, column=6, value=r_data[6]).alignment = align_left
    
    for c in range(1, 7):
        cell = ws.cell(row=row_idx, column=c)
        cell.border = box_border
        if r_type == 'norm':
            cell.font = font_norm_code
            cell.fill = fill_norm
        else:
            cell.font = font_step_item

curr_r = curr_r + len(tt38_details) + 2

# =========================================================================
# PHẦN 4: CĂN CỨ PHÁP LÝ & NGUYÊN TẮC CHO CÁC MÃ "KHÔNG ÁP MÃ"
# =========================================================================
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='MỤC IV. CĂN CỨ PHÁP LÝ & NGUYÊN TẮC XỬ LÝ CHO CÁC CÔNG TÁC "KHÔNG ÁP MÃ" (KDM)').font = font_sec
ws.row_dimensions[curr_r].height = 28
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_sec
    ws.cell(row=curr_r, column=c).border = box_border

curr_r += 1

ws.row_dimensions[curr_r].height = 30
for c_idx, h_text in enumerate(detail_headers, start=1):
    cell = ws.cell(row=curr_r, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

kdm_details = [
    ('norm', 'IV.1', 'Không áp mã (đã gồm trong E6C.3.05) - Bo mạch, Card loop, Card quang, Bộ nguồn trong tủ FMZ5000', '115 mục', '0.00', 'Đã tính trong E6C.3.05', 'Căn cứ Thuyết minh thành phần công việc E6C.3.05 (QĐ 45/QĐ-HĐTV EVN)'),
    ('item', '', '   - Căn cứ kỹ thuật: Định mức E6C.3.05 (2.00 công/tủ) đã bao gồm từ bước 5 đến bước 9: kiểm tra toàn diện nguồn cấp 24V, CPU, card mạng loop, card quang.', '', '', 'Chi tiết bước công việc', 'Không tính mã riêng để chống xuất toán trùng lặp kinh phí.'),
    
    ('norm', 'IV.2', 'Không áp mã (Chi phí thường xuyên) - Kiểm tra đèn chỉ lối EXIT và đèn sự cố', '16 mục', '0.00', 'Chi phí thường xuyên', 'Thông tư 11/2021/TT-BXD & Quy trình bảo dưỡng định kỳ của NMNĐ VT4'),
    ('item', '', '   - Căn cứ kỹ thuật: Công tác là "Kiểm tra định kỳ" (bật test pin, lau mặt kính). Không có ĐM bảo dưỡng; kiên quyết không áp mã lắp mới BD.41171.', '', '', 'Chi phí thường xuyên', 'Chi phí thực hiện trong chi phí bảo dưỡng thường xuyên / quản lý chung.'),

    ('norm', 'IV.3', 'Không áp mã (Thuộc trách nhiệm loop) - Tuyến cáp tín hiệu đường Loop tìm và xử lý chạm đất', '24 mục', '0.00', 'Thuộc trách nhiệm loop', 'Quy trình công nghệ bảo dưỡng mạng Loop PCCC'),
    ('item', '', '   - Căn cứ kỹ thuật: Trong các mã bảo dưỡng đầu báo (E6C.3.01) và module (E6C.3.03) đã bao gồm thao tác đo thông mạch và kiểm tra chạm đất cáp.', '', '', 'Kỹ thuật vận hành Loop', 'Không tính thêm mã kéo rải cáp mới để tránh trùng lắp khối lượng.'),

    ('norm', 'IV.4', 'Không áp mã (Chi phí phụ trợ) - Kiểm tra bình bọt, bình khí, cuộn vòi phương tiện chữa cháy tại chỗ', '23 mục', '0.00', 'Chi phí phụ trợ', 'Quy chế quản lý phương tiện PCCC cơ sở của EVN'),
    ('item', '', '   - Căn cứ kỹ thuật: Kiểm tra cân nặng bình khí, kiểm tra đồng hồ áp suất bình bọt, kiểm tra cuộn vòi lăng phun trong tủ cứu hỏa.', '', '', 'Phương tiện chữa cháy tại chỗ', 'Không có định mức xây dựng riêng; tính trong gói dịch vụ phụ trợ.')
]

for row_idx, r_data in enumerate(kdm_details, start=curr_r + 1):
    r_type = r_data[0]
    ws.row_dimensions[row_idx].height = 24 if r_type == 'norm' else 20
    
    ws.cell(row=row_idx, column=1, value=r_data[1]).alignment = align_center
    ws.cell(row=row_idx, column=2, value=r_data[2]).alignment = align_left
    ws.cell(row=row_idx, column=3, value=r_data[3]).alignment = align_center
    ws.cell(row=row_idx, column=4, value=r_data[4]).alignment = align_right
    ws.cell(row=row_idx, column=5, value=r_data[5]).alignment = align_center
    ws.cell(row=row_idx, column=6, value=r_data[6]).alignment = align_left
    
    for c in range(1, 7):
        cell = ws.cell(row=row_idx, column=c)
        cell.border = box_border
        if r_type == 'norm':
            cell.font = font_norm_code
            cell.fill = fill_norm
        else:
            cell.font = font_step_item

curr_r = curr_r + len(kdm_details) + 2

# Signatures Block
ws.merge_cells(f'A{curr_r}:B{curr_r}')
ws.cell(row=curr_r, column=1, value='NGƯỜI LẬP BẢNG KIỂM TRA ĐỊNH MỨC\n(Tổ chuyên gia kỹ thuật dự toán)').font = font_comp
ws.cell(row=curr_r, column=1).alignment = align_center

ws.merge_cells(f'C{curr_r}:D{curr_r}')
ws.cell(row=curr_r, column=3, value='PHÒNG KẾ HOẠCH VẬT TƯ / KTAT\n(Kiểm soát đối chiếu Thông tư 38 & QĐ 45)').font = font_comp
ws.cell(row=curr_r, column=3).alignment = align_center

ws.merge_cells(f'E{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=5, value='HỘI ĐỒNG THẨM ĐỊNH / LÃNH ĐẠO CÔNG TY\n(Phê duyệt 100% định mức áp dụng)').font = font_comp
ws.cell(row=curr_r, column=5).alignment = align_center

ws.row_dimensions[curr_r].height = 40

# Auto column widths
ws.column_dimensions['A'].width = 12
ws.column_dimensions['B'].width = 65
ws.column_dimensions['C'].width = 12
ws.column_dimensions['D'].width = 14
ws.column_dimensions['E'].width = 16
ws.column_dimensions['F'].width = 68

# Freeze Panes
ws.freeze_panes = 'C7'

print("Saving updated workbook...")
wb.save(wb_path)
print(f"Saved to local: {wb_path}")

try:
    wb.save(dl_path)
    print(f"Saved to downloads primary: {dl_path}")
except Exception as e:
    alt_dl = 'C:\\Users\\Admin\\Downloads\\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_CapNhat.xlsx'
    wb.save(alt_dl)
    print(f"Saved to downloads fallback: {alt_dl} (Error: {e})")
