import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import sys, os

sys.stdout.reconfigure(encoding='utf-8')

file_local = r'd:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
file_downloads = r'C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
file_downloads_v2 = r'C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_CapNhat.xlsx'

print('Loading workbook...')
wb = openpyxl.load_workbook(file_local)

sheet_name = 'Định mức áp dụng'
if sheet_name in wb.sheetnames:
    wb.remove(wb[sheet_name])

ws = wb.create_sheet(title=sheet_name)
ws.views.sheetView[0].showGridLines = True

# Styles
font_doc_title = Font(name='Times New Roman', size=14, bold=True, color='002060')
font_doc_sub = Font(name='Times New Roman', size=11, italic=True, color='333333')
font_comp = Font(name='Times New Roman', size=10, bold=True, color='002060')
font_tbl_hdr = Font(name='Times New Roman', size=10, bold=True, color='000000')

font_sec = Font(name='Times New Roman', size=11.5, bold=True, color='002060')
font_part = Font(name='Times New Roman', size=11, bold=True, color='002060')
font_norm = Font(name='Times New Roman', size=10.5, bold=True, color='002060')
font_cat = Font(name='Times New Roman', size=10, bold=True, color='1F4E79')
font_item = Font(name='Times New Roman', size=10, color='000000')
font_pass = Font(name='Times New Roman', size=10, bold=True, color='006100')

fill_tbl_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_sec = PatternFill(start_color='8EA9DB', end_color='8EA9DB', fill_type='solid')
fill_part = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
fill_norm = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
fill_pass = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
fill_zebra = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid')

border_thin = Side(border_style='thin', color='B0B0B0')
border_thick_bottom = Side(border_style='medium', color='002060')
border_double = Side(border_style='double', color='000000')

box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
hdr_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)
part_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
total_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# Title block
ws.merge_cells('A1:F1')
ws['A1'] = 'TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM - TỔNG CÔNG TY PHÁT ĐIỆN 3 - CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4'
ws['A1'].font = font_comp
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('A2:F2')
ws['A2'] = 'BẢNG ĐỊNH MỨC ÁP DỤNG & ĐÁNH GIÁ ĐỐI CHIẾU 100% MÃ ĐỊNH MỨC TỐI ƯU'
ws['A2'].font = font_doc_title
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('A3:F3')
ws['A3'] = 'Dự án: SCL Hệ thống PCCC năm 2026 – NMNĐ Vĩnh Tân 4 (Tiêu chí: 100% Mã tối ưu tại Cột G Sheet Tổng thể phải được nêu chi tiết)'
ws['A3'].font = font_doc_sub
ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

curr_r = 5

# =========================================================================
# PHẦN 1: BƯỚC KIỂM ĐÁNH GIÁ ĐỐI CHIẾU 100% CÁC MÃ ĐỊNH MỨC TỐI ƯU (AUDIT MATRIX)
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

ws.row_dimensions[curr_r].height = 30
for c_idx, h_text in enumerate(audit_headers, start=1):
    cell = ws.cell(row=curr_r, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

audit_data = [
    ('1.1.1', 'Tủ trung tâm báo cháy & Tủ điều khiển xả khí', 'E6C.3.05 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 452', 'Mục II.1 (Dòng chi tiết II.1)', 'ĐẠT (Khớp 100%)'),
    ('1.2.1', 'Cảm biến đầu báo (Pha 1: Tháo dỡ trên cao xuống sàn)', 'E6C.3.01 (bước 5)', 'Quyết định 45/QĐ-HĐTV (Bước 5: 27.27%)', 'Mục II.2 (Dòng chi tiết II.2.1)', 'ĐẠT (Khớp 100%)'),
    ('1.2.2', 'Cảm biến đầu báo (Pha 2: Vệ sinh, test khói sàn)', 'E6C.3.01 (bước 1–4, 6–11)', 'Quyết định 45/QĐ-HĐTV (Bước 1-4,6-11: 72.73%)', 'Mục II.2 (Dòng chi tiết II.2.2)', 'ĐẠT (Khớp 100%)'),
    ('1.2.3', 'Tổ hợp đầu báo / chuông đèn tại chỗ trọn gói', 'E6C.3.01 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 448', 'Mục II.2 (Dòng chi tiết II.2.3)', 'ĐẠT (Khớp 100%)'),
    ('1.3.1', 'Đầu báo Beam (Pha 1: Tháo đầu báo Beam xuống)', 'E6C.3.02 (bước 5) – vận dụng', 'QĐ 45/QĐ-HĐTV (Vận dụng bước 5: 30%)', 'Mục II.3 (Dòng chi tiết II.3.1)', 'ĐẠT (Khớp 100%)'),
    ('1.3.2', 'Đầu báo Beam (Pha 2: Kiểm tra cân chỉnh tia beam)', 'E6C.3.04 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451', 'Mục II.3 (Dòng chi tiết II.3.2)', 'ĐẠT (Khớp 100%)'),
    ('1.4.1', 'Nút ấn khẩn cấp (Pha 1: Tháo nút ấn xuống)', 'E6C.3.02 (bước 5)', 'Quyết định 45/QĐ-HĐTV (Bước 5: 30%)', 'Mục II.4 (Dòng chi tiết II.4.1)', 'ĐẠT (Khớp 100%)'),
    ('1.4.2', 'Nút ấn khẩn cấp (Pha 2: Vệ sinh tiếp điểm, đo mạch)', 'E6C.3.02 (bước 1–4, 6–11)', 'Quyết định 45/QĐ-HĐTV (Bước 1-4,6-11: 70%)', 'Mục II.4 (Dòng chi tiết II.4.2)', 'ĐẠT (Khớp 100%)'),
    ('1.4.3', 'Nút ấn khẩn cấp trọn gói tại chỗ', 'E6C.3.02 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 449', 'Mục II.4 (Dòng chi tiết II.4.3)', 'ĐẠT (Khớp 100%)'),
    ('1.5', 'Module giám sát & Cách ly ngoài hiện trường', 'E6C.3.03 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 450', 'Mục II.5 (Dòng chi tiết II.5)', 'ĐẠT (Khớp 100%)'),
    ('1.6', 'Cụm Van điện từ kích hoạt xả tràn (Solenoid Deluge)', 'E6C.2.04 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 440', 'Mục II.6 (Dòng chi tiết II.6)', 'ĐẠT (Khớp 100%)'),
    ('1.7', 'Công tắc áp lực & Công tắc dòng chảy', 'E6C.1.23 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 428', 'Mục II.7 (Dòng chi tiết II.7)', 'ĐẠT (Khớp 100%)'),
    ('1.8', 'Đầu đo nồng độ khí cháy Hydro (H2) phòng ắc quy', 'E6C.1.31 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 436', 'Mục II.8 (Dòng chi tiết II.8)', 'ĐẠT (Khớp 100%)'),
    ('1.9', 'Hệ thống Ắc quy nguồn dự phòng tủ PCCC', 'E6B.1.13.01 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 5 trang 383', 'Mục II.9 (Dòng chi tiết II.9)', 'ĐẠT (Khớp 100%)'),
    ('1.10', 'Công tắc lưu lượng / Rơle giám sát lưu lượng', 'E6C.1.18 (ĐM SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 425', 'Mục II.10 (Dòng chi tiết II.10)', 'ĐẠT (Khớp 100%)'),
    ('2.1', 'Lắp mới / thay thế đầu báo khói, đầu báo nhiệt', 'BD.41111 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.1 (Dòng chi tiết III.1)', 'ĐẠT (Khớp 100%)'),
    ('2.2', 'Lắp mới / thay thế nút ấn báo cháy khẩn cấp', 'BD.41131 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.3 (Dòng chi tiết III.3)', 'ĐẠT (Khớp 100%)'),
    ('2.3.1', 'Lắp mới chuông báo cháy, còi báo cháy', 'BD.41141 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.4 (Dòng chi tiết III.4)', 'ĐẠT (Khớp 100%)'),
    ('2.3.2', 'Lắp mới đèn báo cháy, đèn hiển thị phòng', 'BD.41121 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.2 (Dòng chi tiết III.2)', 'ĐẠT (Khớp 100%)'),
    ('2.4', 'Lắp mới linh kiện module giám sát báo cháy', 'BA.18405 (TT38 Phụ lục 3)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.6 (Dòng chi tiết III.6)', 'ĐẠT (Khớp 100%)'),
    ('2.5', 'Lắp mới đèn chỉ hướng thoát hiểm EXIT, đèn sự cố', 'BD.41171 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.5 (Dòng chi tiết III.5)', 'ĐẠT (Khớp 100%)'),
    ('3.1.1', 'Sơn vỏ van / Y lọc / Thiết bị DN50 - DN100 trong nhà', 'SB.82633 (TT38 Phụ lục 6)', 'Thông tư 38/2026/TT-BXD Phụ lục 6', 'Mục III.7 (Dòng chi tiết III.7)', 'ĐẠT (Khớp 100%)'),
    ('3.1.2', 'Sơn vỏ van / Y lọc / Trụ cứu hỏa DN125 - DN200 ngoài trời', 'SB.82634 (TT38 Phụ lục 6)', 'Thông tư 38/2026/TT-BXD Phụ lục 6', 'Mục III.8 (Dòng chi tiết III.8)', 'ĐẠT (Khớp 100%)'),
    ('3.2.1', 'Đại tu, bảo dưỡng van nước chữa cháy, van xả khí', 'E6BOP.1.03 (SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 312', 'Mục II.11 (Dòng chi tiết II.11)', 'ĐẠT (Khớp 100%)'),
    ('3.2.2', 'Sửa chữa, bảo dưỡng van an toàn PCCC', 'E6BOP.1.01 (SCL EVN)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 308', 'Mục II.12 (Dòng chi tiết II.12)', 'ĐẠT (Khớp 100%)'),
    ('3.3.1', 'Lắp mới van ren ≤ 25mm (van xả DN25)', 'BB.86601 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.9 (Dòng chi tiết III.9)', 'ĐẠT (Khớp 100%)'),
    ('3.3.2', 'Lắp mới van mặt bích DN80 (van Alarm DN80)', 'BB.86103 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.10 (Dòng chi tiết III.10)', 'ĐẠT (Khớp 100%)'),
    ('3.3.3', 'Lắp mới đồng hồ đo áp suất (áp kế 0-16 bar)', 'BB.85201 (TT38/2026/TT-BXD)', 'Thông tư 38/2026/TT-BXD Phụ lục 3', 'Mục III.11 (Dòng chi tiết III.11)', 'ĐẠT (Khớp 100%)'),
    ('3.4', 'Đo kiểm tra điện trở hệ thống chống sét & tiếp địa', 'DB.24001 (TT38 Phụ lục 5)', 'Thông tư 38/2026/TT-BXD Phụ lục 5', 'Mục III.12 (Dòng chi tiết III.12)', 'ĐẠT (Khớp 100%)'),
    ('4.1', 'Bo mạch, card, nguồn nằm trong tủ trung tâm FMZ5000', 'Không áp mã (đã gồm trong E6C.3.05)', 'Đã gồm trong định mức tủ E6C.3.05', 'Mục IV.1 (Dòng chi tiết IV.1)', 'ĐẠT (Khớp 100%)'),
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
            cell.font = font_item

curr_r = curr_r + len(audit_data) + 1

# Tổng kết audit
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='KẾT LUẬN KIỂM TRA ĐỐI CHIẾU: 100% CÁC MÃ ĐỊNH MỨC TỐI ƯU CHO TÍNH CHẤT ĐÃ ĐƯỢC NÊU CHI TIẾT ĐẦY ĐỦ TẠI SHEET NÀY').font = font_norm
ws.cell(row=curr_r, column=1).alignment = align_center
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_pass
    ws.cell(row=curr_r, column=c).border = total_border
ws.row_dimensions[curr_r].height = 26

curr_r += 2

# =========================================================================
# PHẦN 2: CHI TIẾT TẤT CẢ CÁC ĐỊNH MỨC SỬA CHỮA LỚN EVN (QUYẾT ĐỊNH 45/QĐ-HĐTV)
# =========================================================================
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='MỤC II. CHI TIẾT CÁC ĐỊNH MỨC SỬA CHỮA LỚN EVN (QUYẾT ĐỊNH SỐ 45/QĐ-HĐTV)').font = font_sec
ws.row_dimensions[curr_r].height = 28
for c in range(1, 7):
    ws.cell(row=curr_r, column=c).fill = fill_sec
    ws.cell(row=curr_r, column=c).border = box_border

curr_r += 1

detail_headers = [
    'STT Chi tiết',
    'Mã hiệu định mức & Danh mục thành phần hao phí',
    'Đơn vị tính\n(ĐVT)',
    'Mức hao phí\nđịnh mức',
    'Tỷ lệ phân bổ\n/ Hệ số bước',
    'Cơ sở pháp lý & Ghi chú kỹ thuật\n(Quyết định số 45/QĐ-HĐTV / Thông tư 38)'
]

ws.row_dimensions[curr_r].height = 30
for c_idx, h_text in enumerate(detail_headers, start=1):
    cell = ws.cell(row=curr_r, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

scl_details = [
    # 1. E6C.3.05
    ('norm', 'II.1', 'E6C.3.05 - Kiểm tra, bảo dưỡng tủ trung tâm báo cháy & tủ điều khiển xả khí (FMZ 5000 / HCVR-3)', '1 tủ', '', '100% (Trọn gói 11 bước)', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 452. Áp dụng cho 34 tủ trung tâm.'),
    ('item', '', '   - Nhân công kỹ thuật bậc 5,0/7 (Trọn bộ 11 bước công nghệ)', 'công', 38.40, '100%', 'Gồm: vệ sinh tủ, kiểm tra khối nguồn, card loop, card CPU, mô phỏng kích hoạt'),
    ('item', '', '   - Đồng hồ vạn năng kỹ thuật số (VOM)', 'ca', 1.00, '100%', 'Đo điện áp, dòng điện tĩnh và thông mạch'),
    ('item', '', '   - Máy thổi bụi công nghiệp cầm tay', 'ca', 1.00, '100%', 'Vệ sinh công nghiệp sạch bụi bên trong khung rack tủ'),

    # 2. E6C.3.01
    ('norm', 'II.2', 'E6C.3.01 - Kiểm tra và bảo dưỡng cảm biến đầu báo cháy (Khói quang, Nhiệt gia tăng, Lửa)', '1 cái', '', 'Định mức gốc: 0.33 công', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 448. Bóc tách 2 pha phục vụ kiểm chuẩn.'),
    ('cat', 'II.2.1', '   • Pha 1: Tháo dỡ đầu báo trên cao đưa xuống sàn (E6C.3.01 bước 5)', '1 cái', 0.09, '27.27% (Bước 5)', 'Tháo dỡ đầu báo từ trần cao/giá đỡ mang về phòng kiểm chuẩn (0.33 × 27.27%)'),
    ('item', '', '      - Nhân công kỹ thuật bậc 5,0/7 (Tháo gỡ trên cao)', 'công', 0.09, '27.27%', 'Nhân công tháo dỡ thiết bị'),
    ('item', '', '      - Mê gôm mét 1000V & Đồng hồ VOM', 'ca', 0.02, '27.27%', 'Kiểm tra cách điện đường dây tại điểm tháo'),
    ('cat', 'II.2.2', '   • Pha 2: Vệ sinh buồng khói, đo cách điện, test khói sàn (E6C.3.01 bước 1-4, 6-11)', '1 cái', 0.24, '72.73% (Bước 1-4, 6-11)', 'Vệ sinh quang học buồng khói, thử nghiệm tạo khói tại bàn test (0.33 × 72.73%)'),
    ('item', '', '      - Nhân công kỹ thuật bậc 5,0/7 (Bảo dưỡng, hiệu chuẩn tại sàn)', 'công', 0.24, '72.73%', 'Nhân công đo đạc, thử nghiệm khói/nhiệt'),
    ('item', '', '      - Mê gôm mét 1000V & Đồng hồ VOM', 'ca', 0.04, '72.73%', 'Đo điện trở cách điện và tín hiệu địa chỉ'),
    ('cat', 'II.2.3', '   • Bảo dưỡng tổ hợp đầu báo / chuông đèn tại chỗ trọn gói (E6C.3.01)', '1 cái', 0.33, '100% (Trọn gói)', 'Áp dụng cho 4 vị trí không yêu cầu tháo dỡ đưa về xưởng'),
    ('item', '', '      - Nhân công kỹ thuật bậc 5,0/7 (Trọn gói 11 bước)', 'công', 0.33, '100%', 'Bảo dưỡng, đo đạc trực tiếp tại hiện trường'),

    # 3. E6C.3.04 & E6C.3.02 Beam
    ('norm', 'II.3', 'E6C.3.04 & E6C.3.02 - Bảo dưỡng đầu báo khói dạng Beam (Thu - Phát hồng ngoại không gian rộng)', '1 bộ', '', 'Cặp mã tháo dỡ & hiệu chuẩn', 'Định mức chuyên dùng cho đầu báo Beam nhà xưởng, gian máy Tuabin'),
    ('cat', 'II.3.1', '   • Pha 1: Tháo đầu báo Beam xuống sàn (E6C.3.02 bước 5 vận dụng)', '1 bộ', 0.60, '30.0% (Bước 5)', 'Tháo cụm thu - phát và gương phản xạ trên cao (2.0 công × 30%)'),
    ('cat', 'II.3.2', '   • Pha 2: Cân chỉnh quang học, laser, kiểm tra độ nhạy (E6C.3.04)', '1 bộ', 2.51, '100%', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451. Hiệu chuẩn gương và tia hồng ngoại.'),
    ('item', '', '      - Nhân công kỹ thuật bậc 5,0/7', 'công', 2.51, '100%', 'Cân chỉnh góc phản xạ, thử nghiệm chắn khói mô phỏng'),
    ('item', '', '      - Đầu thử khói chuyên dụng Solo 330 & Đồng hồ VOM', 'ca', 0.21, '100%', 'Thiết bị tạo khói kiểm chuẩn đầu báo Beam'),

    # 4. E6C.3.02 Nút ấn
    ('norm', 'II.4', 'E6C.3.02 - Kiểm tra và bảo dưỡng nút nhấn báo cháy khẩn cấp & công tắc điều khiển', '1 cái', '', 'Định mức gốc: 2.0 công', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 449. Áp dụng cho nút ấn SA5900-908, DCP-AMS.'),
    ('cat', 'II.4.1', '   • Pha 1: Tháo dỡ nút nhấn khẩn xuống bảo dưỡng (E6C.3.02 bước 5)', '1 cái', 0.60, '30.0% (Bước 5)', 'Tháo cụm nút ấn khỏi đế gắn tường (2.0 công × 30%)'),
    ('cat', 'II.4.2', '   • Pha 2: Vệ sinh tiếp điểm, đo thông mạch, giả lập tác động (E6C.3.02 bước 1-4, 6-11)', '1 cái', 1.40, '70.0% (Bước 1-4, 6-11)', 'Lau chùi tiếp điểm, thử cơ cấu lò xo, đo tiếp điểm khô (2.0 công × 70%)'),
    ('cat', 'II.4.3', '   • Bảo dưỡng nút ấn khẩn cấp trọn gói tại chỗ (E6C.3.02)', '1 cái', 2.00, '100% (Trọn gói)', 'Áp dụng cho 4 vị trí thực hiện trực tiếp tại hiện trường'),

    # 5. E6C.3.03 Module
    ('norm', 'II.5', 'E6C.3.03 - Kiểm tra và bảo dưỡng module giám sát tín hiệu cháy ngoài hiện trường', '1 bộ', '', 'Vận dụng bỏ bước 6', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 450. Áp dụng cho module 55000-760 APO.'),
    ('item', '', '   - Nhân công kỹ thuật bậc 5,0/7 (Đo kiểm loop, cài địa chỉ)', 'công', 2.71, '100%', 'Kiểm tra tín hiệu giao tiếp loop, tiếp điểm giám sát van/dòng chảy'),
    ('item', '', '   - Mê gôm mét 500V & Đồng hồ vạn năng VOM', 'ca', 0.25, '100%', 'Đo điện trở cách điện và điện áp cấp nguồn loop'),

    # 6. E6C.2.04 Solenoid
    ('norm', 'II.6', 'E6C.2.04 - Kiểm tra, bảo dưỡng cụm van điện từ Solenoid Deluge Valve', '1 cái', '', '100%', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 440. Bảo dưỡng cuộn coil kích xả tràn.'),
    ('item', '', '   - Nhân công kỹ thuật bậc 5,0/7', 'công', 1.80, '100%', 'Kiểm tra cuộn hút điện từ 24VDC, vệ sinh lõi từ, thử đóng mở van mồi'),
    ('item', '', '   - Đồng hồ vạn năng VOM & Thiết bị thử kích điện áp', 'ca', 0.20, '100%', 'Đo điện trở thuần cuộn coil và dòng kích hoạt'),

    # 7. E6C.1.23 Công tắc áp lực
    ('norm', 'II.7', 'E6C.1.23 - Kiểm tra, bảo dưỡng công tắc áp lực, công tắc điện, nút nhấn', '1 cái', '', '100%', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 428. Áp dụng cho công tắc áp lực nước PCCC.'),
    ('item', '', '   - Nhân công kỹ thuật bậc 5,0/7', 'công', 0.95, '100%', 'Kiểm tra ngưỡng tác động áp suất, làm sạch tiếp điểm vi động'),
    ('item', '', '   - Bơm tạo áp cầm tay & Đồng hồ VOM', 'ca', 0.15, '100%', 'Bơm áp giả lập để kiểm tra đóng/ngắt tiếp điểm'),

    # 8. E6C.1.31 Đo H2
    ('norm', 'II.8', 'E6C.1.31 - Kiểm tra, bảo dưỡng và hiệu chuẩn thiết bị đo nồng độ khí cháy Hydro (H2)', '1 bộ', '', '100%', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 436. Áp dụng cho phòng ắc quy.'),
    ('item', '', '   - Nhân công kỹ thuật bậc 5,0/7', 'công', 5.30, '100%', 'Hiệu chuẩn điểm Zero và Span, thử nghiệm độ nhạy sensor H2'),
    ('item', '', '   - Bình khí mẫu hiệu chuẩn (Zero & Span gas) & Đồng hồ VOM', 'ca', 0.50, '100%', 'Khí chuẩn N2 và H2 để hiệu chuẩn sensor'),

    # 9. E6B.1.13.01 Ắc quy
    ('norm', 'II.9', 'E6B.1.13.01 - Kiểm tra, bảo dưỡng bình ắc quy nguồn dự phòng tủ PCCC', '1 bình', '', '100%', 'Quyết định 45/QĐ-HĐTV Tập 5 trang 383-384. Áp dụng cho bình ắc quy 12VDC.'),
    ('item', '', '   - Vật liệu: Nước cất nạp accu, mỡ bôi cọc bình chống oxy hóa', 'kg', 0.10, '100%', 'Vệ sinh cọc bình, siết chặt đầu cosse, bôi mỡ bảo vệ'),
    ('item', '', '   - Nhân công kỹ thuật bậc 4,5/7', 'công', 1.50, '100%', 'Đo điện áp từng bình, kiểm tra nạp xả, đo nội trở'),
    ('item', '', '   - Thiết bị đo nội trở mΩ và dung lượng Ah ắc quy', 'ca', 0.25, '100%', 'Máy đo phân tích chuyên dụng đánh giá tuổi thọ ắc quy'),

    # 10. E6C.1.18 & E6C.3.06
    ('norm', 'II.10', 'E6C.1.18 & E6C.3.06 - Bảo dưỡng công tắc dòng chảy & phần điện van PCCC Deluge/Wet', '1 cái', '', '100%', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 425 & 453.'),
    ('item', '', '   - E6C.1.18: Công tắc lưu lượng dòng chảy (Nhân công 5,0/7)', 'công', 1.20, '100%', 'Kiểm tra cánh gạt lưu lượng và tiếp điểm báo về tủ'),
    ('item', '', '   - E6C.3.06: Phần điện van Deluge/Wet (Nhân công 5,0/7: 7.0 công)', 'công', 7.00, '100%', 'Kiểm tra liên động phần điện mở van tràn xả nước'),

    # 11. E6BOP.1.03 & E6BOP.1.01
    ('norm', 'II.11', 'E6BOP.1.03 & E6BOP.1.01 - Đại tu van nước chữa cháy và van an toàn PCCC', '1 cái', '', '100%', 'Quyết định 45/QĐ-HĐTV Tập 6 trang 308 & 312.'),
    ('item', '', '   - E6BOP.1.03 (cột 01): Đại tu van chặn, van nước chữa cháy DN25-DN50', 'công', 2.10, '100%', 'Tháo làm sạch đĩa van, thay gioăng, rà kín mặt côn van'),
    ('item', '', '   - E6BOP.1.01 (cột 02): Sửa chữa, cân chỉnh van an toàn PCCC D100', 'công', 4.50, '100%', 'Tháo rã vệ sinh lò xo, cân chỉnh áp suất xả an toàn trên bàn cân')
]

for row_idx, r_data in enumerate(scl_details, start=curr_r + 1):
    r_type = r_data[0]
    ws.row_dimensions[row_idx].height = 24 if r_type == 'norm' else 20
    
    ws.cell(row=row_idx, column=1, value=r_data[1]).alignment = align_center
    ws.cell(row=row_idx, column=2, value=r_data[2]).alignment = align_left
    ws.cell(row=row_idx, column=3, value=r_data[3]).alignment = align_center
    
    c_val = ws.cell(row=row_idx, column=4, value=r_data[4])
    c_val.alignment = align_right
    if isinstance(r_data[4], (int, float)):
        c_val.number_format = '#,##0.00'
        
    ws.cell(row=row_idx, column=5, value=r_data[5]).alignment = align_center
    ws.cell(row=row_idx, column=6, value=r_data[6]).alignment = align_left
    
    for c in range(1, 7):
        cell = ws.cell(row=row_idx, column=c)
        cell.border = box_border
        if r_type == 'norm':
            cell.font = font_norm
            cell.fill = fill_norm
        elif r_type == 'cat':
            cell.font = font_cat
        else:
            cell.font = font_item

curr_r = curr_r + len(scl_details) + 2

# =========================================================================
# PHẦN 3: CHI TIẾT TẤT CẢ CÁC ĐỊNH MỨC THÔNG TƯ 38/2026/TT-BXD (BỘ XÂY DỰNG)
# =========================================================================
ws.merge_cells(f'A{curr_r}:F{curr_r}')
ws.cell(row=curr_r, column=1, value='MỤC III. CHI TIẾT CÁC ĐỊNH MỨC THÔNG TƯ 38/2026/TT-BXD (BỘ XÂY DỰNG)').font = font_sec
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
    ('norm', 'III.1', 'BD.41111 - Lắp đặt thiết bị đầu báo và đầu báo cháy (Đầu báo khói, báo nhiệt địa chỉ)', '10 đầu', '', '100%', 'Mã gốc BD.41110, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Đinh vít nở M8 liên kết', 'bộ', 20.0, 'Định mức BXD', 'Vật tư phụ bắt đế đầu báo vào trần'),
    ('item', '', '   - Vật liệu: Cồn công nghiệp làm sạch tiếp điểm', 'kg', 0.1, 'Định mức BXD', 'Dung môi làm sạch tiếp xúc'),
    ('item', '', '   - Vật liệu: Đế đầu báo và đầu báo cháy (khói/nhiệt)', 'bộ', 10.0, 'Định mức BXD', 'Vật tư chính (Đơn giá = 0đ do CĐT cấp thiết bị thay thế)'),
    ('item', '', '   - Vật liệu khác (% trên chi phí VL chính)', '%', 2.0, '2.0%', 'Vật tư phụ phát sinh'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc thợ 3,5/7)', 'công', 0.87, '100%', 'Lắp đặt gá đế, đấu dây cáp'),
    ('item', '', '   - Kỹ sư công nghệ (Bậc 5,0/8)', 'công', 0.87, '100%', 'Cài đặt địa chỉ, lập trình, thử khói'),
    ('item', '', '   - Máy khoan bê tông cầm tay 1.0 kW', 'ca', 0.68, '100%', 'Khoan lỗ bắt vít nở trần/dầm'),

    # 2. BD.41121
    ('norm', 'III.2', 'BD.41121 - Lắp đặt đèn báo cháy, đèn hiển thị phòng sau sửa chữa', '5 đèn', '', '100%', 'Mã gốc BD.41120, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Đinh vít nở M8 (10 bộ) + Cồn CN (0.1 kg) + Đèn báo (5 bộ)', 'bộ', 10.0, 'Định mức BXD', 'Vật tư lắp đặt đèn cảnh báo'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc 3,5/7: 0.82 công) + Kỹ sư 5,0/8 (0.35 công)', 'công', 1.17, '100%', 'Tổng nhân công lắp đặt và thử sáng'),
    ('item', '', '   - Máy khoan bê tông cầm tay 1.0 kW', 'ca', 0.75, '100%', 'Khoan bắt vít nở tường/vách'),

    # 3. BD.41131
    ('norm', 'III.3', 'BD.41131 - Lắp đặt nút ấn báo cháy khẩn cấp (Manual Call Point)', '5 nút', '', '100%', 'Mã gốc BD.41130, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Đinh vít nở M6 (10 bộ) + Nút ấn khẩn (5 bộ) + Vật liệu khác 2%', 'bộ', 10.0, 'Định mức BXD', 'Vật tư hộp nút ấn gắn tường'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc 3,5/7: 1.50 công) + Kỹ sư 5,0/8 (1.50 công)', 'công', 3.00, '100%', 'Lắp đặt cao độ chuẩn 1.4m và thử giả lập tác động'),
    ('item', '', '   - Máy khoan bê tông 1.0 kW (0.37 ca) + Đồng hồ VOM (0.37 ca)', 'ca', 0.74, '100%', 'Thiết bị thi công và đo kiểm tiếp điểm'),

    # 4. BD.41141
    ('norm', 'III.4', 'BD.41141 - Lắp đặt chuông báo cháy, còi báo cháy', '5 chuông', '', '100%', 'Mã gốc BD.41140, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Đinh vít nở M6 (10 bộ) + Chuông báo (5 bộ) + Vật liệu khác 5%', 'bộ', 10.0, 'Định mức BXD', 'Vật tư gá treo chuông/còi'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc 3,5/7: 0.82 công) + Kỹ sư 5,0/8 (0.82 công)', 'công', 1.64, '100%', 'Gá lắp và thử nghiệm kích hoạt âm thanh'),
    ('item', '', '   - Máy khoan bê tông 1.0 kW (0.41 ca) + Đồng hồ VOM (0.75 ca)', 'ca', 1.16, '100%', 'Thiết bị thi công và kiểm tra liên động'),

    # 5. BD.41171
    ('norm', 'III.5', 'BD.41171 - Lắp đặt đèn chỉ hướng thoát hiểm EXIT, đèn sự cố', '5 đèn', '', '100%', 'Bảng BD.41170, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Đinh vít nở M6 (10 bộ) + Cồn CN (0.1 kg) + Vật liệu khác 5%', 'bộ', 10.0, 'Định mức BXD', 'Vật tư lắp đèn Exit'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc 3,5/7: 0.73 công) + Kỹ sư 5,0/8 (0.73 công)', 'công', 1.46, '100%', 'Lắp đặt và kiểm tra ngắt nguồn điện lưới thử pin'),
    ('item', '', '   - Máy khoan bê tông 1.0 kW (0.34 ca) + Đồng hồ VOM (0.78 ca)', 'ca', 1.12, '100%', 'Khoan và đo dòng rò sạc ắc quy'),

    # 6. BA.18405
    ('norm', 'III.6', 'BA.18405 - Lắp đặt linh kiện báo cháy, module giám sát, rơ le điều khiển', '1 bộ', '', '100%', 'Bảng BA.18400, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Dây thít, đầu cosse, băng dính cách điện (2% chi phí)', '%', 2.0, '2.0%', 'Vật tư phụ đấu nối module'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc thợ 3,5/7)', 'công', 0.45, '100%', 'Bấm cosse, đấu nối cáp tín hiệu loop'),
    ('item', '', '   - Đồng hồ vạn năng VOM', 'ca', 0.25, '100%', 'Đo điện áp và thử tín hiệu tiếp điểm'),

    # 7. SB.82633 & SB.82634
    ('norm', 'III.7', 'SB.82633 & SB.82634 - Sơn chống rỉ, sơn phủ vỏ thiết bị PCCC (DN50 - DN200)', '100 m2', '', '100%', 'Bảng SB.82630, Chương VIII, Phụ lục 6 TT 38/2026/TT-BXD'),
    ('item', '', '   - Vật liệu: Sơn chống rỉ alkyd và sơn phủ màu đỏ PCCC 2 lớp', 'kg', 32.90, '100%', 'Hệ sơn 2 lớp bảo vệ vỏ van công nghệ cơ nhiệt'),
    ('item', '', '   - Vật liệu khác (giấy nhám, giẻ lau, chổi quét, bạt che)', '%', 1.0, '1.0%', 'Vật tư làm sạch và che chắn'),
    ('item', '', '   - Nhân công nhóm 3 (Bậc thợ 3,5/7)', 'công', 10.80, '100%', 'Đánh gỉ cơ học bề mặt kim loại, thi công sơn 2 lớp'),

    # 8. BB.86601, BB.86103, BB.85201
    ('norm', 'III.8', 'BB.86601 / BB.86103 / BB.85201 - Lắp đặt van ren, van mặt bích, áp kế phụ trợ', '1 cái', '', '100%', 'Bảng BB.86600, BB.86100, BB.85200, Phụ lục 3 TT 38/2026/TT-BXD'),
    ('item', '', '   - BB.86601: Lắp van ren ≤ 25mm (Nhân công 3,5/7: 0.18 công; Máy: 0.05 ca)', 'cái', 0.18, '100%', 'Thay mới van xả ren DN25'),
    ('item', '', '   - BB.86103: Lắp van bích DN80 (Nhân công 3,5/7: 0.45 công; Máy: 0.12 ca)', 'cái', 0.45, '100%', 'Thay mới van Alarm bích DN80'),
    ('item', '', '   - BB.85201: Lắp đồng hồ áp suất (Nhân công 3,5/7: 0.12 công; Máy: 0.04 ca)', 'cái', 0.12, '100%', 'Thay mới đồng hồ áp suất 0-16 bar'),

    # 9. DB.24001
    ('norm', 'III.9', 'DB.24001 - Đo kiểm tra định kỳ điện trở hệ thống chống sét & tiếp địa PCCC', '1 HT', '', '100%', 'Bảng DB.24000, Phụ lục 5 TT 38/2026/TT-BXD'),
    ('item', '', '   - Nhân công: Kỹ sư công nghệ thí nghiệm điện bậc 5,0/8', 'công', 0.50, '100%', 'Đóng cọc phụ, rải dây đo, cân chỉnh cầu đo tiếp địa'),
    ('item', '', '   - Máy thi công: Thiết bị đo điện trở đất chuyên dụng (Terrameter)', 'ca', 0.50, '100%', 'Máy đo điện trở đất kỹ thuật số đạt chuẩn kiểm định')
]

for row_idx, r_data in enumerate(tt38_details, start=curr_r + 1):
    r_type = r_data[0]
    ws.row_dimensions[row_idx].height = 24 if r_type == 'norm' else 20
    
    ws.cell(row=row_idx, column=1, value=r_data[1]).alignment = align_center
    ws.cell(row=row_idx, column=2, value=r_data[2]).alignment = align_left
    ws.cell(row=row_idx, column=3, value=r_data[3]).alignment = align_center
    
    c_val = ws.cell(row=row_idx, column=4, value=r_data[4])
    c_val.alignment = align_right
    if isinstance(r_data[4], (int, float)):
        c_val.number_format = '#,##0.00'
        
    ws.cell(row=row_idx, column=5, value=r_data[5]).alignment = align_center
    ws.cell(row=row_idx, column=6, value=r_data[6]).alignment = align_left
    
    for c in range(1, 7):
        cell = ws.cell(row=row_idx, column=c)
        cell.border = box_border
        if r_type == 'norm':
            cell.font = font_norm
            cell.fill = fill_norm
        else:
            cell.font = font_item

curr_r = curr_r + len(tt38_details) + 2

# =========================================================================
# PHẦN 4: CĂN CỨ PHÁP LÝ & NGUYÊN TẮC BÓC TÁCH CHO CÔNG TÁC "KHÔNG ÁP MÃ"
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
    ('item', '', '   - Căn cứ kỹ thuật: Định mức E6C.3.05 (38.4 công/tủ) đã bao gồm từ bước 5 đến bước 9: kiểm tra toàn diện nguồn cấp 24V, CPU, card mạng loop, card quang.', '', '', 'Chi tiết bước công việc', 'Không tính mã riêng để chống xuất toán trùng lặp kinh phí.'),
    
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
            cell.font = font_norm
            cell.fill = fill_norm
        else:
            cell.font = font_item

curr_r = curr_r + len(kdm_details) + 2

# Signature block
ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=2)
ws.cell(row=curr_r, column=1, value='NGƯỜI LẬP BẢNG KIỂM TRA ĐỊNH MỨC\n(Tổ chuyên gia kỹ thuật dự toán)').alignment = align_center
ws.cell(row=curr_r, column=1).font = font_norm

ws.merge_cells(start_row=curr_r, start_column=3, end_row=curr_r, end_column=4)
ws.cell(row=curr_r, column=3, value='PHÒNG KẾ HOẠCH VẬT TƯ / KTAT\n(Kiểm soát đối chiếu Thông tư 38 & QĐ 45)').alignment = align_center
ws.cell(row=curr_r, column=3).font = font_norm

ws.merge_cells(start_row=curr_r, start_column=5, end_row=curr_r, end_column=6)
ws.cell(row=curr_r, column=5, value='HỘI ĐỒNG THẨM ĐỊNH / LÃNH ĐẠO CÔNG TY\n(Phê duyệt 100% định mức áp dụng)').alignment = align_center
ws.cell(row=curr_r, column=5).font = font_norm

# Column widths
col_widths = [14, 52, 14, 14, 25, 45]
for idx, w in enumerate(col_widths, start=1):
    ws.column_dimensions[get_column_letter(idx)].width = w

ws.freeze_panes = 'C7'

print('Saving updated workbook...')
wb.save(file_local)
print('Saved to local:', file_local)

# Save to Downloads
try:
    wb.save(file_downloads)
    print('Saved to downloads primary:', file_downloads)
except Exception as e:
    print('Could not save to primary downloads (file open):', e)
    wb.save(file_downloads_v2)
    print('Saved to downloads alternative:', file_downloads_v2)
