import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import sys, os

sys.stdout.reconfigure(encoding='utf-8')

file_local = r'd:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
file_downloads = r'C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'

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

font_part = Font(name='Times New Roman', size=11, bold=True, color='002060')
font_norm = Font(name='Times New Roman', size=10.5, bold=True, color='002060')
font_cat = Font(name='Times New Roman', size=10, bold=True, color='1F4E79')
font_item = Font(name='Times New Roman', size=10, color='000000')
font_note = Font(name='Times New Roman', size=9.5, italic=True, color='555555')

fill_tbl_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_part = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
fill_norm = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
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
ws.merge_cells('A1:E1')
ws['A1'] = 'TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM - TỔNG CÔNG TY PHÁT ĐIỆN 3 - CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4'
ws['A1'].font = font_comp
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('A2:E2')
ws['A2'] = 'BẢNG TỔNG HỢP CÁC ĐỊNH MỨC & HAO PHÍ ÁP DỤNG THEO THÔNG TƯ 38/2026/TT-BXD'
ws['A2'].font = font_doc_title
ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

ws.merge_cells('A3:E3')
ws['A3'] = 'Dự án: SCL Hệ thống PCCC năm 2026 – NMNĐ Vĩnh Tân 4 (Chi tiết các thành phần hao phí Vật liệu, Nhân công, Máy thi công và Căn cứ pháp lý)'
ws['A3'].font = font_doc_sub
ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

headers = [
    'STT',
    'Mã hiệu & Danh mục định mức / Thành phần hao phí',
    'Đơn vị tính\n(ĐVT)',
    'Mức hao phí\nđịnh mức',
    'Cơ sở pháp lý\n(Thông tư số 38/2026/TT-BXD của Bộ Xây dựng)'
]

ws.row_dimensions[5].height = 32
for c_idx, h_text in enumerate(headers, start=1):
    cell = ws.cell(row=5, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

# Data definition
# Types: 'part', 'norm', 'cat', 'item'
rows_data = [
    # -------------------------------------------------------------
    # PHẦN I
    # -------------------------------------------------------------
    ('part', 'I', 'CÁC ĐỊNH MỨC LẮP ĐẶT MỚI & THAY THẾ THIẾT BỊ BÁO CHÁY TỰ ĐỘNG (PHỤ LỤC 3)', '', '', 'Thông tư 38/2026/TT-BXD - Phụ lục 3: Định mức dự toán xây dựng công trình - Lắp đặt thiết bị PCCC'),
    
    # 1. BD.41111
    ('norm', '1', 'BD.41111 - Lắp đặt thiết bị đầu báo và đầu báo cháy (Đầu báo khói, báo nhiệt địa chỉ)', '10 đầu', '', 'Mã gốc BD.41110, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   1.1. Hao phí Vật liệu (Materials)', '', '', 'Quy định chi tiết vật tư phụ và thiết bị chính'),
    ('item', '', '      - Đinh vít nở M8', 'bộ', 20.0, 'Vật tư phụ liên kết gá đỡ đế đầu báo'),
    ('item', '', '      - Cồn công nghiệp', 'kg', 0.1, 'Dung môi làm sạch tiếp điểm điện tử'),
    ('item', '', '      - Đế đầu báo và đầu báo cháy (Đầu báo khói/nhiệt)', 'bộ', 10.0, 'Vật tư thiết bị chính (Lưu ý: Đơn giá VL = 0đ do CĐT cấp thiết bị thay thế)'),
    ('item', '', '      - Vật liệu khác', '%', 2.0, 'Tính bằng 2% trên chi phí vật liệu chính'),
    ('cat', '', '   1.2. Hao phí Nhân công (Labor)', '', '', 'Thành phần nhân lực thi công và hiệu chỉnh'),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.87, 'Thao tác lắp đặt cơ khí, gá đế, đấu dây cáp'),
    ('item', '', '      - Kỹ sư công nghệ (Bậc 5,0/8)', 'công', 0.87, 'Cài đặt địa chỉ, lập trình loop, thử nghiệm khói tại bàn test'),
    ('cat', '', '   1.3. Hao phí Máy thi công (Machines)', '', '', 'Dụng cụ và máy móc thi công chuyên dụng'),
    ('item', '', '      - Máy khoan bê tông cầm tay (Công suất: 1,00 kW)', 'ca', 0.68, 'Khoan lỗ bắt vít nở trên trần/dầm bê tông'),
    
    # 2. BD.41121
    ('norm', '2', 'BD.41121 - Lắp đặt đèn báo cháy, đèn hiển thị phòng sau sửa chữa', '5 đèn', '', 'Mã gốc BD.41120, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   2.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Đinh vít nở M8', 'bộ', 10.0, 'Vật tư phụ liên kết giá đỡ đèn'),
    ('item', '', '      - Cồn công nghiệp', 'kg', 0.1, 'Làm sạch tiếp điểm giắc cắm'),
    ('item', '', '      - Đèn báo cháy, đèn hiển thị', 'bộ', 5.0, 'Thiết bị đèn LED cảnh báo (Giá VL = 0đ khi thay thế)'),
    ('item', '', '      - Vật liệu khác', '%', 2.0, 'Tính bằng 2% trên chi phí vật liệu chính'),
    ('cat', '', '   2.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.82, 'Thao tác gá lắp, đấu nối dây nguồn/tín hiệu'),
    ('item', '', '      - Kỹ sư công nghệ (Bậc 5,0/8)', 'công', 0.35, 'Đo kiểm điện áp kích hoạt, thử nghiệm sáng đèn'),
    ('cat', '', '   2.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Máy khoan bê tông cầm tay (Công suất: 1,00 kW)', 'ca', 0.75, 'Khoan gá lắp đèn trên tường/vách ngăn'),

    # 3. BD.41131
    ('norm', '3', 'BD.41131 - Lắp đặt nút ấn báo cháy khẩn cấp (Manual Call Point)', '5 nút', '', 'Mã gốc BD.41130, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   3.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Đinh vít nở M6', 'bộ', 10.0, 'Vật tư liên kết hộp nút ấn vào tường'),
    ('item', '', '      - Cồn công nghiệp', 'kg', 0.1, 'Làm sạch tiếp điểm cơ khí'),
    ('item', '', '      - Nút báo cháy khẩn cấp', 'bộ', 5.0, 'Thiết bị hộp nút ấn khẩn (Giá VL = 0đ khi thay thế)'),
    ('item', '', '      - Vật liệu khác', '%', 2.0, 'Tính bằng 2% trên chi phí vật liệu chính'),
    ('cat', '', '   3.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 1.50, 'Định vị cao độ (1.4m), gá lắp hộp nút ấn, đấu dây'),
    ('item', '', '      - Kỹ sư công nghệ (Bậc 5,0/8)', 'công', 1.50, 'Đo thông mạch tiếp điểm thường hở/đóng, thử nghiệm giả lập'),
    ('cat', '', '   3.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Máy khoan bê tông cầm tay (Công suất: 1,00 kW)', 'ca', 0.37, 'Khoan bắt vít nở tường gạch/bê tông'),
    ('item', '', '      - Đồng hồ vạn năng (VOM kỹ thuật số)', 'ca', 0.37, 'Đo kiểm tra điện trở tiếp điểm và tín hiệu loop'),

    # 4. BD.41141
    ('norm', '4', 'BD.41141 - Lắp đặt chuông báo cháy, còi báo cháy', '5 chuông', '', 'Mã gốc BD.41140, Bảng BD.41100, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   4.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Đinh vít nở M6', 'bộ', 10.0, 'Vật tư gá treo chuông/còi'),
    ('item', '', '      - Cồn công nghiệp', 'kg', 0.1, 'Vệ sinh cuộn dây nam châm điện'),
    ('item', '', '      - Chuông báo cháy / Còi báo động', 'bộ', 5.0, 'Thiết bị chuông 24VDC (Giá VL = 0đ khi thay thế)'),
    ('item', '', '      - Vật liệu khác', '%', 5.0, 'Tính bằng 5% trên chi phí vật liệu chính'),
    ('cat', '', '   4.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.82, 'Gá lắp chuông, đấu nối đường dây điều khiển'),
    ('item', '', '      - Kỹ sư công nghệ (Bậc 5,0/8)', 'công', 0.82, 'Đo dòng kích hoạt 24VDC, đo cường độ âm thanh (dB)'),
    ('cat', '', '   4.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Máy khoan bê tông cầm tay (Công suất: 1,00 kW)', 'ca', 0.41, 'Khoan bắt vít nở'),
    ('item', '', '      - Đồng hồ vạn năng (VOM kỹ thuật số)', 'ca', 0.75, 'Đo dòng tiêu thụ và kiểm tra liên động'),

    # 5. BD.41171
    ('norm', '5', 'BD.41171 - Lắp đặt đèn chỉ hướng thoát nạn EXIT, đèn chiếu sáng sự cố', '5 đèn', '', 'Mã gốc BD.41170, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   5.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Đinh vít nở M6', 'bộ', 10.0, 'Vật tư gá treo đèn trên cửa/vách'),
    ('item', '', '      - Cồn công nghiệp', 'kg', 0.1, 'Vệ sinh giắc cắm nguồn 220VAC/pin'),
    ('item', '', '      - Vật liệu khác', '%', 5.0, 'Tính bằng 5% trên chi phí vật liệu chính'),
    ('cat', '', '   5.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.73, 'Gá lắp đèn, đấu nối dây nguồn cấp điện'),
    ('item', '', '      - Kỹ sư công nghệ (Bậc 5,0/8)', 'công', 0.73, 'Kiểm tra mạch nạp ắc quy dự phòng, thử ngắt điện lưới'),
    ('cat', '', '   5.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Máy khoan bê tông cầm tay (Công suất: 1,00 kW)', 'ca', 0.34, 'Khoan gá lắp đèn'),
    ('item', '', '      - Đồng hồ vạn năng (VOM kỹ thuật số)', 'ca', 0.78, 'Đo điện áp sạc pin và dòng rò'),

    # 6. BA.18405
    ('norm', '6', 'BA.18405 - Lắp đặt rơ le, công tắc dòng chảy, linh kiện module báo cháy', '1 bộ', '', 'Mã hiệu BA.18405, Bảng BA.18400, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   6.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Vật liệu khác (dây thít, đầu cos, băng cách điện)', '%', 2.0, 'Tính bằng 2% trên chi phí vật liệu chính'),
    ('cat', '', '   6.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.45, 'Gá lắp hộp module, bấm cosse, đấu nối cáp tín hiệu loop'),
    ('cat', '', '   6.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Đồng hồ vạn năng (VOM kỹ thuật số)', 'ca', 0.25, 'Kiểm tra thông mạch tiếp điểm khô, kiểm tra điện áp loop'),

    # -------------------------------------------------------------
    # PHẦN II
    # -------------------------------------------------------------
    ('part', 'II', 'CÁC ĐỊNH MỨC SƠN BẢO VỆ CHỐNG ĂN MÒN VỎ THIẾT BỊ, VAN PCCC (PHỤ LỤC 6)', '', '', 'Thông tư 38/2026/TT-BXD - Phụ lục 6: Định mức dự toán xây dựng công trình - Phần Sơn và Bảo vệ bề mặt'),
    
    # 7. SB.82633
    ('norm', '7', 'SB.82633 - Sơn chống rỉ, sơn phủ – sơn trực tiếp lên vỏ thiết bị trong nhà (DN50 - DN100)', '100 m2', '', 'Bảng SB.82630, Chương VIII, Phụ lục 6 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   7.1. Hao phí Vật liệu (Materials)', '', '', 'Hệ sơn 2 lớp: 1 lớp lót chống rỉ alkyd + 1 lớp phủ màu đỏ PCCC'),
    ('item', '', '      - Sơn chống rỉ và sơn phủ hoàn thiện', 'kg', 32.90, 'Định mức sơn vỏ thiết bị đúc bích phức tạp (DN50 - DN100)'),
    ('item', '', '      - Vật liệu khác (giấy nhám, giẻ lau, chổi quét, bạt che)', '%', 1.0, 'Tính bằng 1% trên chi phí sơn chính'),
    ('cat', '', '   7.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 10.80, 'Cạo rỉ, đánh sạch bề mặt, quét sơn lót và sơn phủ 2 lớp'),
    ('cat', '', '   7.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Không sử dụng máy thi công (Thao tác thủ công chổi/ru lô)', '', '', 'Theo quy định thành phần công việc của định mức BXD'),

    # 8. SB.82634
    ('norm', '8', 'SB.82634 - Sơn chống rỉ, sơn phủ – sơn trực tiếp lên vỏ thiết bị ngoài nhà (DN125 - DN200)', '100 m2', '', 'Bảng SB.82630, Chương VIII, Phụ lục 6 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   8.1. Hao phí Vật liệu (Materials)', '', '', 'Sơn chống rỉ và sơn phủ chịu thời tiết ngoài trời'),
    ('item', '', '      - Sơn chống rỉ và sơn phủ hoàn thiện chịu thời tiết', 'kg', 32.90, 'Định mức sơn van xả tràn, van chặn, trụ cứu hỏa (DN125 - DN200)'),
    ('item', '', '      - Vật liệu khác (giấy nhám, giẻ lau, chổi quét, dung môi)', '%', 1.0, 'Tính bằng 1% trên chi phí sơn chính'),
    ('cat', '', '   8.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 10.80, 'Làm sạch cơ học bề mặt kim loại, thi công sơn 2 lớp'),
    ('cat', '', '   8.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Không sử dụng máy thi công', '', '', 'Thao tác thủ công theo định mức TT38'),

    # -------------------------------------------------------------
    # PHẦN III
    # -------------------------------------------------------------
    ('part', 'III', 'CÁC ĐỊNH MỨC LẮP ĐẶT VAN VÀ THIẾT BỊ PHỤ TRỢ CƠ KHÍ PCCC (PHỤ LỤC 3)', '', '', 'Thông tư 38/2026/TT-BXD - Phụ lục 3: Định mức dự toán xây dựng công trình - Lắp đặt đường ống và phụ tùng'),
    
    # 9. BB.86601
    ('norm', '9', 'BB.86601 - Lắp đặt van ren đường kính ren ≤ 25mm (Thay mới van xả DN25)', '1 cái', '', 'Mã hiệu BB.86601, Bảng BB.86600, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   9.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Vật liệu khác (Băng tan teflon, keo làm kín ren, dầu mỡ)', '%', 2.0, 'Tính bằng 2% chi phí vật liệu chính'),
    ('cat', '', '   9.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.18, 'Quấn băng tan, vặn ren định vị van xả, thử kín áp lực'),
    ('cat', '', '   9.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Máy/dụng cụ cầm tay', 'ca', 0.05, 'Mỏ lết, kìm nước siết ren'),

    # 10. BB.86103
    ('norm', '10', 'BB.86103 - Lắp đặt van mặt bích đường kính DN80 (Thay mới van Alarm DN80)', '1 cái', '', 'Mã hiệu BB.86103, Bảng BB.86100, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   10.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Vật liệu khác (Bulong, đai ốc M16, gioăng đệm cao su chịu áp lực)', '%', 3.0, 'Tính bằng 3% chi phí vật liệu chính'),
    ('cat', '', '   10.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.45, 'Vận chuyển van vào vị trí, cân chỉnh bích, siết chéo bu lông'),
    ('cat', '', '   10.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Máy xiết bu lông / Palang xích nâng hạ', 'ca', 0.12, 'Nâng hạ van và siết lực đều các bu lông mặt bích'),

    # 11. BB.85201
    ('norm', '11', 'BB.85201 - Lắp đặt đồng hồ đo áp suất đường ống chữa cháy (Áp kế 0-16 bar)', '1 cái', '', 'Mã hiệu BB.85201, Bảng BB.85200, Phụ lục 3 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   11.1. Hao phí Vật liệu (Materials)', '', '', ''),
    ('item', '', '      - Vật liệu khác (Băng tan teflon, gioăng đệm đồng chịu áp)', '%', 2.0, 'Tính bằng 2% chi phí vật liệu chính'),
    ('cat', '', '   11.2. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Nhân công nhóm 3 (Bậc thợ bình quân 3,5/7)', 'công', 0.12, 'Lắp đồng hồ vào van khóa 3 ngả, quấn băng tan, siết chặt'),
    ('cat', '', '   11.3. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Dụng cụ thi công cầm tay', 'ca', 0.04, 'Cờ lê dẹt chuyên dụng siết chân ren áp kế'),

    # -------------------------------------------------------------
    # PHẦN IV
    # -------------------------------------------------------------
    ('part', 'IV', 'CÁC ĐỊNH MỨC THÍ NGHIỆM ĐIỆN VÀ ĐO ĐẠC HỆ THỐNG TIẾP ĐỊA (PHỤ LỤC 5)', '', '', 'Thông tư 38/2026/TT-BXD - Phụ lục 5: Định mức dự toán xây dựng công trình - Thí nghiệm điện và đo lường'),
    
    # 12. DB.24001
    ('norm', '12', 'DB.24001 - Đo kiểm tra định kỳ điện trở tiếp địa hệ thống chống sét và tiếp địa PCCC', '1 hệ thống', '', 'Bảng DB.24000, Phụ lục 5 TT 38/2026/TT-BXD (Bộ Xây dựng)'),
    ('cat', '', '   12.1. Hao phí Nhân công (Labor)', '', '', ''),
    ('item', '', '      - Kỹ sư công nghệ thí nghiệm điện (Bậc 5,0/8)', 'công', 0.50, 'Đóng cọc phụ, rải dây đo, cân chỉnh cầu đo, lập biên bản đo điện trở đất'),
    ('cat', '', '   12.2. Hao phí Máy thi công (Machines)', '', '', ''),
    ('item', '', '      - Thiết bị đo điện trở đất chuyên dụng (Terrameter)', 'ca', 0.50, 'Máy đo điện trở đất kỹ thuật số chuyên dụng (đạt chuẩn kiểm định)')
]

curr_r = 6
for r_type, stt_val, name_val, unit_val, rate_val, legal_val in rows_data:
    ws.row_dimensions[curr_r].height = 26 if r_type in ['part', 'norm'] else 20
    
    c1 = ws.cell(row=curr_r, column=1, value=stt_val)
    c1.alignment = align_center
    
    c2 = ws.cell(row=curr_r, column=2, value=name_val)
    c2.alignment = align_left
    
    c3 = ws.cell(row=curr_r, column=3, value=unit_val)
    c3.alignment = align_center
    
    c4 = ws.cell(row=curr_r, column=4, value=rate_val)
    c4.alignment = align_right
    if isinstance(rate_val, (int, float)):
        if unit_val == '%':
            c4.number_format = '0.0"%"'
        elif isinstance(rate_val, int) or rate_val.is_integer():
            c4.number_format = '#,##0.0'
        else:
            c4.number_format = '#,##0.00'
            
    c5 = ws.cell(row=curr_r, column=5, value=legal_val)
    c5.alignment = align_left
    
    # Format according to type
    if r_type == 'part':
        for col in range(1, 6):
            cell = ws.cell(row=curr_r, column=col)
            cell.font = font_part
            cell.fill = fill_part
            cell.border = part_border
    elif r_type == 'norm':
        for col in range(1, 6):
            cell = ws.cell(row=curr_r, column=col)
            cell.font = font_norm
            cell.fill = fill_norm
            cell.border = box_border
    elif r_type == 'cat':
        for col in range(1, 6):
            cell = ws.cell(row=curr_r, column=col)
            cell.font = font_cat
            cell.border = box_border
    else: # item
        for col in range(1, 6):
            cell = ws.cell(row=curr_r, column=col)
            cell.font = font_item
            cell.border = box_border
            
    curr_r += 1

# Signature block
curr_r += 2
ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=2)
ws.cell(row=curr_r, column=1, value='NGƯỜI LẬP BẢNG ĐỊNH MỨC\n(Tổ chuyên gia kỹ thuật dự toán)').alignment = align_center
ws.cell(row=curr_r, column=1).font = font_norm

ws.merge_cells(start_row=curr_r, start_column=3, end_row=curr_r, end_column=4)
ws.cell(row=curr_r, column=3, value='PHÒNG KẾ HOẠCH VẬT TƯ / KTAT\n(Kiểm soát đối chiếu Thông tư 38)').alignment = align_center
ws.cell(row=curr_r, column=3).font = font_norm

ws.cell(row=curr_r, column=5, value='HỘI ĐỒNG THẨM ĐỊNH / LÃNH ĐẠO PHÊ DUYỆT\n(Phê duyệt áp dụng định mức)').alignment = align_center
ws.cell(row=curr_r, column=5).font = font_norm

# Column widths
col_widths = [8, 65, 14, 15, 45]
for idx, w in enumerate(col_widths, start=1):
    ws.column_dimensions[get_column_letter(idx)].width = w

ws.freeze_panes = 'C6'

print('Saving updated workbook...')
wb.save(file_local)
print('Saved to local:', file_local)

try:
    wb.save(file_downloads)
    print('Saved to downloads:', file_downloads)
except Exception as e:
    print('Could not save to downloads:', e)
