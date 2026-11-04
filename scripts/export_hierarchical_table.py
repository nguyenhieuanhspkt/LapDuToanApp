import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import sys, os

sys.stdout.reconfigure(encoding='utf-8')

src_file = r'templates/DANH_MUC_CONG_TAC_da_tra_ma.xlsx'
out_local = r'd:\TaskApp_kiet\LapDuToanApp\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
out_downloads = r'C:\Users\Admin\Downloads\BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'

print('Loading source workbook...')
wb_src = openpyxl.load_workbook(src_file, data_only=True)
ws_src = wb_src.active

# Parse 856 services
services = []
for r in range(6, 862):
    val_a = ws_src.cell(row=r, column=1).value
    if isinstance(val_a, (int, float)):
        services.append({
            'stt': int(val_a),
            'name': str(ws_src.cell(row=r, column=2).value or ''),
            'loc': str(ws_src.cell(row=r, column=3).value or ''),
            'desc': str(ws_src.cell(row=r, column=4).value or ''),
            'qty': ws_src.cell(row=r, column=5).value,
            'unit': str(ws_src.cell(row=r, column=6).value or ''),
            'ma_tu': str(ws_src.cell(row=r, column=27).value or '').strip(),
            'ly_do': str(ws_src.cell(row=r, column=28).value or '').strip()
        })

print(f'Parsed {len(services)} services.')

wb_out = openpyxl.Workbook()
wb_out.remove(wb_out.active)

# Styles
font_doc_title = Font(name='Times New Roman', size=14, bold=True, color='002060')
font_doc_sub = Font(name='Times New Roman', size=11, italic=True, color='333333')
font_comp = Font(name='Times New Roman', size=10, bold=True, color='002060')
font_tbl_hdr = Font(name='Times New Roman', size=10, bold=True, color='000000')

font_l1 = Font(name='Times New Roman', size=11, bold=True, color='002060')
font_l2 = Font(name='Times New Roman', size=10, bold=True, color='000000')
font_l3 = Font(name='Times New Roman', size=10)
font_total = Font(name='Times New Roman', size=11, bold=True, color='000000')

fill_tbl_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_l1 = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
fill_l2 = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
fill_total = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
fill_zebra = PatternFill(start_color='FAFAFA', end_color='FAFAFA', fill_type='solid')

border_thin = Side(border_style='thin', color='B0B0B0')
border_thick_bottom = Side(border_style='medium', color='002060')
border_double = Side(border_style='double', color='000000')

box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
hdr_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)
l1_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
total_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# =========================================================================
# SHEET 1: 01_Bang_Phan_Cap_Tong_The
# =========================================================================
ws1 = wb_out.create_sheet(title='01_Bang_Phan_Cap_Tong_The')
ws1.views.sheetView[0].showGridLines = True

# Header company
ws1.merge_cells('A1:G1')
ws1['A1'] = 'TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM - TỔNG CÔNG TY PHÁT ĐIỆN 3 - CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4'
ws1['A1'].font = font_comp
ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws1.merge_cells('A2:G2')
ws1['A2'] = 'BẢNG PHÂN TÍCH PHÂN CẤP ĐỊNH MỨC 856 DỊCH VỤ PCCC (HIERARCHICAL LEVEL TABLE)'
ws1['A2'].font = font_doc_title
ws1['A2'].alignment = Alignment(horizontal='center', vertical='center')

ws1.merge_cells('A3:G3')
ws1['A3'] = 'Dự án: SCL Hệ thống PCCC năm 2026 (Theo QĐ số 1246/QĐ-NĐVT4) – Tiếp cận từ Tính chất nhóm -> Đối tượng -> Hành vi công nghệ'
ws1['A3'].font = font_doc_sub
ws1['A3'].alignment = Alignment(horizontal='center', vertical='center')

headers = [
    'STT Phân tầng',
    'Cấp độ phân tầng\n(Level 1: Nhóm tính chất -> Level 2: Đối tượng -> Level 3: Hành vi)',
    'Số lượng\n(Công tác)',
    'Tỷ lệ\n(%)',
    'Dải STT thực tế trong Dự toán\n(Theo QĐ 1246/QĐ-NĐVT4)',
    'Ghi chú kỹ thuật & Tỷ lệ bước công việc',
    'Mã định mức tối ưu cho tính chất\n(Bộ ĐM & Căn cứ pháp lý)'
]

ws1.row_dimensions[5].height = 32
for col_idx, h_text in enumerate(headers, start=1):
    cell = ws1.cell(row=5, column=col_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

# Data rows for the hierarchy
hierarchy_rows = [
    # Level 1 - Nhóm I
    ('level1', 'I', 'BẢO DƯỠNG HỆ THỐNG ĐIỆN - ĐO LƯỜNG - TỰ ĐỘNG HÓA PCCC', 391, 0.4568,
     'STT 1 ÷ 795 (rải rác trên 55 khu vực nhà máy)', 'Bảo dưỡng thiết bị đang vận hành', 'ĐM SCL EVN (QĐ 45/QĐ-HĐTV)'),
    
    ('level2', '1.1', 'Tủ trung tâm báo cháy & Tủ điều khiển xả khí', 34, 0.0397,
     '', 'Trọn bộ tủ Minimax FMZ 5000 / HCVR-3', 'E6C.3.05 (ĐM SCL EVN)'),
    ('level3', '1.1.1', '   • Kiểm tra, bảo dưỡng tủ trung tâm báo cháy trọn bộ', 34, 0.0397,
     'STT 1, 39, 77, 109, 143, 187, 207, 243, 267, 295, 318, 336, 375, 415, 445, 477, 509, 703, 740...',
     '11 bước công nghệ, hao phí 38.4 công/tủ', 'E6C.3.05 (ĐM SCL EVN)'),
    
    ('level2', '1.2', 'Cảm biến & Đầu báo cháy tự động (Khói quang, Nhiệt, Dây nhiệt)', 174, 0.2033,
     '', 'Chu trình 2 pha bảo dưỡng độc lập', 'E6C.3.01 (ĐM SCL EVN)'),
    ('level3', '1.2.1', '   • Pha 1: Tháo dỡ đầu báo trên cao đưa xuống sàn', 85, 0.0993,
     'STT 17, 49, 87, 119, 153, 194, 214, 250, 274, 302, 325, 343, 382, 422, 452, 486, 526, 762...',
     'Bước 5 định mức: Nhân công × 27.27%', 'E6C.3.01 (bước 5)'),
    ('level3', '1.2.2', '   • Pha 2: Vệ sinh buồng khói, đo cách điện, test khói sàn', 85, 0.0993,
     'STT 18, 50, 88, 120, 154, 195, 215, 251, 275, 303, 326, 344, 383, 423, 453, 487, 527, 763...',
     'Bước 1-4, 6-11: Nhân công × 72.73%', 'E6C.3.01 (bước 1–4, 6–11)'),
    ('level3', '1.2.3', '   • Bảo dưỡng tổ hợp đầu báo / chuông đèn tại chỗ', 4, 0.0047,
     'STT 31, 33, 407, 439', 'Bảo dưỡng trọn gói không tháo dỡ', 'E6C.3.01 (ĐM SCL EVN)'),
    
    ('level2', '1.3', 'Đầu báo Beam thu - phát hồng ngoại', 6, 0.0070,
     '', 'Thiết bị bảo vệ không gian trần cao', 'E6C.3.04 / E6C.3.02 (ĐM SCL EVN)'),
    ('level3', '1.3.1', '   • Pha 1: Tháo đầu báo Beam xuống sàn kiểm tra', 3, 0.0035,
     'STT 180, 236, 471', 'Vận dụng bước 5: Nhân công × 30%', 'E6C.3.02 (bước 5) – vận dụng'),
    ('level3', '1.3.2', '   • Pha 2: Kiểm tra, cân chỉnh tia quang học đầu báo Beam', 3, 0.0035,
     'STT 181, 237, 472', 'Định mức chuyên dùng cho đầu báo Beam', 'E6C.3.04 (ĐM SCL EVN)'),
     
    ('level2', '1.4', 'Nút ấn báo cháy khẩn cấp & Công tắc điều khiển', 102, 0.1192,
     '', 'Nút nhấn khẩn SA5900-908, DCP-AMS', 'E6C.3.02 (ĐM SCL EVN)'),
    ('level3', '1.4.1', '   • Pha 1: Tháo dỡ nút nhấn khẩn xuống bảo dưỡng', 49, 0.0572,
     'STT 9, 47, 85, 117, 151, 192, 212, 248, 272, 300, 323, 341, 380, 420, 450, 484, 524, 744...',
     'Bước 5 định mức: Nhân công × 30%', 'E6C.3.02 (bước 5)'),
    ('level3', '1.4.2', '   • Pha 2: Vệ sinh tiếp điểm, đo thông mạch, giả lập tác động', 49, 0.0572,
     'STT 10, 48, 86, 118, 152, 193, 213, 249, 273, 301, 324, 342, 381, 421, 451, 485, 525, 745...',
     'Bước 1-4, 6-11: Nhân công × 70%', 'E6C.3.02 (bước 1–4, 6–11)'),
    ('level3', '1.4.3', '   • Bảo dưỡng nút ấn khẩn cấp trọn gói tại chỗ', 4, 0.0047,
     'STT 28, 400, 401, 437', 'Thực hiện bảo dưỡng tại chỗ', 'E6C.3.02 (ĐM SCL EVN)'),
     
    ('level2', '1.5', 'Module giám sát & Cách ly ngoài hiện trường', 33, 0.0385,
     'STT 24, 29, 62, 67, 100, 105, 134, 137, 178, 203, 235, 260, 288, 313, 332, 356, 370, 396, 435, 465, 476, 735, 772...',
     'Đo loop hiện trường (vận dụng bỏ bước 6)', 'E6C.3.03 (ĐM SCL EVN)'),
     
    ('level2', '1.6', 'Cụm Van điện từ kích hoạt xả tràn (Solenoid Deluge)', 13, 0.0152,
     'STT 36, 74, 106, 141, 205, 264, 292, 316, 334, 358, 398, 438, 795',
     'Kiểm tra cuộn coil điện từ kích van tràn', 'E6C.2.04 (ĐM SCL EVN)'),
     
    ('level2', '1.7', 'Công tắc áp lực & Công tắc dòng chảy', 10, 0.0117,
     'STT 30, 68, 101, 135, 204, 261, 289, 314, 397, 467',
     'Kiểm tra tiếp điểm điện đóng/ngắt tín hiệu', 'E6C.1.23 (ĐM SCL EVN)'),
     
    ('level2', '1.8', 'Đầu đo nồng độ khí cháy Hydro (H2)', 4, 0.0047,
     'STT 32, 70, 103, 405', 'Đo kiểm chuẩn cảm biến H2 phòng ắc quy', 'E6C.1.31 (ĐM SCL EVN)'),
     
    ('level2', '1.9', 'Hệ thống Ắc quy nguồn dự phòng tủ PCCC', 5, 0.0058,
     'STT 506, 513, 521, 625, 648', 'Đo nội trở, kiểm tra dung lượng nạp xả ắc quy', 'E6B.1.13.01 (ĐM SCL EVN)'),
     
    ('level2', '1.10', 'Thiết bị điện - đo lường khác (Lưu lượng, rơle, chống sét)', 10, 0.0117,
     'STT 37, 75, 107, 142, 410, 541, 545, 567...',
     'Kiểm tra công tắc lưu lượng, tiếp địa', 'E6C.1.18 / DB.24001 (ĐM SCL/TT38)'),

    # Level 1 - Nhóm II
    ('level1', 'II', 'LẮP ĐẶT MỚI & THAY THẾ THIẾT BỊ HƯ HỎNG', 195, 0.2278,
     'STT 11 ÷ 764 (rải rác trên 55 khu vực)', 'Thay mới linh kiện/thiết bị bị hư hỏng', 'ĐM BỘ XÂY DỰNG (TT38/2026/TT-BXD)'),
     
    ('level2', '2.1', 'Cảm biến & Đầu báo cháy tự động', 82, 0.0958,
     'STT 11, 20, 52, 59, 90, 97, 122, 129, 156, 163, 196, 202, 216, 222, 252, 258, 276, 282, 304, 310, 327, 345, 384, 424, 454, 488, 528, 746...',
     'Lắp đặt mới đầu báo khói/nhiệt địa chỉ', 'BD.41111 (TT38/2026/TT-BXD)'),
     
    ('level2', '2.2', 'Nút ấn báo cháy khẩn cấp', 48, 0.0561,
     'STT 19, 58, 96, 128, 162, 201, 221, 257, 281, 309, 331, 350, 389, 429, 459, 493, 533, 761...',
     'Lắp đặt mới nút nhấn báo cháy khẩn cấp', 'BD.41131 (TT38/2026/TT-BXD)'),
     
    ('level2', '2.3', 'Thiết bị cảnh báo âm thanh - ánh sáng', 56, 0.0654,
     '', 'Chuông còi đèn báo cháy', 'BD.41141 / BD.41121 (TT38)'),
    ('level3', '2.3.1', '   • Chuông báo cháy, còi báo cháy', 33, 0.0385,
     'STT 23, 61, 99, 133, 167, 225, 285, 335, 354, 393, 433, 463, 497, 537, 749...',
     'Lắp đặt mới chuông báo cháy', 'BD.41141 (TT38/2026/TT-BXD)'),
    ('level3', '2.3.2', '   • Đèn báo cháy, đèn hiển thị phòng', 23, 0.0269,
     'STT 27, 35, 66, 73, 105, 140, 171, 229, 287, 339, 360, 437, 467, 501, 764...',
     'Lắp đặt mới đèn báo cháy, đèn hiển thị', 'BD.41121 (TT38/2026/TT-BXD)'),
     
    ('level2', '2.4', 'Module giám sát ngoại vi hiện trường', 7, 0.0082,
     'STT 63, 261, 289, 357, 397, 436, 466',
     'Lắp đặt thay mới linh kiện/module báo cháy', 'BA.18405 (TT38 Phụ lục 3)'),
     
    ('level2', '2.5', 'Đèn chỉ lối thoát nạn & Chiếu sáng sự cố', 2, 0.0023,
     'STT 651, 657', 'Lắp đặt mới đèn chỉ hướng thoát hiểm EXIT', 'BD.41171 (TT38/2026/TT-BXD)'),

    # Level 1 - Nhóm III
    ('level1', 'III', 'CƠ KHÍ & SƠN BẢO VỆ CHỐNG ĂN MÒN CỤM VAN PCCC', 92, 0.1075,
     'STT 565 ÷ 856 (Trạm bơm, xưởng hóa, cảng, nhiên liệu)', 'Bảo vệ chống rỉ vỏ van và cơ khí đường ống', 'TT38 PHỤ LỤC 6 & ĐM SCL CƠ KHÍ EVN'),
     
    ('level2', '3.1', 'Cụm Van PCCC và phụ tùng đường ống (Sơn bảo vệ)', 67, 0.0783,
     '', 'Sơn chống rỉ, sơn phủ vỏ van (2 lớp)', 'SB.8263x (TT38 Phụ lục 6)'),
    ('level3', '3.1.1', '   • Vỏ van / Y lọc / Phụ tùng kích cỡ DN50 - DN100', 57, 0.0666,
     'STT 777, 778, 779, 781, 782, 784, 786, 788, 794, 800, 802, 804, 807, 810, 813, 816, 819, 822, 825, 829, 832, 836, 840, 848, 851...',
     'Sơn vỏ thiết bị DN ≤ 100mm', 'SB.82633 (TT38 Phụ lục 6)'),
    ('level3', '3.1.2', '   • Vỏ van / Y lọc / Trụ cứu hỏa kích cỡ DN125 - DN200', 10, 0.0117,
     'STT 796, 808, 811, 814, 817, 820, 823, 827, 830, 852',
     'Sơn vỏ thiết bị DN 125 ÷ 200mm', 'SB.82634 (TT38 Phụ lục 6)'),
     
    ('level2', '3.2', 'Cụm Van cơ khí hệ thống nước chữa cháy', 19, 0.0222,
     'STT 565, 568, 571, 574, 577, 580, 583, 792, 793, 846, 853, 854, 855, 856...',
     'Đại tu van chặn, van an toàn, van xả khí', 'E6BOP.1.03 / E6BOP.1.01 (SCL EVN)'),
     
    ('level2', '3.3', 'Thay mới van và phụ tùng cơ khí phụ trợ', 6, 0.0070,
     'STT 785, 792, 793, 794, 828, 844',
     'Lắp mới van ren DN25, van bích DN80, áp kế', 'BB.86601 / BB.86103 / BB.85201 (TT38)'),

    # Level 1 - Nhóm IV
    ('level1', 'IV', 'CÔNG TÁC ĐÃ GỘP TRỌN GÓI HOẶC KHÔNG TÍNH MÃ RIÊNG', 178, 0.2079,
     'STT 2 ÷ 845 (rải rác trên 55 khu vực)', 'Chống trùng lặp định mức & Chi phí TX', 'KHÔNG ÁP MÃ (KDM)'),
     
    ('level2', '4.1', 'Bo mạch, Card, Bộ nguồn trong Tủ trung tâm FMZ 5000', 115, 0.1343,
     '', 'Đã tính trọn gói trong 38.4 công của tủ mẹ', 'Không áp mã (đã gồm trong E6C.3.05)'),
    ('level3', '4.1.1', '   • Bộ nguồn 24VDC cấp nguồn tủ trung tâm', 17, 0.0199,
     'STT 2, 40, 78, 110, 144, 188, 208, 244, 268, 296, 319, 337, 376, 416, 446, 480, 505',
     'Vệ sinh, kiểm tra điện áp đã gồm ở Bước 5-9 tủ', 'KDM (đã gồm trong E6C.3.05)'),
    ('level3', '4.1.2', '   • Card mạng cáp quang kết nối liên tủ', 16, 0.0187,
     'STT 8, 46, 84, 116, 150, 191, 211, 247, 271, 299, 322, 340, 379, 419, 449, 483',
     'Kiểm tra truyền thông quang đã gồm trong tủ', 'KDM (đã gồm trong E6C.3.05)'),
    ('level3', '4.1.3', '   • Module / Card Loop quản lý địa chỉ đầu báo', 17, 0.0199,
     'STT 3, 41, 79, 111, 145, 189, 209, 245, 269, 297, 320, 338, 377, 417, 447, 481, 507',
     'Kiểm tra mạch loop gắn trên rack tủ', 'KDM (đã gồm trong E6C.3.05)'),
    ('level3', '4.1.4', '   • Module nhánh mở rộng đường tín hiệu', 16, 0.0187,
     'STT 4, 42, 80, 112, 146, 190, 210, 246, 270, 298, 321, 339, 378, 418, 448, 479',
     'Bo mạch cắm trong tủ', 'KDM (đã gồm trong E6C.3.05)'),
    ('level3', '4.1.5', '   • Module điều khiển & kích hoạt rơ le', 34, 0.0397,
     'STT 5, 43, 81, 113, 147, 192, 212, 248, 272, 300, 323, 341, 380, 420, 450, 482, 508...',
     'Module rơle điều khiển nằm trong tủ', 'KDM (đã gồm trong E6C.3.05)'),
    ('level3', '4.1.6', '   • Các module kết nối, hiển thị phụ khác trong tủ', 15, 0.0175,
     'STT 6, 7, 44, 45, 82, 83, 114, 115, 148, 149...',
     'Các linh kiện phần cứng nội bộ tủ', 'KDM (đã gồm trong E6C.3.05)'),
     
    ('level2', '4.2', 'Hệ thống đèn Exit thoát hiểm và Đèn sự cố', 16, 0.0187,
     'STT 543, 599, 600, 617, 618, 634, 635, 654, 655, 660, 661, 704, 744, 775, 776...',
     'Kiểm tra bật test ắc quy, lau mặt kính', 'Không áp mã (Chi phí thường xuyên)'),
     
    ('level2', '4.3', 'Tuyến cáp tín hiệu đường Loop và thiết bị ngoại vi', 24, 0.0280,
     'STT 38, 76, 108, 143, 206, 266, 294, 317, 335, 359, 374, 400, 414, 444, 474, 508, 736, 770...',
     'Xử lý chạm đất, rút dây kiểm tra cáp', 'Không áp mã (Thuộc trách nhiệm loop)'),
     
    ('level2', '4.4', 'Công tác phụ trợ và kiểm tra phương tiện tại chỗ', 23, 0.0269,
     'STT 542, 544, 569, 570, 601, 602, 636, 637, 737, 771, 773, 841, 842, 845...',
     'Kiểm tra bình bọt, bình khí, cuộn vòi', 'Không áp mã (Chi phí phụ trợ)')
]

curr_r = 6
for row_type, stt_val, name_val, cnt_val, pct_val, stt_range_val, note_val, norm_val in hierarchy_rows:
    ws1.row_dimensions[curr_r].height = 24 if row_type != 'level1' else 28
    
    c1 = ws1.cell(row=curr_r, column=1, value=stt_val)
    c1.alignment = align_center
    
    c2 = ws1.cell(row=curr_r, column=2, value=name_val)
    c2.alignment = align_left
    
    c3 = ws1.cell(row=curr_r, column=3, value=cnt_val)
    c3.alignment = align_right
    c3.number_format = '#,##0'
    
    c4 = ws1.cell(row=curr_r, column=4, value=pct_val)
    c4.alignment = align_center
    c4.number_format = '0.0%'
    
    c5 = ws1.cell(row=curr_r, column=5, value=stt_range_val)
    c5.alignment = align_left
    
    c6 = ws1.cell(row=curr_r, column=6, value=note_val)
    c6.alignment = align_left
    
    c7 = ws1.cell(row=curr_r, column=7, value=norm_val)
    c7.alignment = align_left
    
    # Styling based on level
    if row_type == 'level1':
        for col in range(1, 8):
            cell = ws1.cell(row=curr_r, column=col)
            cell.font = font_l1
            cell.fill = fill_l1
            cell.border = l1_border
    elif row_type == 'level2':
        for col in range(1, 8):
            cell = ws1.cell(row=curr_r, column=col)
            cell.font = font_l2
            cell.fill = fill_l2
            cell.border = box_border
    else:
        for col in range(1, 8):
            cell = ws1.cell(row=curr_r, column=col)
            cell.font = font_l3
            cell.border = box_border
            
    curr_r += 1

# Total row
ws1.row_dimensions[curr_r].height = 28
c1 = ws1.cell(row=curr_r, column=1, value='TỔNG')
c1.alignment = align_center

c2 = ws1.cell(row=curr_r, column=2, value='TOÀN BỘ CÔNG TÁC PHẦN DỊCH VỤ')
c2.alignment = align_left

c3 = ws1.cell(row=curr_r, column=3, value=856)
c3.alignment = align_right
c3.number_format = '#,##0'

c4 = ws1.cell(row=curr_r, column=4, value=1.0)
c4.alignment = align_center
c4.number_format = '0.0%'

c5 = ws1.cell(row=curr_r, column=5, value='STT 1 ĐẾN STT 856')
c5.alignment = align_center

c6 = ws1.cell(row=curr_r, column=6, value='Khớp 100% cơ cấu danh mục dự toán')
c6.alignment = align_left

c7 = ws1.cell(row=curr_r, column=7, value='TỐI ƯU HÓA TOÀN DIỆN (EVN & TT38)')
c7.alignment = align_left

for col in range(1, 8):
    cell = ws1.cell(row=curr_r, column=col)
    cell.font = font_total
    cell.fill = fill_total
    cell.border = total_border

# Signature block
curr_r += 3
ws1.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=2)
ws1.cell(row=curr_r, column=1, value='NGƯỜI LẬP BÁO CÁO\n(Tổ chuyên gia rà soát)').alignment = align_center
ws1.cell(row=curr_r, column=1).font = font_l2

ws1.merge_cells(start_row=curr_r, start_column=3, end_row=curr_r, end_column=5)
ws1.cell(row=curr_r, column=3, value='PHÒNG KẾ HOẠCH VẬT TƯ - KỸ THUẬT AN TOÀN\n(Kiểm soát kỹ thuật & định mức)').alignment = align_center
ws1.cell(row=curr_r, column=3).font = font_l2

ws1.merge_cells(start_row=curr_r, start_column=6, end_row=curr_r, end_column=7)
ws1.cell(row=curr_r, column=6, value='HỘI ĐỒNG THẨM ĐỊNH / LÃNH ĐẠO CÔNG TY\n(Phê duyệt phương án kỹ thuật)').alignment = align_center
ws1.cell(row=curr_r, column=6).font = font_l2

# Set widths
col_widths1 = [12, 45, 12, 10, 38, 35, 32]
for idx, w in enumerate(col_widths1, start=1):
    ws1.column_dimensions[get_column_letter(idx)].width = w

ws1.freeze_panes = 'C6'

# =========================================================================
# SHEET 2: 02_Chi_Tiet_856_Dich_Vu
# =========================================================================
ws2 = wb_out.create_sheet(title='02_Chi_Tiet_856_Dich_Vu')
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells('A1:J1')
ws2['A1'] = 'DANH MỤC CHI TIẾT 856 DỊCH VỤ PCCC ÁNH XẠ THEO CẤU TRÚC PHÂN TẦNG'
ws2['A1'].font = font_doc_title
ws2['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws2.merge_cells('A2:J2')
ws2['A2'] = 'Bảng tra cứu chi tiết từng công tác: Tên việc, Khu vực, Khối lượng, ĐVT, Cấp độ Level và Mã định mức tối ưu'
ws2['A2'].font = font_doc_sub
ws2['A2'].alignment = Alignment(horizontal='center', vertical='center')

headers2 = [
    'STT Dự toán', 'Khu vực / Vị trí', 'Tên công việc / Hàng hóa dịch vụ', 'Khối lượng', 'ĐVT',
    'Level 1 (Nhóm tính chất)', 'Level 2 (Đối tượng thiết bị)', 'Level 3 (Hành vi công nghệ)',
    'Mã định mức đề xuất tối ưu', 'Lý do & Căn cứ kỹ thuật'
]

ws2.row_dimensions[4].height = 28
for col_idx, h_text in enumerate(headers2, start=1):
    cell = ws2.cell(row=4, column=col_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

# Function to map each service
def map_item_hierarchy(it):
    name = it['name'].lower()
    ma = it['ma_tu']
    
    if 'không áp mã (đã gồm trong e6c.3.05)' in ma.lower():
        l1 = 'IV. ĐÃ GỘP TRỌN GÓI / KHÔNG ÁP MÃ'
        l2 = '4.1. Bo mạch, Card, Nguồn tủ FMZ5000'
        if 'nguồn' in name: l3 = 'Bộ nguồn 24VDC'
        elif 'quang' in name: l3 = 'Card mạng cáp quang'
        elif 'loop' in name: l3 = 'Module / Card Loop'
        elif 'nhánh' in name: l3 = 'Module nhánh'
        elif 'điều khiển' in name: l3 = 'Module điều khiển'
        elif 'rơ le' in name or 'relay' in name: l3 = 'Module rơ le'
        else: l3 = 'Module mở rộng khác'
        return l1, l2, l3
        
    if ma == 'Không áp mã':
        l1 = 'IV. ĐÃ GỘP TRỌN GÓI / KHÔNG ÁP MÃ'
        if any(k in name for k in ['exit', 'sự cố', 'thoát hiểm']):
            l2 = '4.2. Đèn Exit & Sự cố'
            l3 = 'Kiểm tra định kỳ đèn Exit/sự cố'
        elif 'chạm đất' in name or 'cáp tín hiệu' in name:
            l2 = '4.3. Tuyến cáp tín hiệu'
            l3 = 'Tìm lỗi chạm đất, kiểm tra cáp'
        else:
            l2 = '4.4. Công tác phụ trợ'
            l3 = 'Kiểm tra bình bọt/khí/hộp cứu hỏa'
        return l1, l2, l3

    if ma.startswith('SB.') or 'sơn' in name or ma.startswith('E6BOP.') or ma.startswith('BB.'):
        l1 = 'III. CƠ KHÍ & SƠN BẢO VỆ CHỐNG ĂN MÒN'
        if ma == 'SB.82633':
            l2 = '3.1. Sơn vỏ van & phụ tùng'
            l3 = 'Sơn vỏ van/Y lọc DN50-DN100'
        elif ma == 'SB.82634':
            l2 = '3.1. Sơn vỏ van & phụ tùng'
            l3 = 'Sơn vỏ van/Y lọc DN125-DN200'
        elif ma.startswith('E6BOP.'):
            l2 = '3.2. Van cơ khí PCCC'
            l3 = 'Đại tu van nước/an toàn/xả khí'
        else:
            l2 = '3.3. Thay mới van & phụ tùng'
            l3 = f'Lắp thay mới van/áp kế ({ma})'
        return l1, l2, l3

    if ma.startswith('BD.') or ma.startswith('BA.'):
        l1 = 'II. LẮP ĐẶT MỚI & THAY THẾ (TT38)'
        if ma == 'BD.41111':
            l2 = '2.1. Cảm biến báo cháy'
            l3 = 'Lắp mới đầu báo khói/nhiệt'
        elif ma == 'BD.41131':
            l2 = '2.2. Nút ấn báo cháy'
            l3 = 'Lắp mới nút ấn khẩn'
        elif ma == 'BD.41141':
            l2 = '2.3. Cảnh báo âm thanh - ánh sáng'
            l3 = 'Lắp mới chuông báo cháy'
        elif ma == 'BD.41121':
            l2 = '2.3. Cảnh báo âm thanh - ánh sáng'
            l3 = 'Lắp mới đèn báo cháy'
        elif ma == 'BD.41171':
            l2 = '2.5. Đèn Exit & Sự cố'
            l3 = 'Lắp mới đèn Exit'
        elif 'ba.18405' in ma.lower():
            l2 = '2.4. Module hiện trường'
            l3 = 'Lắp mới linh kiện module BA.18405'
        else:
            l2 = '2.6. Lắp mới khác'
            l3 = f'Lắp mới ({ma})'
        return l1, l2, l3

    # Group I: SCL EVN
    l1 = 'I. BẢO DƯỠNG ĐIỆN - ĐO LƯỜNG - TỰ ĐỘNG HÓA'
    if 'e6c.3.05' in ma.lower():
        l2 = '1.1. Tủ trung tâm FMZ5000/HCVR-3'
        l3 = 'Bảo dưỡng tủ trung tâm (Bước 1-11)'
    elif 'e6c.3.01' in ma.lower():
        l2 = '1.2. Cảm biến đầu báo'
        if 'bước 5' in ma.lower(): l3 = 'Pha 1: Tháo đầu báo xuống sàn (Bước 5)'
        elif 'bước 1' in ma.lower(): l3 = 'Pha 2: Vệ sinh test khói sàn (Bước 1-4, 6-11)'
        else: l3 = 'Bảo dưỡng cảm biến trọn gói'
    elif 'e6c.3.02' in ma.lower():
        if 'beam' in name:
            l2 = '1.3. Đầu báo Beam'
            l3 = 'Pha 1: Tháo đầu báo Beam xuống (Bước 5 vận dụng)'
        else:
            l2 = '1.4. Nút ấn khẩn'
            if 'bước 5' in ma.lower(): l3 = 'Pha 1: Tháo nút ấn xuống (Bước 5)'
            elif 'bước 1' in ma.lower(): l3 = 'Pha 2: Vệ sinh đo tiếp điểm (Bước 1-4, 6-11)'
            else: l3 = 'Bảo dưỡng nút ấn trọn gói'
    elif 'e6c.3.04' in ma.lower():
        l2 = '1.3. Đầu báo Beam'
        l3 = 'Pha 2: Cân chỉnh tia beam E6C.3.04'
    elif 'e6c.3.03' in ma.lower():
        l2 = '1.5. Module loop hiện trường'
        l3 = 'Bảo dưỡng module loop E6C.3.03'
    elif 'e6c.2.04' in ma.lower():
        l2 = '1.6. Van điện từ Solenoid'
        l3 = 'Bảo dưỡng coil van solenoid E6C.2.04'
    elif 'e6c.1.23' in ma.lower():
        l2 = '1.7. Công tắc áp lực / dòng chảy'
        l3 = 'Bảo dưỡng tiếp điểm công tắc E6C.1.23'
    elif 'e6c.1.31' in ma.lower():
        l2 = '1.8. Đầu đo khí H2'
        l3 = 'Đo kiểm chuẩn cảm biến H2 E6C.1.31'
    elif 'e6b.1.13' in ma.lower():
        l2 = '1.9. Ắc quy dự phòng'
        l3 = 'Bảo dưỡng bình ắc quy E6B.1.13'
    else:
        l2 = '1.10. Thiết bị điện khác'
        l3 = f'Bảo dưỡng ({ma})'
        
    return l1, l2, l3

for idx, it in enumerate(services, start=1):
    r_idx = idx + 4
    ws2.row_dimensions[r_idx].height = 20
    
    l1, l2, l3 = map_item_hierarchy(it)
    
    ws2.cell(row=r_idx, column=1, value=it['stt']).alignment = align_center
    ws2.cell(row=r_idx, column=2, value=it['loc']).alignment = align_left
    ws2.cell(row=r_idx, column=3, value=it['name']).alignment = align_left
    
    c_q = ws2.cell(row=r_idx, column=4, value=it['qty'])
    c_q.alignment = align_right
    if isinstance(it['qty'], (int, float)):
        c_q.number_format = '#,##0.00'
        
    ws2.cell(row=r_idx, column=5, value=it['unit']).alignment = align_center
    ws2.cell(row=r_idx, column=6, value=l1).alignment = align_left
    ws2.cell(row=r_idx, column=7, value=l2).alignment = align_left
    ws2.cell(row=r_idx, column=8, value=l3).alignment = align_left
    
    c_ma = ws2.cell(row=r_idx, column=9, value=it['ma_tu'])
    c_ma.alignment = align_center
    c_ma.font = font_l2
    
    ws2.cell(row=r_idx, column=10, value=it['ly_do']).alignment = align_left
    
    fill_row = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
    for c in range(1, 11):
        cell = ws2.cell(row=r_idx, column=c)
        cell.border = box_border
        if c != 9:
            cell.font = font_l3
        if fill_row.fill_type:
            cell.fill = fill_row

col_widths2 = [10, 25, 45, 10, 8, 30, 28, 30, 24, 38]
for idx, w in enumerate(col_widths2, start=1):
    ws2.column_dimensions[get_column_letter(idx)].width = w

ws2.freeze_panes = 'D5'

print('Saving workbook...')
wb_out.save(out_local)
print(f'Successfully saved to local: {out_local}')

try:
    wb_out.save(out_downloads)
    print(f'Successfully saved to Downloads: {out_downloads}')
except Exception as e:
    print(f'Could not save to Downloads: {e}')
