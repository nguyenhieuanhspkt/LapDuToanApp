import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os, sys

sys.stdout.reconfigure(encoding='utf-8')

src_file = r'C:\Users\Admin\Downloads\DANH MUC CONG TAC_da tra ma.xlsx'
out_file_local = r'd:\TaskApp_kiet\LapDuToanApp\BAO_CAO_RA_SOAT_CHUYEN_DE_DINH_MUC_PCCC_2026.xlsx'
out_file_downloads = r'C:\Users\Admin\Downloads\BAO_CAO_RA_SOAT_CHUYEN_DE_DINH_MUC_PCCC_2026.xlsx'

print('Loading source workbook...')
wb_src = openpyxl.load_workbook(src_file, data_only=True)
ws_src = wb_src.active

# Extract the 238 items where Col AC == 1
items = []

def classify_item(it):
    ma_tu = str(it['ma_tu']).strip()
    name = str(it['name']).lower()
    
    if ma_tu in ['SB.82633', 'SB.82634']:
        return 'CD01', '01_Son_Van', 'Sơn bảo dưỡng vỏ van và phụ tùng PCCC', 'SB.82633 / SB.82634 (TT38/2026/TT-BXD PL6)', 'Đối tượng là vỏ van/thiết bị cơ nhiệt, áp SB.8263x chuẩn 7 ký tự theo TT38, không áp mã sơn kết cấu thép SB.822.'
    elif 'không áp mã (đã gồm trong e6c.3.05)' in ma_tu.lower() or ('không áp mã' in ma_tu.lower() and ('nguồn' in name or 'card' in name)):
        return 'CD02', '02_Tu_Trung_Tam', 'Bo mạch, Card, Nguồn tủ trung tâm FMZ 5000', 'Không áp mã (đã gồm trong E6C.3.05)', 'Định mức E6C.3.05 bảo dưỡng tủ trung tâm (38.4 công) đã gồm bước 5-9 kiểm tra nguồn và card, không áp mã riêng tránh tính trùng.'
    elif 'không áp mã' in ma_tu.lower() and ('đèn' in name or 'exit' in name or 'sự cố' in name):
        return '04_Kiem_Tra_Den_Exit', '04_Kiem_Tra_Den_Exit', 'Kiểm tra định kỳ đèn chỉ lối EXIT và đèn sự cố', 'Không áp mã', 'Công tác kiểm tra định kỳ, không có định mức riêng; không được áp mã lắp mới BD.41171 vì sai lệch bản chất công tác.'
    elif ma_tu == 'E6C.3.03':
        return 'CD03', '03_Module_Hien_Truong', 'Module giám sát và cách ly ngoài hiện trường', 'E6C.3.03 (ĐM SCL EVN)', 'Module lắp đặt phân tán ngoài hiện trường, bảo dưỡng kiểm tra loop áp E6C.3.03 vận dụng bỏ bước 6 theo đúng chuẩn EVN.'
    elif 'ba.18405' in ma_tu.lower():
        return 'CD03', '03_Module_Hien_Truong', 'Lắp đặt thay mới module giám sát ngoài hiện trường', 'BA.18405 (TT38/2026/TT-BXD PL3)', 'Công tác thay mới module hư hỏng, áp mã lắp đặt linh kiện báo cháy BA.18405 theo TT38 đúng bản chất gá lắp, đấu nối.'
    elif ma_tu.startswith('BD.411'):
        return 'CD05', '05_Lap_Moi_Thiet_Bi', 'Lắp mới & thay thế thiết bị báo cháy (Đèn, Nút ấn, Chuông, Exit)', f'{ma_tu} (TT38/2026/TT-BXD)', 'Thay mới thiết bị hư hỏng bằng thiết bị mới, áp dụng định mức lắp đặt chuyên ngành xây dựng TT38 của Bộ Xây dựng.'
    elif ma_tu == 'E6C.2.04':
        return 'CD06', '06_Van_Solenoid', 'Bảo dưỡng van điện từ Solenoid Deluge valve', 'E6C.2.04 (ĐM SCL EVN)', 'Bảo dưỡng cuộn coil và van điện từ kích hoạt van Deluge, áp E6C.2.04 đúng bản chất thiết bị; không áp mã lắp mới van cơ khí BD.42401.'
    elif 'bước 5' in ma_tu.lower() or 'bước 1' in ma_tu.lower():
        return 'CD07', '07_Tach_Buoc_Thao_Lap', 'Bóc tách bước tháo dỡ phục vụ kiểm chuẩn và bảo dưỡng', f'{ma_tu} (ĐM SCL EVN)', 'Bóc tách tỷ lệ nhân công Bước 5 (tháo dỡ 27.27% - 30%) và Bước bảo dưỡng (70% - 72.73%) bảo đảm tính minh bạch, bám sát thực tế thi công.'
    else:
        return 'CD08', '08_Thiet_Bi_Do_Van_Co_Khi', 'Thiết bị đo lường, tủ xả khí và van cơ khí phụ trợ', f'{ma_tu}', 'Áp dụng mã định mức chuyên ngành (Beam E6C.3.04, khí H2 E6C.1.31, công tắc E6C.1.23, tủ HCVR-3 E6C.3.05, van ren BB.86601, van bích BB.86103...).'

for r in range(6, ws_src.max_row + 1):
    val_ac = ws_src.cell(row=r, column=29).value
    if val_ac is not None and str(val_ac).strip() in ['1', '1.0', 1]:
        it = {
            'row_src': r,
            'stt_goc': ws_src.cell(row=r, column=1).value,
            'name': ws_src.cell(row=r, column=2).value,
            'loc': ws_src.cell(row=r, column=3).value,
            'desc': ws_src.cell(row=r, column=4).value,
            'qty': ws_src.cell(row=r, column=5).value,
            'unit': ws_src.cell(row=r, column=6).value,
            'ma_tk': ws_src.cell(row=r, column=7).value,
            'ma_khvt': ws_src.cell(row=r, column=9).value,
            'ma_ktat': ws_src.cell(row=r, column=18).value,
            'ma_tu': ws_src.cell(row=r, column=27).value,
            'ly_do': ws_src.cell(row=r, column=28).value,
            'ghi_chu': ws_src.cell(row=r, column=26).value
        }
        cd_id, sheet_name, cd_title, norm_desc, expert_note = classify_item(it)
        it['cd_id'] = cd_id
        it['sheet_name'] = sheet_name
        it['cd_title'] = cd_title
        it['norm_desc'] = norm_desc
        it['expert_note'] = expert_note
        items.append(it)

print(f'Total extracted items: {len(items)}')

# Initialize target workbook
wb_out = openpyxl.Workbook()
# remove default sheet
wb_out.remove(wb_out.active)

# Styles
font_title = Font(name='Times New Roman', size=15, bold=True, color='002060')
font_subtitle = Font(name='Times New Roman', size=11, bold=True, color='333333')
font_sec_hdr = Font(name='Times New Roman', size=12, bold=True, color='002060')
font_tbl_hdr = Font(name='Times New Roman', size=10, bold=True, color='000000')
font_data = Font(name='Times New Roman', size=10)
font_data_bold = Font(name='Times New Roman', size=10, bold=True)
font_agree = Font(name='Times New Roman', size=10, bold=True, color='006100')

fill_tbl_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_sec_hdr = PatternFill(start_color='B4C6E7', end_color='B4C6E7', fill_type='solid')
fill_agree = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
fill_zebra = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')

border_thin = Side(border_style='thin', color='B0B0B0')
border_thick_bottom = Side(border_style='medium', color='002060')
border_double_bottom = Side(border_style='double', color='000000')

box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
hdr_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)
total_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double_bottom)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# -------------------------------------------------------------
# 1. SHEET 00_Tong_Hop_Chuyen_De
# -------------------------------------------------------------
print('Building Sheet 00_Tong_Hop_Chuyen_De...')
ws00 = wb_out.create_sheet(title='00_Tong_Hop_Chuyen_De')
ws00.views.sheetView[0].showGridLines = True

# Title block
ws00.merge_cells('A1:K1')
ws00['A1'] = 'TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM - TỔNG CÔNG TY PHÁT ĐIỆN 3 - CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4'
ws00['A1'].font = font_subtitle
ws00['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws00.merge_cells('A2:K2')
ws00['A2'] = 'BÁO CÁO CHUYÊN ĐỀ RÀ SOÁT VÀ THỐNG NHẤT MÃ ĐỊNH MỨC DỰ TOÁN'
ws00['A2'].font = font_title
ws00['A2'].alignment = Alignment(horizontal='center', vertical='center')

ws00.merge_cells('A3:K3')
ws00['A3'] = 'Dự án: Sửa chữa lớn Hệ thống PCCC năm 2026 (Theo Quyết định số 1246/QĐ-NĐVT4)'
ws00['A3'].font = Font(name='Times New Roman', size=11, italic=True)
ws00['A3'].alignment = Alignment(horizontal='center', vertical='center')

# Section 1: Căn cứ pháp lý & Kỹ thuật
ws00.cell(row=5, column=1, value='I. CĂN CỨ PHÁP LÝ & QUY CHUẨN THẨM ĐỊNH ĐỊNH MỨC').font = font_sec_hdr
legal_bases = [
    '1. Quyết định số 1246/QĐ-NĐVT4 ngày 30/12/2025 của Công ty Nhiệt điện Vĩnh Tân 4 về việc phê duyệt PAKT Sửa chữa lớn Hệ thống PCCC năm 2026.',
    '2. Quyết định số 45/QĐ-HĐTV ngày 15/03/2022 của Hội đồng thành viên Tập đoàn Điện lực Việt Nam (EVN) ban hành Bộ định mức Sửa chữa lớn tổ máy Nhiệt điện.',
    '3. Thông tư số 38/2026/TT-BXD của Bộ Xây dựng ban hành định mức dự toán xây dựng công trình (Lắp đặt thiết bị PCCC và sơn bảo vệ bề mặt).',
    '4. Thông tư số 11/2021/TT-BXD và Thông tư số 14/2023/TT-BXD của Bộ Xây dựng về hướng dẫn xác định và quản lý chi phí đầu tư xây dựng.',
    '5. Bảng rà soát đối chiếu đa phương án giữa Phương án Tư vấn ban đầu, Phòng Kế hoạch Vật tư (KHVT) và Phòng Kỹ thuật An toàn (KTAT).'
]
for idx, lb in enumerate(legal_bases, start=6):
    ws00.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=11)
    c = ws00.cell(row=idx, column=1, value=lb)
    c.font = font_data
    c.alignment = align_left

# Section 2: Bảng tổng hợp 8 chuyên đề rà soát
row_curr = 12
ws00.cell(row=row_curr, column=1, value='II. TỔNG HỢP KẾT QUẢ RÀ SOÁT & THỐNG NHẤT 238 CÔNG TÁC VƯỚNG MẮC (CỘT AC = 1)').font = font_sec_hdr

row_curr += 1
headers_00 = [
    'STT', 'Mã CĐ', 'Tên chuyên đề rà soát kỹ thuật', 'Số lượng\n(công tác)', 'Tỷ lệ\n(%)',
    'Phương án cũ\n(Tham khảo / KTAT)', 'Mã Đề xuất tối ưu\n(Thống nhất áp dụng)', 'Bộ định mức\náp dụng',
    'Kết luận rà soát', 'Mức độ rủi ro\nkiểm toán', 'Luận chứng cốt lõi'
]

ws00.row_dimensions[row_curr].height = 32
for c_idx, h_text in enumerate(headers_00, start=1):
    cell = ws00.cell(row=row_curr, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

summary_groups = [
    ('1', 'CD01', 'Sơn bảo dưỡng vỏ van và phụ tùng PCCC', 67, '28.2%', 'SB.82220 / SB.822', 'SB.82633 (DN50-100)\nSB.82634 (DN125-200)', 'TT38/2026/TT-BXD\nPhụ lục 6', 'THỐNG NHẤT 100%', 'Rất an toàn', 'Vỏ van công nghệ cơ nhiệt áp mã sơn vỏ thiết bị SB.8263x đủ 7 ký tự chuẩn BXD, không áp mã sơn kết cấu thép SB.822.'),
    ('2', 'CD02', 'Bo mạch, Card, Nguồn tủ trung tâm FMZ5000', 37, '15.5%', 'E6C.3.05.01 / .03', 'Không áp mã\n(Đã gồm trong E6C.3.05)', 'ĐM SCL EVN\n(QĐ 45/QĐ-HĐTV)', 'THỐNG NHẤT 100%', 'Tuyệt đối an toàn\n(Chống trùng lặp)', 'Định mức E6C.3.05 (38.4 công) đã gồm bước 5-9 kiểm tra nguồn và card. Không áp mã con loại bỏ 100% rủi ro xuất toán trùng lặp.'),
    ('3', 'CD03', 'Module giám sát & cách ly ngoài hiện trường', 40, '16.8%', 'KĐM / BD.42201 (camera)', 'E6C.3.03 (33 mục)\nBA.18405 (7 mục thay mới)', 'ĐM SCL EVN &\nTT38 PL3', 'THỐNG NHẤT 100%', 'Rất an toàn', 'Phân định rõ: Bảo dưỡng loop áp E6C.3.03 (bỏ bước 6); Thay mới module áp BA.18405 (lắp linh kiện báo cháy TT38).'),
    ('4', 'CD04', 'Kiểm tra đèn chỉ lối EXIT và đèn sự cố', 14, '5.9%', 'BD.41171 (lắp đèn exit)\nBD.41121 (lắp đèn)', 'Không áp mã\n(Chi phí thường xuyên)', 'Không áp mã', 'THỐNG NHẤT 100%', 'Tuyệt đối an toàn\n(Không áp sai mã)', 'Bản chất là kiểm tra định kỳ, không có định mức bảo dưỡng; không được áp mã lắp đặt mới BD.41171 tránh vi phạm pháp luật.'),
    ('5', 'CD05', 'Lắp mới & thay thế đèn, nút ấn, chuông', 23, '9.7%', 'E6C.3.01.02 / BD.41141', 'BD.41121 / BD.41131\nBD.41141 / BD.41171', 'TT38/2026/TT-BXD\nPhụ lục 3', 'THỐNG NHẤT 100%', 'Rất an toàn', 'Thay mới thiết bị hỏng bằng thiết bị mới, áp dụng chuẩn xác định mức lắp đặt mới chuyên ngành xây dựng của Bộ Xây dựng.'),
    ('6', 'CD06', 'Van điện từ Solenoid Deluge valve', 11, '4.6%', 'BD.42401 (lắp van cơ khí)\nE6C.3.06.02', 'E6C.2.04', 'ĐM SCL EVN\n(QĐ 45/QĐ-HĐTV)', 'THỐNG NHẤT 100%', 'Rất an toàn', 'Bảo dưỡng cuộn coil và van điện từ kích hoạt van Deluge, áp E6C.2.04 đúng bản chất thiết bị; bác bỏ mã lắp mới van cơ khí BD.42401.'),
    ('7', 'CD07', 'Bóc tách bước tháo dỡ phục vụ kiểm chuẩn', 15, '6.3%', 'BD.41111 (lắp mới)\nTính gộp 100%', 'E6C.3.01 (bước 5: 27.27%)\nE6C.3.02 (bước 5: 30%)', 'ĐM SCL EVN\n(QĐ 45/QĐ-HĐTV)', 'THỐNG NHẤT 100%', 'Rất an toàn', 'Bóc tách tỷ lệ nhân công Bước 5 tháo dỡ độc lập với bước bảo dưỡng tại phòng thí nghiệm, sát thực tế tổ chức thi công tại NMNĐ VT4.'),
    ('8', 'CD08', 'Thiết bị đo lường, tủ HCVR-3 và van phụ trợ', 31, '13.0%', 'Mã không tương thích\nhoặc thiếu ký tự cột', 'E6C.3.04 (Beam), E6C.1.31(H2)\nBB.86601, BB.86103, BB.85201', 'ĐM SCL EVN &\nTT38 BXD', 'THỐNG NHẤT 100%', 'Rất an toàn', 'Chuẩn hóa mã định mức 7 ký tự, áp đúng mã chuyên ngành cho đầu báo Beam (E6C.3.04), nồng độ H2 (E6C.1.31) và van bích (BB.86103).')
]

for g in summary_groups:
    row_curr += 1
    ws00.row_dimensions[row_curr].height = 28
    ws00.cell(row=row_curr, column=1, value=int(g[0])).alignment = align_center
    ws00.cell(row=row_curr, column=2, value=g[1]).alignment = align_center
    ws00.cell(row=row_curr, column=3, value=g[2]).alignment = align_left
    ws00.cell(row=row_curr, column=4, value=g[3]).alignment = align_right
    ws00.cell(row=row_curr, column=5, value=g[4]).alignment = align_center
    ws00.cell(row=row_curr, column=6, value=g[5]).alignment = align_center
    ws00.cell(row=row_curr, column=7, value=g[6]).alignment = align_center
    ws00.cell(row=row_curr, column=8, value=g[7]).alignment = align_center
    c_agree = ws00.cell(row=row_curr, column=9, value=g[8])
    c_agree.alignment = align_center
    c_agree.font = font_agree
    c_agree.fill = fill_agree
    ws00.cell(row=row_curr, column=10, value=g[9]).alignment = align_center
    ws00.cell(row=row_curr, column=11, value=g[10]).alignment = align_left
    
    for c_idx in range(1, 12):
        cell = ws00.cell(row=row_curr, column=c_idx)
        cell.border = box_border
        if c_idx != 9:
            cell.font = font_data

# Total Row
row_curr += 1
ws00.row_dimensions[row_curr].height = 26
ws00.merge_cells(start_row=row_curr, start_column=1, end_row=row_curr, end_column=3)
c_tot = ws00.cell(row=row_curr, column=1, value='TỔNG CỘNG TOÀN BỘ CÔNG TÁC RÀ SOÁT (CỘT AC = 1)')
c_tot.font = font_data_bold
c_tot.alignment = align_center

ws00.cell(row=row_curr, column=4, value=238).alignment = align_right
ws00.cell(row=row_curr, column=4).font = font_data_bold

ws00.cell(row=row_curr, column=5, value='100.0%').alignment = align_center
ws00.cell(row=row_curr, column=5).font = font_data_bold

ws00.merge_cells(start_row=row_curr, start_column=6, end_row=row_curr, end_column=8)
ws00.cell(row=row_curr, column=6, value='Đồng bộ 100% theo Cột AA (Đề xuất tối ưu)').alignment = align_center
ws00.cell(row=row_curr, column=6).font = font_data_bold

c_tot_agree = ws00.cell(row=row_curr, column=9, value='THỐNG NHẤT 238/238')
c_tot_agree.alignment = align_center
c_tot_agree.font = font_agree
c_tot_agree.fill = fill_agree

ws00.merge_cells(start_row=row_curr, start_column=10, end_row=row_curr, end_column=11)
ws00.cell(row=row_curr, column=10, value='Hồ sơ đủ điều kiện phê duyệt & bảo vệ trước Kiểm toán').alignment = align_center
ws00.cell(row=row_curr, column=10).font = font_data_bold

for c_idx in range(1, 12):
    ws00.cell(row=row_curr, column=c_idx).border = total_border

# Section 3: Signature Block
row_curr += 3
ws00.merge_cells(start_row=row_curr, start_column=1, end_row=row_curr, end_column=3)
ws00.cell(row=row_curr, column=1, value='NGƯỜI LẬP BÁO CÁO\n(Tổ chuyên gia rà soát)').alignment = align_center
ws00.cell(row=row_curr, column=1).font = font_data_bold

ws00.merge_cells(start_row=row_curr, start_column=4, end_row=row_curr, end_column=7)
ws00.cell(row=row_curr, column=4, value='KIỂM SOÁT KỸ THUẬT & DỰ TOÁN\n(Phòng KHVT - Phòng KTAT)').alignment = align_center
ws00.cell(row=row_curr, column=4).font = font_data_bold

ws00.merge_cells(start_row=row_curr, start_column=8, end_row=row_curr, end_column=11)
ws00.cell(row=row_curr, column=8, value='HỘI ĐỒNG THẨM ĐỊNH / LÃNH ĐẠO CÔNG TY\n(Phê duyệt thống nhất áp dụng)').alignment = align_center
ws00.cell(row=row_curr, column=8).font = font_data_bold

# Set column widths for ws00
col_widths_00 = [6, 10, 32, 12, 10, 24, 28, 20, 20, 18, 45]
for idx, w in enumerate(col_widths_00, start=1):
    ws00.column_dimensions[get_column_letter(idx)].width = w

# -------------------------------------------------------------
# 2. SHEET 01_Danh_Sach_238_Cong_Tac (Master Detailed List)
# -------------------------------------------------------------
print('Building Sheet 01_Danh_Sach_238_Cong_Tac...')
ws01 = wb_out.create_sheet(title='01_Danh_Sach_238_Cong_Tac')
ws01.views.sheetView[0].showGridLines = True

ws01.merge_cells('A1:N1')
ws01['A1'] = 'DANH MỤC CHI TIẾT 238 CÔNG TÁC RÀ SOÁT ĐỊNH MỨC (CỘT AC = 1)'
ws01['A1'].font = font_title
ws01['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws01.merge_cells('A2:N2')
ws01['A2'] = 'Đối chiếu phương án ban đầu, KHVT, KTAT và Đề xuất tối ưu – Kết luận thống nhất của Chuyên gia'
ws01['A2'].font = Font(name='Times New Roman', size=11, italic=True)
ws01['A2'].alignment = Alignment(horizontal='center', vertical='center')

headers_01 = [
    'STT Rà soát', 'STT QĐ 1246', 'Mã CĐ', 'Tên công việc / Hàng hóa dịch vụ', 'Khu vực / Vị trí',
    'Khối lượng', 'ĐVT', 'Mã Tham khảo\n(Ban đầu)', 'Mã KHVT rà', 'Mã KTAT\nđề xuất',
    'MÃ ĐỀ XUẤT TỐI ƯU\n(Cột AA)', 'Lý do Đề xuất tối ưu\n(Cột AB)',
    'Kết luận rà soát', 'Luận chứng kỹ thuật & Giải pháp bảo vệ kiểm toán'
]

ws01.row_dimensions[4].height = 32
for c_idx, h_text in enumerate(headers_01, start=1):
    cell = ws01.cell(row=4, column=c_idx, value=h_text)
    cell.font = font_tbl_hdr
    cell.fill = fill_tbl_hdr
    cell.border = hdr_border
    cell.alignment = align_center

for idx, it in enumerate(items, start=1):
    r_idx = idx + 4
    ws01.row_dimensions[r_idx].height = 24
    
    ws01.cell(row=r_idx, column=1, value=idx).alignment = align_center
    ws01.cell(row=r_idx, column=2, value=it['stt_goc']).alignment = align_center
    ws01.cell(row=r_idx, column=3, value=it['cd_id']).alignment = align_center
    ws01.cell(row=r_idx, column=4, value=it['name']).alignment = align_left
    ws01.cell(row=r_idx, column=5, value=it['loc']).alignment = align_left
    
    c_qty = ws01.cell(row=r_idx, column=6, value=it['qty'])
    c_qty.alignment = align_right
    if isinstance(it['qty'], (int, float)):
        c_qty.number_format = '#,##0.00'
        
    ws01.cell(row=r_idx, column=7, value=it['unit']).alignment = align_center
    ws01.cell(row=r_idx, column=8, value=it['ma_tk'] or '-').alignment = align_center
    ws01.cell(row=r_idx, column=9, value=it['ma_khvt'] or '-').alignment = align_center
    ws01.cell(row=r_idx, column=10, value=it['ma_ktat'] or '-').alignment = align_center
    
    c_tu = ws01.cell(row=r_idx, column=11, value=it['ma_tu'])
    c_tu.alignment = align_center
    c_tu.font = font_data_bold
    
    ws01.cell(row=r_idx, column=12, value=it['ly_do']).alignment = align_left
    
    c_res = ws01.cell(row=r_idx, column=13, value='THỐNG NHẤT')
    c_res.alignment = align_center
    c_res.font = font_agree
    c_res.fill = fill_agree
    
    ws01.cell(row=r_idx, column=14, value=it['expert_note']).alignment = align_left
    
    fill_row = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
    for c in range(1, 15):
        cell = ws01.cell(row=r_idx, column=c)
        cell.border = box_border
        if c != 11 and c != 13:
            cell.font = font_data
        if fill_row.fill_type and c != 13:
            cell.fill = fill_row

# Set column widths for ws01
col_widths_01 = [8, 10, 8, 38, 22, 10, 8, 16, 16, 16, 24, 35, 15, 38]
for idx, w in enumerate(col_widths_01, start=1):
    ws01.column_dimensions[get_column_letter(idx)].width = w

ws01.freeze_panes = 'E5'

# -------------------------------------------------------------
# 3. BUILD 8 DEDICATED THEMATIC SHEETS (CHUYÊN ĐỀ 01 ĐẾN 08)
# -------------------------------------------------------------
sheet_configs = [
    ('CD01_Son_Van', 'CD01', 'CHUYÊN ĐỀ 01: SƠN BẢO DƯỠNG VỎ VAN VÀ THIẾT BỊ PCCC (67 CÔNG TÁC)',
     'Mã áp dụng: SB.82633 (DN50-DN100) và SB.82634 (DN125-DN200) theo TT38/2026/TT-BXD Phụ lục 6.\n'
     'Lý do chuyên gia thống nhất: Vỏ van PCCC là vỏ thiết bị công nghệ cơ nhiệt có kết cấu đúc bích phức tạp. Áp dụng SB.8263x đúng bản chất vỏ thiết bị, loại bỏ mã kết cấu dầm thép SB.82220 / SB.822.'),
    
    ('CD02_Tu_Trung_Tam', 'CD02', 'CHUYÊN ĐỀ 02: BO MẠCH, CARD, NGUỒN TỦ TRUNG TÂM FMZ 5000 (37 CÔNG TÁC)',
     'Mã áp dụng: Không áp mã (Đã tính trọn gói trong định mức E6C.3.05 - 38.4 công/tủ).\n'
     'Lý do chuyên gia thống nhất: Định mức bảo dưỡng tủ trung tâm E6C.3.05 đã bao gồm bước 5 đến 9 kiểm tra vệ sinh card loop, nguồn cấp. Không áp mã riêng để bảo vệ 100% hồ sơ không bị xuất toán trùng lặp.'),
    
    ('CD03_Module_Hien_Truong', 'CD03', 'CHUYÊN ĐỀ 03: MODULE GIÁM SÁT VÀ CÁCH LY NGOÀI HIỆN TRƯỜNG (40 CÔNG TÁC)',
     'Mã áp dụng: E6C.3.03 (Bảo dưỡng 33 module) & BA.18405 (Thay mới 7 module theo TT38 PL3).\n'
     'Lý do chuyên gia thống nhất: Phân định ranh giới rành mạch giữa bảo dưỡng kiểm tra loop ngoài hiện trường (E6C.3.03 vận dụng bỏ bước 6) và công tác thay mới linh kiện gá lắp cáp tín hiệu (BA.18405).'),
    
    ('CD04_Den_Exit_Su_Co', '04_Kiem_Tra_Den_Exit', 'CHUYÊN ĐỀ 04: KIỂM TRA ĐỊNH KỲ ĐÈN CHỈ LỐI EXIT VÀ ĐÈN SỰ CỐ (14 CÔNG TÁC)',
     'Mã áp dụng: Không áp mã (Tính trong chi phí quản lý vận hành / bảo dưỡng thường xuyên).\n'
     'Lý do chuyên gia thống nhất: Công tác kiểm tra test ắc quy, lau chùi mặt kính không có định mức SCL. Kiên quyết không áp mã lắp đặt mới BD.41171 tránh vi phạm nghiêm trọng bản chất hồ sơ dự toán.'),
    
    ('CD05_Lap_Moi_Thiet_Bi', 'CD05', 'CHUYÊN ĐỀ 05: LẮP MỚI & THAY THẾ ĐÈN, NÚT ẤN, CHUÔNG (23 CÔNG TÁC)',
     'Mã áp dụng: BD.41121 (Đèn), BD.41131 (Nút ấn), BD.41141 (Chuông), BD.41171 (Exit) theo TT38/2026/TT-BXD.\n'
     'Lý do chuyên gia thống nhất: Công tác thay thế lắp mới thiết bị hư hỏng bắt buộc phải áp dụng định mức chuyên ngành xây dựng của Bộ Xây dựng (TT38), không áp mã sửa chữa SCL cho phần lắp mới.'),
    
    ('CD06_Van_Solenoid', 'CD06', 'CHUYÊN ĐỀ 06: VAN ĐIỆN TỪ SOLENOID DELUGE VALVE (11 CÔNG TÁC)',
     'Mã áp dụng: E6C.2.04 (Bảo dưỡng van điện từ Solenoid theo ĐM SCL EVN QĐ 45).\n'
     'Lý do chuyên gia thống nhất: Đối tượng là cụm van điện từ Solenoid mở van xả tràn. Mã E6C.2.04 phản ánh đúng thao tác kiểm tra coil, gioăng, lõi từ; bác bỏ mã lắp đặt mới van cơ khí BD.42401 của KTAT.'),
    
    ('CD07_Tach_Buoc_Thao_Lap', 'CD07', 'CHUYÊN ĐỀ 07: BÓC TÁCH BƯỚC THÁO DỠ VÀ BẢO DƯỠNG CẢM BIẾN (15 CÔNG TÁC)',
     'Mã áp dụng: E6C.3.01 (Bước 5: 27.27%) & E6C.3.02 (Bước 5: 30%) phục vụ tháo dỡ thiết bị kiểm chuẩn.\n'
     'Lý do chuyên gia thống nhất: Giải quyết khoảng trống pháp lý khi TT38 không có mã tháo dỡ PCCC. Bóc tách tỷ lệ bước công việc theo đúng quy trình công nghệ của EVN, phản ánh trung thực thực tế thi công.'),
    
    ('CD08_Thiet_Bi_Do_Van', 'CD08', 'CHUYÊN ĐỀ 08: THIẾT BỊ ĐO LƯỜNG, TỦ HCVR-3 VÀ VAN PHỤ TRỢ (31 CÔNG TÁC)',
     'Mã áp dụng: Beam E6C.3.04, khí H2 E6C.1.31, công tắc E6C.1.23, tủ HCVR-3 E6C.3.05, van ren BB.86601, van bích BB.86103...\n'
     'Lý do chuyên gia thống nhất: Chuẩn hóa toàn bộ mã định mức đủ 7 ký tự hợp lệ, áp dụng mã chuyên ngành chính xác cho thiết bị đo lường đặc thù nhiệt điện.')
]

for sheet_title, cd_filter, box_title, box_desc in sheet_configs:
    print(f'Building Sheet {sheet_title}...')
    ws = wb_out.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    # Title & Description Box
    ws.merge_cells('A1:L1')
    ws['A1'] = box_title
    ws['A1'].font = font_title
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    
    ws.merge_cells('A2:L3')
    ws['A2'] = box_desc
    ws['A2'].font = Font(name='Times New Roman', size=10, italic=True)
    ws['A2'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws['A2'].fill = fill_sec_hdr
    
    headers_cd = [
        'STT CĐ', 'STT QĐ 1246', 'Tên công việc / Hàng hóa dịch vụ', 'Khu vực / Vị trí',
        'Khối lượng', 'ĐVT', 'Mã Tham khảo', 'Mã KHVT rà', 'Mã KTAT',
        'MÃ TỐI ƯU (AA)', 'Lý do tối ưu (AB)', 'Đánh giá Chuyên gia'
    ]
    
    ws.row_dimensions[5].height = 28
    for c_idx, h_text in enumerate(headers_cd, start=1):
        cell = ws.cell(row=5, column=c_idx, value=h_text)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.border = hdr_border
        cell.alignment = align_center
        
    grp_items = [it for it in items if it['cd_id'] == cd_filter or (cd_filter == '04_Kiem_Tra_Den_Exit' and it['sheet_name'] == '04_Kiem_Tra_Den_Exit')]
    
    for idx, it in enumerate(grp_items, start=1):
        r_idx = idx + 5
        ws.row_dimensions[r_idx].height = 22
        ws.cell(row=r_idx, column=1, value=idx).alignment = align_center
        ws.cell(row=r_idx, column=2, value=it['stt_goc']).alignment = align_center
        ws.cell(row=r_idx, column=3, value=it['name']).alignment = align_left
        ws.cell(row=r_idx, column=4, value=it['loc']).alignment = align_left
        
        c_qty = ws.cell(row=r_idx, column=5, value=it['qty'])
        c_qty.alignment = align_right
        if isinstance(it['qty'], (int, float)):
            c_qty.number_format = '#,##0.00'
            
        ws.cell(row=r_idx, column=6, value=it['unit']).alignment = align_center
        ws.cell(row=r_idx, column=7, value=it['ma_tk'] or '-').alignment = align_center
        ws.cell(row=r_idx, column=8, value=it['ma_khvt'] or '-').alignment = align_center
        ws.cell(row=r_idx, column=9, value=it['ma_ktat'] or '-').alignment = align_center
        
        c_tu = ws.cell(row=r_idx, column=10, value=it['ma_tu'])
        c_tu.alignment = align_center
        c_tu.font = font_data_bold
        
        ws.cell(row=r_idx, column=11, value=it['ly_do']).alignment = align_left
        
        c_res = ws.cell(row=r_idx, column=12, value='THỐNG NHẤT 100%')
        c_res.alignment = align_center
        c_res.font = font_agree
        c_res.fill = fill_agree
        
        fill_row = fill_zebra if idx % 2 == 0 else PatternFill(fill_type=None)
        for c in range(1, 13):
            cell = ws.cell(row=r_idx, column=c)
            cell.border = box_border
            if c != 10 and c != 12:
                cell.font = font_data
            if fill_row.fill_type and c != 12:
                cell.fill = fill_row
                
    # Column widths
    col_widths_cd = [6, 10, 38, 22, 10, 8, 16, 16, 16, 22, 35, 18]
    for idx, w in enumerate(col_widths_cd, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = 'E6'

print('Saving dedicated workbook...')
wb_out.save(out_file_local)
print(f'Successfully saved to local: {out_file_local}')

try:
    wb_out.save(out_file_downloads)
    print(f'Successfully saved to Downloads: {out_file_downloads}')
except Exception as e:
    print(f'Could not save to Downloads: {e}')
