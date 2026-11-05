import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb_path = 'd:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
dl_path = 'C:/Users/Admin/Downloads/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'

print("Loading workbook...")
wb = openpyxl.load_workbook(wb_path)

ws_hp = wb['hao phi VT']

# Scan hao phi VT to extract detail items for each resource type
# Only include rows that actually apply (hs > 0)
nc_groups = {}
may_groups = {}
vl_groups = {}

curr_task_stt = None
curr_task_name = None
curr_task_khuvuc = None

for r in range(7, ws_hp.max_row + 1):
    stt = ws_hp.cell(r, 1).value
    code = ws_hp.cell(r, 2).value
    loai = ws_hp.cell(r, 3).value
    khu_vuc = ws_hp.cell(r, 4).value
    name = ws_hp.cell(r, 5).value
    dvt = ws_hp.cell(r, 6).value
    hs = ws_hp.cell(r, 12).value

    if loai == 'CÔNG TÁC':
        curr_task_stt = stt
        curr_task_name = name
        curr_task_khuvuc = khu_vuc
    elif loai == 'NHÂN CÔNG':
        if hs and float(hs) > 0:
            clean_name = str(name).lstrip('- ').strip()
            nc_groups.setdefault(clean_name, []).append({
                'hp_row': r,
                'col_letter': 'N',
                'stt': curr_task_stt,
                'task_name': curr_task_name,
                'khu_vuc': curr_task_khuvuc,
                'code': code,
                'dvt': dvt
            })
    elif loai in ['MÁY THI CÔNG', 'DỤNG CỤ', 'THIẾT BỊ']:
        if hs and float(hs) > 0:
            clean_name = str(name).lstrip('- ').strip()
            col_letter = 'O' if loai == 'DỤNG CỤ' else 'P'
            may_groups.setdefault(clean_name, []).append({
                'hp_row': r,
                'col_letter': col_letter,
                'stt': curr_task_stt,
                'task_name': curr_task_name,
                'khu_vuc': curr_task_khuvuc,
                'code': code,
                'dvt': dvt
            })
    elif loai == 'VẬT LIỆU':
        if hs and float(hs) > 0:
            clean_name = str(name).lstrip('- ').strip()
            vl_groups.setdefault(clean_name, []).append({
                'hp_row': r,
                'col_letter': 'M',
                'stt': curr_task_stt,
                'task_name': curr_task_name,
                'khu_vuc': curr_task_khuvuc,
                'code': code,
                'dvt': dvt
            })

# Prices map: Đơn giá nhân công chuẩn theo Quyết định số 177/QĐ-HĐTV ngày 29/10/2026 của HĐTV EVN
# kết hợp Quyết định số 311/QĐ-SXD ngày 22/12/2023 của Sở Xây dựng Bình Thuận (Vùng IV - NMNĐ Vĩnh Tân 4)
# tính theo hệ số cấp bậc quy định tại Thông tư số 13/2021/TT-BXD của Bộ Xây dựng.
nc_prices = {
    # 1. Nhóm 3 (Bậc 5,0/7) theo QĐ 177/QĐ-HĐTV & QĐ 45/QĐ-HĐTV: 262.300 * 1.94 / 1.52 = 334.778 đ/công
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 theo QĐ 177 & QĐ 45)': 334778,
    'Nhân công tháo dỡ trên cao (Nhóm 3 - Bậc 5,0/7 - E6C.3.01 bước 5)': 334778,
    'Nhân công bảo dưỡng, kiểm chuẩn sàn (Nhóm 3 - Bậc 5,0/7)': 334778,
    'Nhân công tháo dỡ trên cao (Nhóm 3 - Bậc 5,0/7 - E6C.3.02 bước 5)': 334778,
    'Nhân công hiệu chuẩn, thử nghiệm sàn (Nhóm 3 - Bậc 5,0/7)': 334778,
    'Nhân công tháo dỡ beam trên cao (Vận dụng E6C.3.02 bước 5)': 334778,
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Cân chỉnh beam)': 334778,
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Module loop)': 334778,
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Cụm van Deluge)': 334778,
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Đại tu van nước)': 334778,
    'Thợ cơ khí sửa chữa bậc 4,0/7 (Đại tu van nước PCCC)': 334778,

    # 2. Nhóm 4 (Bậc 5,0/7) theo QĐ 177/QĐ-HĐTV & QĐ 45/QĐ-HĐTV: 272.100 * 1.94 / 1.52 = 347.286 đ/công
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Van Solenoid)': 347286,
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Công tắc áp lực)': 347286,
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Đo lưu lượng)': 347286,
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Hiệu chuẩn khí H2)': 347286,
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Kiểm tra ắc quy)': 347286,
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Van an toàn)': 347286,
    'Thợ cơ khí sửa chữa bậc 4,5/7 (Đại tu van an toàn)': 347286,

    # 3. Nhân công theo Thông tư 38/2026/TT-BXD & TT 13/2021/TT-BXD (Bình Thuận Vùng IV):
    'Thợ lắp đặt thiết bị PCCC bậc 3,5/7': 262300, # Nhóm 3 bậc 3,5/7: 262.300
    'Thợ lắp đặt đường ống PCCC bậc 3,5/7': 262300, # Nhóm 3 bậc 3,5/7: 262.300
    'Thợ lắp đặt bậc 3,5/7 (Lắp van ren ≤ DN25)': 262300, # Nhóm 3 bậc 3,5/7: 262.300
    'Thợ lắp đặt bậc 3,5/7 (Lắp van bích DN80)': 262300, # Nhóm 3 bậc 3,5/7: 262.300
    'Thợ lắp đặt bậc 3,5/7 (Lắp đồng hồ áp suất)': 262300, # Nhóm 3 bậc 3,5/7: 262.300
    'Thợ lắp đặt thiết bị PCCC bậc 3,0/7 (Đầu phun xả tràn)': 239866, # Nhóm 3 bậc 3,0/7: 262.300 * 1.39 / 1.52 = 239.866
    'Thợ cơ điện bậc 4,0/7 (Bơm chữa cháy)': 276809, # Nhóm 2 bậc 4,0/7: 255.000 * 1.65 / 1.52 = 276.809
    'Thợ điện công nghiệp bậc 3,0/7 (Đèn sự cố & EXIT)': 233191, # Nhóm 2 bậc 3,0/7: 255.000 * 1.39 / 1.52 = 233.191
    'Thợ điện bậc 3,0/7 (Tháo chuông còi báo cháy)': 233191, # Nhóm 2 bậc 3,0/7: 255.000 * 1.39 / 1.52 = 233.191
    'Thợ sơn bậc 3,5/7': 231200, # Nhóm 1 bậc 3,5/7: 231.200
    'Thợ sơn bậc 3,5/7 (Làm việc ngoài trời)': 231200, # Nhóm 1 bậc 3,5/7: 231.200
    'Kỹ sư công nghệ thí nghiệm điện bậc 5,0/8': 312339 # Kỹ sư bậc 5/8: 285.800 * 1.53 / 1.40 = 312.339
}

# Legal basis mapping for Master Rows
nc_basis = {
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 theo QĐ 177 & QĐ 45)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công tháo dỡ trên cao (Nhóm 3 - Bậc 5,0/7 - E6C.3.01 bước 5)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công bảo dưỡng, kiểm chuẩn sàn (Nhóm 3 - Bậc 5,0/7)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công tháo dỡ trên cao (Nhóm 3 - Bậc 5,0/7 - E6C.3.02 bước 5)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công hiệu chuẩn, thử nghiệm sàn (Nhóm 3 - Bậc 5,0/7)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công tháo dỡ beam trên cao (Vận dụng E6C.3.02 bước 5)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Cân chỉnh beam)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Module loop)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Cụm van Deluge)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 3 - Bậc 5,0/7 - Đại tu van nước)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Thợ cơ khí sửa chữa bậc 4,0/7 (Đại tu van nước PCCC)': 'QĐ 177/QĐ-HĐTV (Nhóm 3) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Van Solenoid)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Công tắc áp lực)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Đo lưu lượng)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Hiệu chuẩn khí H2)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Kiểm tra ắc quy)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Nhân công kỹ thuật (Nhóm 4 - Bậc 5,0/7 - Van an toàn)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Thợ cơ khí sửa chữa bậc 4,5/7 (Đại tu van an toàn)': 'QĐ 177/QĐ-HĐTV (Nhóm 4) & QĐ 311/QĐ-SXD Bình Thuận Vùng IV',
    'Thợ lắp đặt thiết bị PCCC bậc 3,5/7': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 3 bậc 3,5/7)',
    'Thợ lắp đặt đường ống PCCC bậc 3,5/7': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 3 bậc 3,5/7)',
    'Thợ lắp đặt bậc 3,5/7 (Lắp van ren ≤ DN25)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 3 bậc 3,5/7)',
    'Thợ lắp đặt bậc 3,5/7 (Lắp van bích DN80)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 3 bậc 3,5/7)',
    'Thợ lắp đặt bậc 3,5/7 (Lắp đồng hồ áp suất)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 3 bậc 3,5/7)',
    'Thợ lắp đặt thiết bị PCCC bậc 3,0/7 (Đầu phun xả tràn)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 3 bậc 3,0/7)',
    'Thợ cơ điện bậc 4,0/7 (Bơm chữa cháy)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 2 bậc 4,0/7)',
    'Thợ điện công nghiệp bậc 3,0/7 (Đèn sự cố & EXIT)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 2 bậc 3,0/7)',
    'Thợ điện bậc 3,0/7 (Tháo chuông còi báo cháy)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 2 bậc 3,0/7)',
    'Thợ sơn bậc 3,5/7': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 1 bậc 3,5/7)',
    'Thợ sơn bậc 3,5/7 (Làm việc ngoài trời)': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Nhóm 1 bậc 3,5/7)',
    'Kỹ sư công nghệ thí nghiệm điện bậc 5,0/8': 'TT 38/2026 & QĐ 311/QĐ-SXD Bình Thuận Vùng IV (Kỹ sư bậc 5,0/8)'
}

may_prices = {
    'Đồng hồ vạn năng kỹ thuật số (VOM)': 60000,
    'Đồng hồ vạn năng (VOM)': 60000,
    'Đồng hồ vạn năng VOM': 60000,
    'Đồng hồ vạn năng VOM & Thiết bị kích': 80000,
    'Máy thổi bụi công nghiệp cầm tay': 50000,
    'Đầu thử nhiệt Solo 461 / Đầu thử khói': 150000,
    'Mê gôm mét 1000V': 90000,
    'Mê gôm mét 500V': 80000,
    'Máy đo điện trở cách điện': 95000,
    'Máy kiểm tra áp lực nước mini cầm tay': 120000,
    'Thiết bị đo rung chấn & Ampe kìm công nghiệp': 200000,
    'Đồng hồ đo áp lực lưu lượng chuyên dụng': 180000,
    'Bơm tạo áp cầm tay & Đồng hồ VOM': 110000,
    'Bình khí mẫu hiệu chuẩn (Zero & Span gas)': 1500000,
    'Đầu thử khói Solo 330 / Kính ngắm laser': 160000,
    'Dụng cụ gỡ đầu báo khói Solo 200': 40000,
    'Thiết bị đo nội trở và dung lượng accu': 300000,
    'Bàn cân áp suất / Bơm tạo áp': 350000,
    'Thiết bị nâng hạ / Palăng xích & Cờ lê lực': 150000,
    'Bàn thử áp lực van an toàn chuyên dùng': 650000,
    'Thiết bị đo điện trở đất chuyên dụng (Terrameter)': 250000
}

vl_prices = {
    'Sơn chống gỉ Alkyd đỏ PCCC (100 m2)': 68000,
    'Sơn phủ màu đỏ PCCC (100 m2)': 85000,
    'Nước cất nạp accu, mỡ bôi cọc bình': 45000
}

# Styling definitions
font_comp = Font(name='Times New Roman', size=10, bold=True, color='002060')
font_title = Font(name='Times New Roman', size=13, bold=True, color='002060')
font_sub = Font(name='Times New Roman', size=10, italic=True, color='444444')

font_tbl_hdr = Font(name='Times New Roman', size=9.5, bold=True, color='000000')
font_master = Font(name='Times New Roman', size=9.5, bold=True, color='002060')
font_detail = Font(name='Times New Roman', size=9.0, color='000000')
font_total = Font(name='Times New Roman', size=10, bold=True, color='9C0006')

fill_tbl_hdr = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_master_nc = PatternFill(start_color='EDF2F8', end_color='EDF2F8', fill_type='solid')
fill_master_may = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
fill_master_vl = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
fill_total = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')

border_thin = Side(border_style='thin', color='B0B0B0')
border_thick_bottom = Side(border_style='medium', color='002060')
border_double = Side(border_style='double', color='000000')

box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
hdr_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thick_bottom)
total_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_double)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

def create_resource_sheet(wb, sheet_title, resource_type_label, groups_dict, prices_dict, fill_master):
    if sheet_title in wb.sheetnames:
        wb.remove(wb[sheet_title])
        
    ws = wb.create_sheet(title=sheet_title)
    ws.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws.merge_cells('A1:I1')
    ws['A1'] = 'TẬP ĐOÀN ĐIỆN LỰC VIỆT NAM - TỔNG CÔNG TY PHÁT ĐIỆN 3 - CÔNG TY NHIỆT ĐIỆN VĨNH TÂN 4'
    ws['A1'].font = font_comp
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A2:I2')
    ws['A2'] = f'BẢNG TỔNG HỢP VÀ PHÂN BỔ CHI TIẾT {resource_type_label.upper()} (MÔ HÌNH MASTER - DETAIL)'
    ws['A2'].font = font_title
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A3:I3')
    ws['A3'] = f'Dự án: SCL Hệ thống PCCC năm 2026 – Chuẩn bị Đơn giá chi tiết từ Sheet "hao phi VT" (Tổng khối lượng = SUM row con)'
    ws['A3'].font = font_sub
    ws['A3'].alignment = Alignment(horizontal='center', vertical='center')

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 25
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 10

    # Headers at Row 5
    headers = [
        'STT',
        'Mã hiệu\ntài nguyên',
        f'Tên {resource_type_label.lower()} (Master) &\nTên công tác áp dụng trong 856 dịch vụ (Detail)',
        'Khu vực /\nHạng mục công trình',
        'Đơn vị\ntính (ĐVT)',
        'Khối lượng hao phí\n(Row cha = SUM row con)',
        'Đơn giá tạm tính\n(VNĐ)',
        'Thành tiền\n(VNĐ)',
        'Căn cứ pháp lý /\nGhi chú liên kết'
    ]

    ws.row_dimensions[5].height = 32
    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=c_idx, value=h_text)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.border = hdr_border
        cell.alignment = align_center

    curr_r = 6
    master_idx = 1
    master_rows_for_total = []

    for res_name, details in groups_dict.items():
        master_r = curr_r
        master_rows_for_total.append(master_r)
        
        # Get representative item info
        sample = details[0]
        res_code = sample['code']
        res_dvt = sample['dvt']
        unit_price = prices_dict.get(res_name, 350000)
        
        child_start_r = curr_r + 1
        child_end_r = curr_r + len(details)
        
        # 1. Ghi ROW CHA (MASTER ROW)
        ws.row_dimensions[curr_r].height = 24
        ws.cell(row=curr_r, column=1, value=master_idx).alignment = align_center
        ws.cell(row=curr_r, column=2, value=res_code).alignment = align_center
        ws.cell(row=curr_r, column=3, value=f"[LOẠI {master_idx}] {res_name.upper()}").alignment = align_left
        ws.cell(row=curr_r, column=4, value=f"Tổng hợp toàn bộ {len(details)} công tác áp dụng").alignment = align_left
        ws.cell(row=curr_r, column=5, value=res_dvt).alignment = align_center
        
        # Công thức SUM toàn bộ row con
        c_qty_master = ws.cell(row=curr_r, column=6, value=f"=SUM(F{child_start_r}:F{child_end_r})")
        c_qty_master.alignment = align_right
        c_qty_master.number_format = '#,##0.0000'
        
        # Đơn giá tạm tính
        c_price_master = ws.cell(row=curr_r, column=7, value=unit_price)
        c_price_master.alignment = align_right
        c_price_master.number_format = '#,##0'
        
        # Thành tiền = Khối lượng * Đơn giá
        c_amount_master = ws.cell(row=curr_r, column=8, value=f"=F{curr_r}*G{curr_r}")
        c_amount_master.alignment = align_right
        c_amount_master.number_format = '#,##0'
        
        if resource_type_label == 'Nhân công':
            basis_text = nc_basis.get(res_name, f"Tổng hao phí {len(details)} công tác")
        else:
            basis_text = f"Tổng hao phí {len(details)} công tác"
        ws.cell(row=curr_r, column=9, value=basis_text).alignment = align_left
        
        for c in range(1, 10):
            cell = ws.cell(row=curr_r, column=c)
            cell.font = font_master
            cell.fill = fill_master
            cell.border = box_border
            
        curr_r += 1
        
        # 2. Ghi CÁC ROW CON (DETAIL ROWS)
        for d in details:
            hp_row = d['hp_row']
            col_letter = d['col_letter']
            task_stt = d['stt']
            task_name = d['task_name']
            task_khuvuc = d['khu_vuc']
            
            ws.row_dimensions[curr_r].height = 19
            ws.cell(row=curr_r, column=1, value=f"STT {task_stt}").alignment = align_center
            ws.cell(row=curr_r, column=2, value=d['code']).alignment = align_center
            ws.cell(row=curr_r, column=3, value=f"   • STT {task_stt}: {task_name}").alignment = align_left
            ws.cell(row=curr_r, column=4, value=str(task_khuvuc)).alignment = align_left
            ws.cell(row=curr_r, column=5, value=res_dvt).alignment = align_center
            
            # LINK TRỰC TIẾP ĐẾN SHEET 'hao phi VT'
            c_qty_detail = ws.cell(row=curr_r, column=6, value=f"='hao phi VT'!{col_letter}{hp_row}")
            c_qty_detail.alignment = align_right
            c_qty_detail.number_format = '#,##0.0000'
            
            # Đơn giá & Thành tiền để trống ở row con
            ws.cell(row=curr_r, column=7, value='').alignment = align_right
            ws.cell(row=curr_r, column=8, value='').alignment = align_right
            
            ws.cell(row=curr_r, column=9, value=f"Link từ 'hao phi VT' ô {col_letter}{hp_row}").alignment = align_left
            
            for c in range(1, 10):
                cell = ws.cell(row=curr_r, column=c)
                cell.font = font_detail
                cell.border = box_border
                
            curr_r += 1
            
        master_idx += 1

    # Dòng TỔNG CỘNG TOÀN TOÀN BỘ SHEET
    total_r = curr_r
    ws.row_dimensions[total_r].height = 26
    ws.cell(row=total_r, column=1, value="TỔNG").alignment = align_center
    ws.cell(row=total_r, column=2, value="").alignment = align_center
    ws.cell(row=total_r, column=3, value=f"TỔNG CỘNG CHI PHÍ {resource_type_label.upper()} (TẤT CẢ CÁC LOẠI MASTER)").alignment = align_left
    ws.cell(row=total_r, column=4, value="").alignment = align_left
    ws.cell(row=total_r, column=5, value="").alignment = align_center
    
    # Tổng khối lượng (nếu cùng ĐVT) hoặc sum các master
    qty_sum_formula = "+".join([f"F{mr}" for mr in master_rows_for_total])
    ws.cell(row=total_r, column=6, value=f"={qty_sum_formula}").alignment = align_right
    ws.cell(row=total_r, column=6).number_format = '#,##0.00'
    
    ws.cell(row=total_r, column=7, value="").alignment = align_right
    
    # Tổng thành tiền = Sum thành tiền các master row
    amount_sum_formula = "+".join([f"H{mr}" for mr in master_rows_for_total])
    c_tot_amount = ws.cell(row=total_r, column=8, value=f"={amount_sum_formula}")
    c_tot_amount.alignment = align_right
    c_tot_amount.number_format = '#,##0'
    
    ws.cell(row=total_r, column=9, value=f"Chuẩn bị nguồn dữ liệu cho Đơn giá chi tiết").alignment = align_left
    
    for c in range(1, 10):
        cell = ws.cell(row=total_r, column=c)
        cell.font = font_total
        cell.fill = fill_total
        cell.border = total_border

    # Set Column widths
    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 30
    ws.column_dimensions['E'].width = 11
    ws.column_dimensions['F'].width = 18
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 30
    
    ws.freeze_panes = 'E6'
    print(f"Sheet '{sheet_title}' created successfully with {total_r} rows.")

# 1. Tạo sheet Nhân công
print("\nCreating sheet 'Nhân công'...")
create_resource_sheet(wb, 'Nhân công', 'Nhân công', nc_groups, nc_prices, fill_master_nc)

# 2. Tạo sheet Máy thi công
print("\nCreating sheet 'Máy thi công'...")
create_resource_sheet(wb, 'Máy thi công', 'Máy thi công', may_groups, may_prices, fill_master_may)

# 3. Tạo sheet Vật liệu
print("\nCreating sheet 'Vật liệu'...")
create_resource_sheet(wb, 'Vật liệu', 'Vật liệu', vl_groups, vl_prices, fill_master_vl)

# Save workbooks
print("\nSaving updated workbook locally...")
wb.save(wb_path)
print(f"Saved: {wb_path}")

print("Saving to Downloads...")
try:
    wb.save(dl_path)
    print(f"Saved: {dl_path}")
except Exception as e:
    alt_path = 'C:/Users/Admin/Downloads/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_CapNhat.xlsx'
    wb.save(alt_path)
    print(f"Fallback saved: {alt_path} (Error: {e})")

print("\nAll 3 resource sheets successfully created!")
