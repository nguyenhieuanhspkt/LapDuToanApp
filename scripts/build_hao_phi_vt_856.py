import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb_path = 'd:/TaskApp_kiet/LapDuToanApp/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'
dl_path = 'C:/Users/Admin/Downloads/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC.xlsx'

print("Loading workbook...")
wb = openpyxl.load_workbook(wb_path)

ws_02 = wb['02_Chi_Tiet_856_Dich_Vu']
sheet_name = 'hao phi VT'

if sheet_name in wb.sheetnames:
    wb.remove(wb[sheet_name])

ws = wb.create_sheet(title=sheet_name)
ws.views.sheetView[0].showGridLines = True

# Styling
font_tbl_hdr = Font(name='Times New Roman', size=9.5, bold=True, color='000000')
font_cong_tac = Font(name='Times New Roman', size=9.5, bold=True, color='002060')
font_item = Font(name='Times New Roman', size=9.0, color='000000')
font_zero = Font(name='Times New Roman', size=9.0, italic=True, color='888888')

fill_tbl_hdr = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
fill_cong_tac = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')

border_thin = Side(border_style='thin', color='C0C0C0')
border_med = Side(border_style='medium', color='002060')
box_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
hdr_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_med)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

# Title / Blank Rows 1-4
ws.row_dimensions[1].height = 15
ws.row_dimensions[2].height = 15
ws.row_dimensions[3].height = 15
ws.row_dimensions[4].height = 15

# Row 5-6: Headers matching template image
ws.merge_cells('A5:A6')
ws.merge_cells('B5:B6')
ws.merge_cells('C5:C6')
ws.merge_cells('D5:D6')
ws.merge_cells('E5:E6')
ws.merge_cells('F5:F6')
ws.merge_cells('G5:G6')
ws.merge_cells('H5:K5') # MỨC HAO PHÍ
ws.merge_cells('L5:L6')
ws.merge_cells('M5:P5') # KHỐI LƯỢNG HAO PHÍ
ws.merge_cells('Q5:Q6')

ws['A5'] = 'STT'
ws['B5'] = 'Mã hiệu\nđịnh mức'
ws['C5'] = 'Phân loại'
ws['D5'] = 'Giai đoạn /\nKhu vực'
ws['E5'] = 'Tên công tác / Thành phần hao phí chi tiết'
ws['F5'] = 'Đơn vị'
ws['G5'] = 'Khối lượng\ncông trình (Q)'

ws['H5'] = 'MỨC HAO PHÍ'
ws['H6'] = 'Vật liệu'
ws['I6'] = 'Nhân công'
ws['J6'] = 'Công cụ dụng\ncụ'
ws['K6'] = 'Máy thi công'

ws['L5'] = 'Hệ số'

ws['M5'] = 'KHỐI LƯỢNG HAO PHÍ'
ws['M6'] = 'Vật liệu'
ws['N6'] = 'Nhân công'
ws['O6'] = 'Công cụ dụng cụ'
ws['P6'] = 'Máy thi công'

ws['Q5'] = 'Căn cứ pháp lý định mức\n(QĐ 45/QĐ-HĐTV & TT 38/2026)'

ws.row_dimensions[5].height = 24
ws.row_dimensions[6].height = 24

for r in [5, 6]:
    for c in range(1, 18):
        cell = ws.cell(r, c)
        cell.font = font_tbl_hdr
        cell.fill = fill_tbl_hdr
        cell.border = box_border
        cell.alignment = align_center

curr_r = 7

# Function to get component template for any code
def get_components_for_code(code_str):
    c = str(code_str).strip()
    
    # 1. Tủ trung tâm E6C.3.05
    if c == 'E6C.3.05':
        return [
            ('NC_E6C.3.05', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Trọn bộ 11 bước công nghệ)', 'công', 'I', "='Định mức áp dụng'!D61", 1.0, 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451'),
            ('M_E6C.3.05_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng kỹ thuật số (VOM)', 'ca', 'K', "='Định mức áp dụng'!D62", 1.0, 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451'),
            ('M_E6C.3.05_BUI', 'MÁY THI CÔNG', 'Hiện trường', '- Máy thổi bụi công nghiệp cầm tay', 'ca', 'K', "='Định mức áp dụng'!D63", 1.0, 'Quyết định 45/QĐ-HĐTV Tập 6 trang 451')
        ]
        
    # 2. Không áp mã (đã gồm trong E6C.3.05)
    elif 'Không áp mã (đã gồm trong E6C.3.05)' in c:
        return [
            ('NC_E6C.3.05', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Đã gồm trong tủ E6C.3.05)', 'công', 'I', "='Định mức áp dụng'!D61", 0.0, 'Đã tính trong tủ trung tâm E6C.3.05 (Hệ số = 0)'),
            ('M_E6C.3.05_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM) (Đã gồm trong tủ E6C.3.05)', 'ca', 'K', "='Định mức áp dụng'!D62", 0.0, 'Đã tính trong tủ trung tâm E6C.3.05 (Hệ số = 0)'),
            ('M_E6C.3.05_BUI', 'MÁY THI CÔNG', 'Hiện trường', '- Máy thổi bụi cầm tay (Đã gồm trong tủ E6C.3.05)', 'ca', 'K', "='Định mức áp dụng'!D63", 0.0, 'Đã tính trong tủ trung tâm E6C.3.05 (Hệ số = 0)')
        ]

    # 3. E6C.3.01 (bước 5) - Pha 1 tháo dỡ
    elif 'E6C.3.01' in c and 'bước 5' in c:
        return [
            ('NC_E6C.3.01_B5', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công tháo dỡ trên cao (E6C.3.01 bước 5)', 'công', 'I', "='Định mức áp dụng'!D71", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448 (Bước 5: 30%)'),
            ('NC_E6C.3.01_B1-11', 'NHÂN CÔNG', 'Phòng kiểm chuẩn', '- Nhân công bảo dưỡng, kiểm chuẩn sàn (Bước 1-4, 6-11)', 'công', 'I', 0.24, 0.0, 'Không áp dụng tại pha tháo dỡ (Hệ số = 0)'),
            ('M_E6C.3.01_VOM', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D80", 0.0, 'Không áp dụng tại pha tháo dỡ (Hệ số = 0)'),
            ('M_E6C.3.01_MEG', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Mê gôm mét 1000V', 'ca', 'K', "='Định mức áp dụng'!D81", 0.0, 'Không áp dụng tại pha tháo dỡ (Hệ số = 0)')
        ]

    # 4. E6C.3.01 (bước 1–4, 6–11) - Pha 2 kiểm chuẩn sàn
    elif 'E6C.3.01' in c and ('1-4' in c or '1–4' in c):
        return [
            ('NC_E6C.3.01_B5', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công tháo dỡ trên cao (E6C.3.01 bước 5)', 'công', 'I', "='Định mức áp dụng'!D71", 0.0, 'Đã tính tại pha tháo dỡ (Hệ số = 0)'),
            ('NC_E6C.3.01_B1-11', 'NHÂN CÔNG', 'Phòng kiểm chuẩn', '- Nhân công bảo dưỡng, kiểm chuẩn sàn (Bước 1-4, 6-11)', 'công', 'I', 0.24, 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448 (Bước 1-4,6-11: 70%)'),
            ('M_E6C.3.01_VOM', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D80", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448'),
            ('M_E6C.3.01_MEG', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Mê gôm mét 1000V', 'ca', 'K', "='Định mức áp dụng'!D81", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448')
        ]

    # 5. E6C.3.01 trọn gói
    elif c == 'E6C.3.01':
        return [
            ('NC_E6C.3.01', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Trọn bộ 11 bước)', 'công', 'I', "='Định mức áp dụng'!D79", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448'),
            ('M_E6C.3.01_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D80", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448'),
            ('M_E6C.3.01_MEG', 'MÁY THI CÔNG', 'Hiện trường', '- Mê gôm mét 1000V', 'ca', 'K', "='Định mức áp dụng'!D81", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 448')
        ]

    # 6. E6C.3.02 (bước 5) - Pha 1 tháo dỡ
    elif 'E6C.3.02' in c and 'bước 5' in c:
        return [
            ('NC_E6C.3.02_B5', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công tháo dỡ trên cao (E6C.3.02 bước 5)', 'công', 'I', "='Định mức áp dụng'!D89", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449 (Bước 5: 30%)'),
            ('NC_E6C.3.02_B1-11', 'NHÂN CÔNG', 'Phòng kiểm chuẩn', '- Nhân công hiệu chuẩn, thử nghiệm sàn (Bước 1-4, 6-11)', 'công', 'I', 1.40, 0.0, 'Không áp dụng tại pha tháo dỡ (Hệ số = 0)'),
            ('M_E6C.3.02_SOLO', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Đầu thử nhiệt Solo 461 / Đầu thử khói', 'ca', 'K', "='Định mức áp dụng'!D98", 0.0, 'Không áp dụng tại pha tháo dỡ (Hệ số = 0)'),
            ('M_E6C.3.02_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D99", 0.0, 'Không áp dụng tại pha tháo dỡ (Hệ số = 0)')
        ]

    # 7. E6C.3.02 (bước 1–4, 6–11) - Pha 2 kiểm chuẩn sàn
    elif 'E6C.3.02' in c and ('1-4' in c or '1–4' in c):
        return [
            ('NC_E6C.3.02_B5', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công tháo dỡ trên cao (E6C.3.02 bước 5)', 'công', 'I', "='Định mức áp dụng'!D89", 0.0, 'Đã tính tại pha tháo dỡ (Hệ số = 0)'),
            ('NC_E6C.3.02_B1-11', 'NHÂN CÔNG', 'Phòng kiểm chuẩn', '- Nhân công hiệu chuẩn, thử nghiệm sàn (Bước 1-4, 6-11)', 'công', 'I', 1.40, 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449 (Bước 1-4,6-11: 70%)'),
            ('M_E6C.3.02_SOLO', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Đầu thử nhiệt Solo 461 / Đầu thử khói', 'ca', 'K', "='Định mức áp dụng'!D98", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449'),
            ('M_E6C.3.02_VOM', 'MÁY THI CÔNG', 'Phòng kiểm chuẩn', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D99", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449')
        ]

    # 8. E6C.3.02 trọn gói
    elif c == 'E6C.3.02':
        return [
            ('NC_E6C.3.02', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Trọn bộ 11 bước)', 'công', 'I', "='Định mức áp dụng'!D97", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449'),
            ('M_E6C.3.02_SOLO', 'MÁY THI CÔNG', 'Hiện trường', '- Đầu thử nhiệt Solo 461 / Đầu thử khói', 'ca', 'K', "='Định mức áp dụng'!D98", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449'),
            ('M_E6C.3.02_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D99", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 449')
        ]

    # 9. E6C.3.04 (Beam cân chỉnh)
    elif 'E6C.3.04' in c:
        return [
            ('NC_E6C.3.04', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Cân chỉnh beam quang)', 'công', 'I', "='Định mức áp dụng'!D115", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 451'),
            ('M_E6C.3.04_SOLO', 'MÁY THI CÔNG', 'Hiện trường', '- Đầu thử khói Solo 330 / Kính ngắm laser', 'ca', 'K', "='Định mức áp dụng'!D116", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 451'),
            ('M_E6C.3.04_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D117", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 451'),
            ('DC_E6C.3.04_S200', 'DỤNG CỤ', 'Hiện trường', '- Dụng cụ gỡ đầu báo khói Solo 200', 'ca', 'J', "='Định mức áp dụng'!D118", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 451')
        ]

    # 10. E6C.3.03 (Module loop)
    elif 'E6C.3.03' in c:
        return [
            ('NC_E6C.3.03', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Kiểm tra module loop)', 'công', 'I', "='Định mức áp dụng'!D132", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 450'),
            ('M_E6C.3.03_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D133", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 450'),
            ('M_E6C.3.03_MEG', 'MÁY THI CÔNG', 'Hiện trường', '- Mê gôm mét 500V', 'ca', 'K', "='Định mức áp dụng'!D134", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 450')
        ]

    # 11. E6C.2.04 (Solenoid)
    elif 'E6C.2.04' in c:
        return [
            ('NC_E6C.2.04', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Bảo dưỡng van điện từ solenoid)', 'công', 'I', "='Định mức áp dụng'!D197", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 440'),
            ('M_E6C.2.04_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng VOM & Thiết bị kích', 'ca', 'K', "='Định mức áp dụng'!D198", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 440')
        ]

    # 12. E6C.1.23 (Công tắc áp lực/dòng chảy)
    elif 'E6C.1.23' in c:
        return [
            ('NC_E6C.1.23', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Kiểm tra công tắc áp lực)', 'công', 'I', "='Định mức áp dụng'!D201", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 428'),
            ('M_E6C.1.23_BOM', 'MÁY THI CÔNG', 'Hiện trường', '- Bơm tạo áp cầm tay & Đồng hồ VOM', 'ca', 'K', "='Định mức áp dụng'!D202", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 428')
        ]

    # 13. E6C.1.18 (Đồng hồ lưu lượng)
    elif 'E6C.1.18' in c:
        return [
            ('NC_E6C.1.18', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Kiểm tra đo lưu lượng)', 'công', 'I', "='Định mức áp dụng'!D205", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 425'),
            ('M_E6C.1.18_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng VOM', 'ca', 'K', "='Định mức áp dụng'!D206", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 425')
        ]

    # 14. E6C.3.06 (Van Deluge)
    elif 'E6C.3.06' in c:
        return [
            ('NC_E6C.3.06', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Cụm van Deluge phần điện)', 'công', 'I', "='Định mức áp dụng'!D152", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 453'),
            ('M_E6C.3.06_CD', 'MÁY THI CÔNG', 'Hiện trường', '- Máy đo điện trở cách điện', 'ca', 'K', "='Định mức áp dụng'!D153", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 453'),
            ('M_E6C.3.06_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D154", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 453'),
            ('M_E6C.3.06_BAN', 'MÁY THI CÔNG', 'Hiện trường', '- Bàn cân áp suất / Bơm tạo áp', 'ca', 'K', "='Định mức áp dụng'!D155", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 453')
        ]

    # 15. E6C.1.31 (Đo khí H2)
    elif 'E6C.1.31' in c:
        return [
            ('NC_E6C.1.31', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Hiệu chuẩn đầu đo khí H2)', 'công', 'I', "='Định mức áp dụng'!D171", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 436'),
            ('M_E6C.1.31_GAS', 'MÁY THI CÔNG', 'Hiện trường', '- Bình khí mẫu hiệu chuẩn (Zero & Span gas)', 'ca', 'K', "='Định mức áp dụng'!D172", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 436'),
            ('M_E6C.1.31_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng kỹ thuật số (VOM)', 'ca', 'K', "='Định mức áp dụng'!D173", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 436')
        ]

    # 16. E6B.1.13.01 (Bình ắc quy)
    elif 'E6B.1.13.01' in c:
        return [
            ('NC_E6B.1.13.01', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kỹ thuật bậc 5,0/7 (Kiểm tra bình ắc quy)', 'công', 'I', "='Định mức áp dụng'!D191", 1.0, 'QĐ 45/QĐ-HĐTV Tập 5 trang 383'),
            ('M_E6B.1.13.01_TR', 'MÁY THI CÔNG', 'Hiện trường', '- Thiết bị đo nội trở và dung lượng accu', 'ca', 'K', "='Định mức áp dụng'!D192", 1.0, 'QĐ 45/QĐ-HĐTV Tập 5 trang 383'),
            ('M_E6B.1.13.01_VOM', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ vạn năng (VOM)', 'ca', 'K', "='Định mức áp dụng'!D193", 1.0, 'QĐ 45/QĐ-HĐTV Tập 5 trang 383'),
            ('VL_E6B.1.13.01', 'VẬT LIỆU', 'Hiện trường', '- Nước cất nạp accu, mỡ bôi cọc bình', 'kg', 'H', "='Định mức áp dụng'!D194", 1.0, 'QĐ 45/QĐ-HĐTV Tập 5 trang 383')
        ]

    # 17. E6BOP.1.03 (Đại tu van nước cơ khí)
    elif 'E6BOP.1.03' in c or 'E6BOP.2.01' in c or 'E6BOP.2.16' in c or 'E6BOP.5.03' in c or 'E6BOP.6.01' in c:
        return [
            ('NC_E6BOP.1.03', 'NHÂN CÔNG', 'Hiện trường', '- Thợ cơ khí sửa chữa bậc 4,0/7 (Đại tu van nước PCCC)', 'công', 'I', "='Định mức áp dụng'!D209", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 312'),
            ('M_E6BOP.1.03', 'MÁY THI CÔNG', 'Hiện trường', '- Thiết bị nâng hạ / Palăng xích & Cờ lê lực', 'ca', 'K', "='Định mức áp dụng'!D210", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 312')
        ]

    # 18. E6BOP.1.01 (Đại tu van an toàn)
    elif 'E6BOP.1.01' in c:
        return [
            ('NC_E6BOP.1.01', 'NHÂN CÔNG', 'Hiện trường', '- Thợ cơ khí sửa chữa bậc 4,5/7 (Đại tu van an toàn)', 'công', 'I', "='Định mức áp dụng'!D213", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 308'),
            ('M_E6BOP.1.01', 'MÁY THI CÔNG', 'Hiện trường', '- Bàn thử áp lực van an toàn chuyên dùng', 'ca', 'K', "='Định mức áp dụng'!D214", 1.0, 'QĐ 45/QĐ-HĐTV Tập 6 trang 308')
        ]

    # 19. BD.41111 (Họng nước / Đầu báo)
    elif 'BD.41111' in c:
        return [
            ('NC_BD.41111', 'NHÂN CÔNG', 'Hiện trường', '- Thợ lắp đặt thiết bị PCCC bậc 3,5/7', 'công', 'I', "='Định mức áp dụng'!D219", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3'),
            ('M_BD.41111', 'MÁY THI CÔNG', 'Hiện trường', '- Máy kiểm tra áp lực nước mini cầm tay', 'ca', 'K', "='Định mức áp dụng'!D220", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 20. BD.41121 (Trụ cứu hỏa / Đèn báo cháy)
    elif 'BD.41121' in c:
        return [
            ('NC_BD.41121', 'NHÂN CÔNG', 'Hiện trường', '- Thợ lắp đặt đường ống PCCC bậc 3,5/7', 'công', 'I', "='Định mức áp dụng'!D222", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3'),
            ('M_BD.41121', 'MÁY THI CÔNG', 'Hiện trường', '- Đồng hồ đo áp lực lưu lượng chuyên dụng', 'ca', 'K', "='Định mức áp dụng'!D223", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 21. BD.41131 (Cụm máy bơm / Nút ấn)
    elif 'BD.41131' in c:
        return [
            ('NC_BD.41131', 'NHÂN CÔNG', 'Hiện trường', '- Thợ cơ điện bậc 4,0/7 (Bơm chữa cháy)', 'công', 'I', "='Định mức áp dụng'!D225", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3'),
            ('M_BD.41131', 'MÁY THI CÔNG', 'Hiện trường', '- Thiết bị đo rung chấn & Ampe kìm công nghiệp', 'ca', 'K', "='Định mức áp dụng'!D226", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 22. BD.41141 (Đầu phun Deluge / Chuông)
    elif 'BD.41141' in c:
        return [
            ('NC_BD.41141', 'NHÂN CÔNG', 'Hiện trường', '- Thợ lắp đặt thiết bị PCCC bậc 3,0/7 (Đầu phun xả tràn)', 'công', 'I', "='Định mức áp dụng'!D228", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 23. BD.41171 (Đèn EXIT)
    elif 'BD.41171' in c:
        return [
            ('NC_BD.41171', 'NHÂN CÔNG', 'Hiện trường', '- Thợ điện công nghiệp bậc 3,0/7 (Đèn sự cố & EXIT)', 'công', 'I', "='Định mức áp dụng'!D230", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 24. BA.18405 (Tháo dỡ chuông còi)
    elif 'BA.18405' in c:
        return [
            ('NC_BA.18405', 'NHÂN CÔNG', 'Hiện trường', '- Thợ điện bậc 3,0/7 (Tháo chuông còi báo cháy)', 'công', 'I', "='Định mức áp dụng'!D232", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 25. SB.82633 (Sơn trong nhà)
    elif 'SB.82633' in c:
        return [
            ('VL_SB.82633_LOT', 'VẬT LIỆU', 'Hiện trường', '- Sơn chống gỉ Alkyd đỏ PCCC (100 m2)', 'kg', 'H', "='Định mức áp dụng'!D234", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 6'),
            ('VL_SB.82633_PHU', 'VẬT LIỆU', 'Hiện trường', '- Sơn phủ màu đỏ PCCC (100 m2)', 'kg', 'H', "='Định mức áp dụng'!D235", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 6'),
            ('NC_SB.82633', 'NHÂN CÔNG', 'Hiện trường', '- Thợ sơn bậc 3,5/7', 'công', 'I', "='Định mức áp dụng'!D236", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 6')
        ]

    # 26. SB.82634 (Sơn ngoài trời)
    elif 'SB.82634' in c:
        return [
            ('NC_SB.82634', 'NHÂN CÔNG', 'Hiện trường', '- Thợ sơn bậc 3,5/7 (Làm việc ngoài trời)', 'công', 'I', "='Định mức áp dụng'!D238", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 6')
        ]

    # 27. BB.86601 (Van ren)
    elif 'BB.86601' in c:
        return [
            ('NC_BB.86601', 'NHÂN CÔNG', 'Hiện trường', '- Thợ lắp đặt bậc 3,5/7 (Lắp van ren ≤ DN25)', 'công', 'I', "='Định mức áp dụng'!D240", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 28. BB.86103 (Van bích)
    elif 'BB.86103' in c:
        return [
            ('NC_BB.86103', 'NHÂN CÔNG', 'Hiện trường', '- Thợ lắp đặt bậc 3,5/7 (Lắp van bích DN80)', 'công', 'I', "='Định mức áp dụng'!D241", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 29. BB.85201 (Đồng hồ áp)
    elif 'BB.85201' in c:
        return [
            ('NC_BB.85201', 'NHÂN CÔNG', 'Hiện trường', '- Thợ lắp đặt bậc 3,5/7 (Lắp đồng hồ áp suất)', 'công', 'I', "='Định mức áp dụng'!D242", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 3')
        ]

    # 30. DB.24001 (Tiếp địa)
    elif 'DB.24001' in c:
        return [
            ('NC_DB.24001', 'NHÂN CÔNG', 'Hiện trường', '- Kỹ sư công nghệ thí nghiệm điện bậc 5,0/8', 'công', 'I', "='Định mức áp dụng'!D244", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 5'),
            ('M_DB.24001', 'MÁY THI CÔNG', 'Hiện trường', '- Thiết bị đo điện trở đất chuyên dụng (Terrameter)', 'ca', 'K', "='Định mức áp dụng'!D245", 1.0, 'Thông tư 38/2026/TT-BXD Phụ lục 5')
        ]

    # 31. Các mã Không áp mã khác (bảo dưỡng thường xuyên / phụ trợ)
    else:
        return [
            ('NC_KDM', 'NHÂN CÔNG', 'Hiện trường', '- Nhân công kiểm tra phụ trợ / kiểm tra thường xuyên', 'công', 'I', 0.10, 0.0, 'Chi phí bảo dưỡng thường xuyên / quản lý chung (Hệ số = 0)')
        ]

print("Processing 856 tasks...")

# Populate 856 tasks
for r in range(5, ws_02.max_row + 1):
    stt_val = ws_02.cell(r, 1).value
    khu_vuc = ws_02.cell(r, 2).value
    ten_ct = ws_02.cell(r, 3).value
    q_val = ws_02.cell(r, 4).value
    dvt_val = ws_02.cell(r, 5).value
    ma_dm_val = ws_02.cell(r, 9).value
    can_cu_val = ws_02.cell(r, 10).value
    
    if stt_val is None:
        continue
        
    q_row = curr_r
    
    # 1. DÒNG CÔNG TÁC (Chính xác STT từ 1 đến 856)
    ws.row_dimensions[curr_r].height = 20
    ws.cell(row=curr_r, column=1, value=stt_val).alignment = align_center
    ws.cell(row=curr_r, column=2, value=str(ma_dm_val)).alignment = align_center
    ws.cell(row=curr_r, column=3, value='CÔNG TÁC').alignment = align_center
    ws.cell(row=curr_r, column=4, value=str(khu_vuc)).alignment = align_left
    ws.cell(row=curr_r, column=5, value=str(ten_ct)).alignment = align_left
    ws.cell(row=curr_r, column=6, value=str(dvt_val)).alignment = align_center
    
    # Q value
    try:
        q_num = float(q_val)
        c_q = ws.cell(row=curr_r, column=7, value=q_num)
        c_q.alignment = align_right
        c_q.number_format = '#,##0.00' if q_num != int(q_num) else '#,##0'
    except:
        ws.cell(row=curr_r, column=7, value=q_val).alignment = align_right
        
    ws.cell(row=curr_r, column=17, value=str(can_cu_val)).alignment = align_left
    
    for c in range(1, 18):
        cell = ws.cell(row=curr_r, column=c)
        cell.font = font_cong_tac
        cell.fill = fill_cong_tac
        cell.border = box_border
        
    curr_r += 1
    
    # 2. CÁC DÒNG THÀNH PHẦN HAO PHÍ CỦA ĐỊNH MỨC
    components = get_components_for_code(ma_dm_val)
    for comp in components:
        comp_code, comp_type, comp_stage, comp_name, comp_dvt, rate_col_letter, rate_formula_or_val, he_so_val, comp_basis = comp
        
        ws.row_dimensions[curr_r].height = 19
        ws.cell(row=curr_r, column=1, value='').alignment = align_center
        ws.cell(row=curr_r, column=2, value=comp_code).alignment = align_center
        ws.cell(row=curr_r, column=3, value=comp_type).alignment = align_center
        ws.cell(row=curr_r, column=4, value=comp_stage).alignment = align_left
        ws.cell(row=curr_r, column=5, value=comp_name).alignment = align_left
        ws.cell(row=curr_r, column=6, value=comp_dvt).alignment = align_center
        
        # Determine column index for Mức hao phí (H=8, I=9, J=10, K=11)
        # and Khối lượng hao phí (M=13, N=14, O=15, P=16)
        col_map = {'H': (8, 13, 'M'), 'I': (9, 14, 'N'), 'J': (10, 15, 'O'), 'K': (11, 16, 'P')}
        rate_col_idx, qty_col_idx, qty_col_letter = col_map[rate_col_letter]
        
        # Mức hao phí (Link formula hoặc giá trị)
        c_rate = ws.cell(row=curr_r, column=rate_col_idx, value=rate_formula_or_val)
        c_rate.alignment = align_right
        c_rate.number_format = '0.0000'
        
        # Hệ số (Cột L = 12)
        c_hs = ws.cell(row=curr_r, column=12, value=float(he_so_val))
        c_hs.alignment = align_center
        c_hs.number_format = '0.00'
        
        # Khối lượng hao phí (Công thức: = Q * Mức hao phí * Hệ số)
        c_qty = ws.cell(row=curr_r, column=qty_col_idx, value=f"=G${q_row}*{rate_col_letter}{curr_r}*L{curr_r}")
        c_qty.alignment = align_right
        c_qty.number_format = '0.0000'
        
        # Căn cứ pháp lý
        ws.cell(row=curr_r, column=17, value=comp_basis).alignment = align_left
        
        # Borders and fonts
        for c in range(1, 18):
            cell = ws.cell(row=curr_r, column=c)
            cell.border = box_border
            if he_so_val == 0.0:
                cell.font = font_zero
            else:
                cell.font = font_item
                
        curr_r += 1

print(f"Total rows generated: {curr_r - 1}")

# Column widths
ws.column_dimensions['A'].width = 7
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 14
ws.column_dimensions['D'].width = 24
ws.column_dimensions['E'].width = 45
ws.column_dimensions['F'].width = 10
ws.column_dimensions['G'].width = 14
ws.column_dimensions['H'].width = 11
ws.column_dimensions['I'].width = 11
ws.column_dimensions['J'].width = 13
ws.column_dimensions['K'].width = 12
ws.column_dimensions['L'].width = 9
ws.column_dimensions['M'].width = 13
ws.column_dimensions['N'].width = 13
ws.column_dimensions['O'].width = 15
ws.column_dimensions['P'].width = 13
ws.column_dimensions['Q'].width = 35

ws.freeze_panes = 'F7'

print("Saving local...")
wb.save(wb_path)
print(f"Saved: {wb_path}")

print("Saving downloads...")
try:
    wb.save(dl_path)
    print(f"Saved: {dl_path}")
except Exception as e:
    print(f"Downloads lock, fallback: {e}")
    alt_path = 'C:/Users/Admin/Downloads/BANG_PHAN_TICH_PHAN_CAP_856_DICH_VU_PCCC_CapNhat.xlsx'
    wb.save(alt_path)
    print(f"Saved fallback: {alt_path}")
