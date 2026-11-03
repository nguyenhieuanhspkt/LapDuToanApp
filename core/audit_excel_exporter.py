"""
Excel Exporter for Semantic Task Auditor
Xuất Báo cáo Thẩm tra & Bóc tách Dữ liệu Thô PAKT (.xlsx)
Author: Tổ Thẩm Định EVN
"""

import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import Dict, Any, List

class AuditExcelExporter:
    """Trình sinh file Excel Báo cáo Thẩm tra Dữ liệu Thô PAKT chuẩn mực."""

    def __init__(self):
        # Palette màu chuẩn EVN & Thẩm định
        self.color_navy = "002060"
        self.color_header_fill = "002060"
        self.color_header_font = "FFFFFF"
        
        # Màu trạng thái cảnh báo
        self.fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        self.font_green = Font(name="Times New Roman", size=10, bold=True, color="375623")

        self.fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        self.font_yellow = Font(name="Times New Roman", size=10, bold=True, color="9C6500")

        self.fill_red = PatternFill(start_color="FFD2D2", end_color="FFD2D2", fill_type="solid")
        self.font_red = Font(name="Times New Roman", size=10, bold=True, color="9C0006")

        # Viền ô
        thin = Side(border_style="thin", color="D9D9D9")
        thick_bottom = Side(border_style="medium", color="002060")
        self.border_cell = Border(left=thin, right=thin, top=thin, bottom=thin)
        self.border_header = Border(left=thin, right=thin, top=thin, bottom=thick_bottom)

    def export_workbook(self, audit_result: Dict[str, Any], project_title: str = "HỒ SƠ PAKT VĨNH TÂN 4") -> bytes:
        """Tạo workbook và trả về byte stream."""
        wb = openpyxl.Workbook()
        ws_default = wb.active
        wb.remove(ws_default)

        # 1. Sheet 1: Bảng Thẩm tra V-O-S-C
        self._build_sheet_vosc(wb, audit_result["audited_rows"], project_title)

        # 2. Sheet 2: Bảng Thống kê Cụm mẫu lặp
        self._build_sheet_clusters(wb, audit_result["summary_stats"]["clusters"], project_title)

        # 3. Sheet 3: Bảng Tổng hợp Trạng thái Hồ sơ
        self._build_sheet_summary(wb, audit_result["summary_stats"], project_title)

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def _build_sheet_vosc(self, wb: openpyxl.Workbook, rows: List[Dict[str, Any]], project_title: str):
        ws = wb.create_sheet(title="Bang_ThamTra_VOSC")
        ws.views.sheetView[0].showGridLines = True

        # Tiêu đề báo cáo
        ws.merge_cells("A1:W1")
        ws["A1"] = "BÁO CÁO THẨM TRA NGỮ NGHĨA & BÓC TÁCH BẢN CHẤT KỸ THUẬT DỮ LIỆU THÔ PAKT"
        ws["A1"].font = Font(name="Times New Roman", size=15, bold=True, color=self.color_navy)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        ws.merge_cells("A2:W2")
        ws["A2"] = f"Dự án / Công trình: {project_title.upper()} | Trích xuất phân rã V-O-S-C & Đối chiếu Nội hàm Định mức Gốc"
        ws["A2"].font = Font(name="Times New Roman", size=11, italic=True, color="555555")
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Header Columns
        headers = [
            ("STT thô", 8),
            ("Tên công việc thô", 35),
            ("Mô tả thô", 30),
            ("ĐVT thô", 10),
            ("KL thô", 10),
            # 4 Cột VOSC độc lập
            ("Hành động (V)", 18),
            ("Đối tượng (O)", 22),
            ("Quy cách (S)", 22),
            ("Bối cảnh (C)", 22),
            # Cảnh báo trọng tâm
            ("CẦN LÀM RÕ GÌ? (HÀNH ĐỘNG THẨM TRA)", 45),
            ("Trạng thái", 16),
            # Bản chất
            ("Làm cái gì? (What)", 32),
            ("Làm như thế nào? (How)", 38),
            # Định mức đề xuất & Bạch hóa
            ("Mã ĐM đề xuất", 14),
            ("Tên định mức gốc", 30),
            ("ĐVT gốc", 10),
            ("Căn cứ pháp lý", 25),
            ("Thành phần công việc gốc (Các bước)", 40),
            ("Hao phí nhân công gốc", 35),
            ("Hao phí máy thi công gốc", 25),
            # Quy đổi
            ("Hệ số quy đổi", 12),
            ("KL quy đổi", 12),
            # Cụm
            ("Cụm mẫu lặp", 28),
            ("Số dòng cùng cụm", 14)
        ]

        header_row = 4
        ws.row_dimensions[header_row].height = 32

        for col_idx, (col_name, width) in enumerate(headers, start=1):
            cell = ws.cell(row=header_row, column=col_idx, value=col_name)
            cell.font = Font(name="Times New Roman", size=10, bold=True, color=self.color_header_font)
            cell.fill = PatternFill(start_color=self.color_header_fill, end_color=self.color_header_fill, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = self.border_header
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # Data Rows
        current_row = 5
        font_data = Font(name="Times New Roman", size=10)
        align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
        align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        align_right = Alignment(horizontal="right", vertical="center")

        for r in rows:
            ws.row_dimensions[current_row].height = 42

            ws.cell(row=current_row, column=1, value=r["stt"]).alignment = align_center
            ws.cell(row=current_row, column=2, value=r["raw_job_name"]).alignment = align_left
            ws.cell(row=current_row, column=3, value=r["raw_desc"]).alignment = align_left
            ws.cell(row=current_row, column=4, value=r["raw_unit"]).alignment = align_center
            c_vol = ws.cell(row=current_row, column=5, value=r["raw_volume"])
            c_vol.alignment = align_right
            c_vol.number_format = "#,##0.00"

            # 4 Cột VOSC
            ws.cell(row=current_row, column=6, value=r["v_action"]).alignment = align_left
            ws.cell(row=current_row, column=7, value=r["o_object"]).alignment = align_left
            ws.cell(row=current_row, column=8, value=r["s_spec"]).alignment = align_left
            ws.cell(row=current_row, column=9, value=r["c_context"]).alignment = align_left

            # Cột CẦN LÀM RÕ GÌ? (Tô màu)
            c_clarify = ws.cell(row=current_row, column=10, value=r["can_lam_ro"])
            c_clarify.alignment = align_left

            c_status = ws.cell(row=current_row, column=11, value=r["audit_status"])
            c_status.alignment = align_center

            # Áp màu theo level
            level = r["audit_level"]
            if level == "RED":
                c_clarify.fill = self.fill_red
                c_clarify.font = self.font_red
                c_status.fill = self.fill_red
                c_status.font = self.font_red
            elif level == "YELLOW":
                c_clarify.fill = self.fill_yellow
                c_clarify.font = self.font_yellow
                c_status.fill = self.fill_yellow
                c_status.font = self.font_yellow
            else:
                c_clarify.fill = self.fill_green
                c_clarify.font = self.font_green
                c_status.fill = self.fill_green
                c_status.font = self.font_green

            # Bản chất
            ws.cell(row=current_row, column=12, value=r["nature_what"]).alignment = align_left
            ws.cell(row=current_row, column=13, value=r["nature_how"]).alignment = align_left

            # Nội hàm định mức
            ws.cell(row=current_row, column=14, value=r["prop_norm_code"]).alignment = align_center
            ws.cell(row=current_row, column=15, value=r["prop_norm_name"]).alignment = align_left
            ws.cell(row=current_row, column=16, value=r["prop_norm_unit"]).alignment = align_center
            ws.cell(row=current_row, column=17, value=r["prop_legal_basis"]).alignment = align_left
            ws.cell(row=current_row, column=18, value=r["prop_work_steps"]).alignment = align_left
            ws.cell(row=current_row, column=19, value=r["prop_labor_rate"]).alignment = align_left
            ws.cell(row=current_row, column=20, value=r["prop_machine_rate"]).alignment = align_left

            # Quy đổi
            c_conv = ws.cell(row=current_row, column=21, value=r["conv_factor"])
            c_conv.alignment = align_right
            c_conv.number_format = "#,##0.0000"

            c_cvol = ws.cell(row=current_row, column=22, value=r["converted_vol"])
            c_cvol.alignment = align_right
            c_cvol.number_format = "#,##0.0000"

            # Cụm
            ws.cell(row=current_row, column=23, value=r["cluster_id"]).alignment = align_left
            ws.cell(row=current_row, column=24, value=r["cluster_count"]).alignment = align_center

            # Format font & border
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=current_row, column=col_idx)
                if col_idx not in [10, 11]:  # Đã format màu riêng ở trên
                    cell.font = font_data
                cell.border = self.border_cell

            current_row += 1

        # BẬT CHẾ ĐỘ AUTOFILTER TOÀN BỘ CỘT
        last_col_letter = get_column_letter(len(headers))
        ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{current_row - 1}"
        ws.freeze_panes = "F5"  # Cố định từ cột V (Hành động) trở đi

    def _build_sheet_clusters(self, wb: openpyxl.Workbook, clusters: List[Dict[str, Any]], project_title: str):
        ws = wb.create_sheet(title="Cum_Mau_LapLai")
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:E1")
        ws["A1"] = "BẢNG GOM CỤM MẪU CÔNG VIỆC LẶP LẠI TRONG CHUỖI DỮ LIỆU THÔ"
        ws["A1"].font = Font(name="Times New Roman", size=14, bold=True, color=self.color_navy)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        headers = [
            ("STT Cụm", 10),
            ("Tên Cụm Mẫu Lặp (Hành Động V + Đối Tượng O)", 45),
            ("Số Dòng Xuất Hiện (Tần Suất)", 20),
            ("Tỷ Lệ %", 14),
            ("Khuyến Nghị Xử Lý Đồng Loạt", 45)
        ]

        h_row = 3
        ws.row_dimensions[h_row].height = 26
        for col_idx, (col_name, width) in enumerate(headers, start=1):
            cell = ws.cell(row=h_row, column=col_idx, value=col_name)
            cell.font = Font(name="Times New Roman", size=10, bold=True, color=self.color_header_font)
            cell.fill = PatternFill(start_color=self.color_header_fill, end_color=self.color_header_fill, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.border_header
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        curr = 4
        font_data = Font(name="Times New Roman", size=10)
        for idx, cl in enumerate(clusters, start=1):
            ws.row_dimensions[curr].height = 25
            c_name = cl["cluster_name"]
            
            # Gợi ý xử lý
            if "Bảo dưỡng" in c_name:
                sol = "Áp dụng đồng loạt mã E6C hoặc Biến thể A1 loại trừ vật tư mua mới."
            elif "Lắp mới" in c_name:
                sol = "Áp dụng định mức xây lắp TT38/2026 nguyên bản (Biến thể A0)."
            elif "Chưa xác định" in c_name or "Khái quát" in c_name:
                sol = "Yêu cầu đơn vị lập PAKT giải trình, bóc tách lại; không áp ĐM đơn chiếc."
            else:
                sol = "Kiểm tra đối chiếu hồ sơ thiết kế và áp mã theo quy trình chuẩn."

            ws.cell(row=curr, column=1, value=idx).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=curr, column=2, value=c_name).alignment = Alignment(horizontal="left", vertical="center")
            
            c_cnt = ws.cell(row=curr, column=3, value=cl["count"])
            c_cnt.alignment = Alignment(horizontal="center", vertical="center")
            c_cnt.number_format = "#,##0"

            c_pct = ws.cell(row=curr, column=4, value=cl["pct"] / 100.0)
            c_pct.alignment = Alignment(horizontal="right", vertical="center")
            c_pct.number_format = "0.0%"

            ws.cell(row=curr, column=5, value=sol).alignment = Alignment(horizontal="left", vertical="center")

            for col_idx in range(1, 6):
                c = ws.cell(row=curr, column=col_idx)
                c.font = font_data
                c.border = self.border_cell

            curr += 1

        last_col = get_column_letter(5)
        ws.auto_filter.ref = f"A{h_row}:{last_col}{curr - 1}"

    def _build_sheet_summary(self, wb: openpyxl.Workbook, stats: Dict[str, Any], project_title: str):
        ws = wb.create_sheet(title="TongHop_TrangThai_HoSo")
        ws.views.sheetView[0].showGridLines = True

        ws.merge_cells("A1:D1")
        ws["A1"] = "BẢNG ĐÁNH GIÁ CHẤT LƯỢNG HỒ SƠ DỮ LIỆU THÔ PAKT"
        ws["A1"].font = Font(name="Times New Roman", size=14, bold=True, color=self.color_navy)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 25

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 16

        headers = ["STT", "Hạng Mục Đánh Giá", "Số Lượng Dòng", "Tỷ Lệ %"]
        h_row = 3
        for col_idx, h in enumerate(headers, start=1):
            cell = ws.cell(row=h_row, column=col_idx, value=h)
            cell.font = Font(name="Times New Roman", size=10, bold=True, color=self.color_header_font)
            cell.fill = PatternFill(start_color=self.color_header_fill, end_color=self.color_header_fill, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.border_header

        tot = stats["total_items"]
        g_cnt = stats["green_count"]
        y_cnt = stats["yellow_count"]
        r_cnt = stats["red_count"]

        rows_meta = [
            (1, "🟢 Hồ sơ rõ ràng (Đủ điều kiện áp định mức)", g_cnt, self.fill_green, self.font_green),
            (2, "🟡 Hồ sơ cần làm rõ điều kiện / Biến thể A1 / Quy đổi", y_cnt, self.fill_yellow, self.font_yellow),
            (3, "🔴 Hồ sơ mập mờ / Chưa đủ điều kiện (Cấm ép định mức)", r_cnt, self.fill_red, self.font_red)
        ]

        curr = 4
        for stt, name, cnt, fill, font in rows_meta:
            ws.row_dimensions[curr].height = 24
            ws.cell(row=curr, column=1, value=stt).alignment = Alignment(horizontal="center", vertical="center")
            
            c_name = ws.cell(row=curr, column=2, value=name)
            c_name.alignment = Alignment(horizontal="left", vertical="center")
            c_name.fill = fill
            c_name.font = font

            c_val = ws.cell(row=curr, column=3, value=cnt)
            c_val.alignment = Alignment(horizontal="center", vertical="center")
            c_val.number_format = "#,##0"

            c_pct = ws.cell(row=curr, column=4, value=(cnt / tot) if tot else 0)
            c_pct.alignment = Alignment(horizontal="right", vertical="center")
            c_pct.number_format = "0.0%"

            for col_idx in range(1, 5):
                cell = ws.cell(row=curr, column=col_idx)
                if col_idx != 2:
                    cell.font = Font(name="Times New Roman", size=10)
                cell.border = self.border_cell

            curr += 1

        # Dòng tổng cộng
        ws.row_dimensions[curr].height = 25
        ws.cell(row=curr, column=1, value="").border = self.border_cell
        c_tot_label = ws.cell(row=curr, column=2, value="TỔNG CỘNG SỐ DÒNG CÔNG VIỆC THÔ")
        c_tot_label.font = Font(name="Times New Roman", size=10, bold=True)
        c_tot_label.border = self.border_cell

        c_tot_val = ws.cell(row=curr, column=3, value=tot)
        c_tot_val.font = Font(name="Times New Roman", size=10, bold=True)
        c_tot_val.alignment = Alignment(horizontal="center", vertical="center")
        c_tot_val.border = self.border_cell

        c_tot_pct = ws.cell(row=curr, column=4, value=1.0)
        c_tot_pct.font = Font(name="Times New Roman", size=10, bold=True)
        c_tot_pct.alignment = Alignment(horizontal="right", vertical="center")
        c_tot_pct.number_format = "0.0%"
        c_tot_pct.border = self.border_cell
