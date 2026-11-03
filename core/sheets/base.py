# -*- coding: utf-8 -*-
"""
Module: core.sheets.base
Tầng nền tảng (Base Sheet Builder) chuẩn hóa thể thức trình bày Excel Dự toán theo chuẩn F1:
- Phông chữ chuẩn Times New Roman đồng nhất toàn bộ Workbook
- Tiêu đề biểu mẫu 14pt IN HOA, thông tin công trình/hạng mục nổi bật
- Lưới kỹ thuật (Grid Layout) kẻ khung liền mạch, tối đa mật độ thông tin
- Cố định dòng/cột hiển thị (Freeze Panes) & Bộ lọc tự động (AutoFilter)
- Thiết lập trang in chuẩn A4/A3, lặp dòng tiêu đề khi sang trang
- Khối chữ ký thẩm định 3 bên chuẩn EVN

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, Any, Optional, List
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from core.norm_manager import NormManager


class BaseSheetBuilder:
    """Lớp cơ sở cung cấp hệ thống style và các phương thức tiện ích chuẩn F1."""

    def __init__(self, norm_manager: Optional[NormManager] = None):
        self.norm_manager = norm_manager or NormManager()
        self.init_styles()

    def init_styles(self):
        """Khởi tạo toàn bộ phông chữ, màu nền, đường viền và căn lề chuẩn F1/EVN."""
        # 1. PHÔNG CHỮ CHUẨN TIMES NEW ROMAN
        self.font_title = Font(name="Times New Roman", size=14, bold=True)
        self.font_subtitle = Font(name="Times New Roman", size=11, italic=True)
        self.font_project = Font(name="Times New Roman", size=11, bold=True)
        self.font_unit = Font(name="Times New Roman", size=10, italic=True)
        self.font_header = Font(name="Times New Roman", size=11, bold=True)
        self.font_code = Font(name="Times New Roman", size=11, bold=True)
        self.font_regular = Font(name="Times New Roman", size=11)
        self.font_bold = Font(name="Times New Roman", size=11, bold=True)
        self.font_cat = Font(name="Times New Roman", size=11, bold=True, italic=True, color="002060")
        self.font_blue = Font(name="Times New Roman", size=11, color="002060")
        self.font_total = Font(name="Times New Roman", size=11, bold=True, color="002060")
        self.font_warning = Font(name="Times New Roman", size=11, bold=True, color="C00000")
        self.font_signature_role = Font(name="Times New Roman", size=11, bold=True)
        self.font_signature_note = Font(name="Times New Roman", size=10, italic=True)

        # Phông cho sheet minh chứng tra cứu
        self.font_ev_title = Font(name="Times New Roman", size=11, bold=True, color="002060")
        self.font_ev_header = Font(name="Times New Roman", size=10, bold=True)
        self.font_ev_reg = Font(name="Times New Roman", size=10)
        self.font_ev_bold = Font(name="Times New Roman", size=10, bold=True)
        self.font_ev_cat = Font(name="Times New Roman", size=10, bold=True, italic=True)

        # 2. MÀU NỀN CÓ Ý NGHĨA NGHIỆP VỤ
        # Nền tiêu đề bảng chuẩn F1 (Xanh nhạt #D9E1F2)
        self.fill_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        # Nền mã hiệu công tác (Xanh lam nhạt #BDD7EE)
        self.fill_code = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        # Nền ô nhập liệu / tham chiếu giá (Vàng nhạt #FFF2CC)
        self.fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        # Nền tiêu đề sheet minh chứng
        self.fill_ev_header = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

        # 3. ĐƯỜNG VIỀN KỸ THUẬT (GRID LAYOUT)
        thin_side = Side(border_style="thin", color="A6A6A6")
        thick_top = Side(border_style="medium", color="002060")
        double_bottom = Side(border_style="double", color="002060")

        self.border_cell = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        self.border_subtotal = Border(left=thin_side, right=thin_side, top=thick_top, bottom=thin_side)
        self.border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)

        # 4. CĂN LỀ DỮ LIỆU CHUẨN MỰC
        self.align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
        self.align_left = Alignment(horizontal="left", vertical="center")
        self.align_right = Alignment(horizontal="right", vertical="center")

    def setup_header_banner(
        self,
        ws: Worksheet,
        title: str,
        project_name: str,
        category_name: str,
        max_col: int = 10
    ) -> None:
        """
        Thiết lập cụm thông tin đầu biểu mẫu: Tiêu đề in hoa 14pt, Công trình, Hạng mục.
        Chỉ merge cell ở hàng tiêu đề A1:max_col 1.
        """
        ws.row_dimensions[1].height = 28
        ws.row_dimensions[2].height = 20
        ws.row_dimensions[3].height = 20

        # Tiêu đề biểu mẫu
        ws["A1"] = title.upper()
        ws["A1"].font = self.font_title
        ws["A1"].alignment = self.align_center

        if max_col > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)

        # Tên công trình & Hạng mục
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A2"].font = self.font_project
        ws["A2"].alignment = self.align_left

        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["A3"].font = self.font_project
        ws["A3"].alignment = self.align_left

    def setup_table_view(
        self,
        ws: Worksheet,
        freeze_cell: str = "C6",
        filter_ref: Optional[str] = None
    ) -> None:
        """
        Cấu hình hiển thị bảng tính:
        - Bật lưới kỹ thuật showGridLines
        - Freeze Panes cố định dòng/cột
        - Bật AutoFilter tra cứu dữ liệu
        """
        ws.views.sheetView[0].showGridLines = True
        if freeze_cell:
            ws.freeze_panes = freeze_cell
        if filter_ref:
            ws.auto_filter.ref = filter_ref

    def setup_page_printing(
        self,
        ws: Worksheet,
        orientation: str = "landscape",
        repeat_rows: str = "5:5"
    ) -> None:
        """
        Thiết lập thông số trang in chuẩn phục vụ xuất PDF hoặc in ấn trực tiếp:
        - Khổ giấy A4, vừa khớp chiều ngang trang in (Fit to Width = 1)
        - Cố định lặp lại dòng tiêu đề khi sang trang mới (print_title_rows)
        """
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        if orientation == "portrait":
            ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        else:
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

        if repeat_rows:
            ws.print_title_rows = repeat_rows

        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

    def add_signature_block(
        self,
        ws: Worksheet,
        start_row: int,
        roles: Optional[List[str]] = None,
        date_str: str = "Ngày ..... tháng ..... năm 2026",
        col_left: int = 2,
        col_mid: int = 5,
        col_right: int = 9
    ) -> int:
        """
        Chèn khối chữ ký thẩm định 3 bên chuẩn EVN ở cuối biểu mẫu:
        - Bên trái: Người lập
        - Ở giữa: Người kiểm tra
        - Bên phải: Người chủ trì / Đại diện Chủ đầu tư
        """
        if roles is None:
            roles = ["NGƯỜI LẬP BIỂU", "NGƯỜI KIỂM TRA", "ĐẠI DIỆN CHỦ ĐẦU TƯ / TỔ TRƯỞNG"]

        r_date = start_row + 2
        # Dòng ngày tháng bên phải
        ws.cell(r_date, col_right, date_str).font = self.font_signature_note
        ws.cell(r_date, col_right).alignment = self.align_center

        # Dòng chức danh
        r_role = r_date + 1
        cols = [col_left, col_mid, col_right]
        for idx, role in enumerate(roles):
            c_idx = cols[idx]
            ws.cell(r_role, c_idx, role).font = self.font_signature_role
            ws.cell(r_role, c_idx).alignment = self.align_center

        # Dòng ghi chú ký tên
        r_note = r_role + 1
        for c_idx in cols:
            ws.cell(r_note, c_idx, "(Ký, ghi rõ họ tên)").font = self.font_signature_note
            ws.cell(r_note, c_idx).alignment = self.align_center

        # Khoảng trống ký tên (5 hàng)
        r_sign_space = r_note + 5
        return r_sign_space

    def determine_group(self, item: Dict[str, Any]) -> str:
        """Xác định nhóm công tác chuẩn phục vụ phân nhóm."""
        grp = (item.get("group") or "").strip()
        if grp:
            return grp
        code = (item.get("norm_code") or "").strip()
        if code.startswith("BD.41") or code.startswith("BD.42"):
            return "Hệ thống báo cháy tự động và thiết bị ngoại vi (Lắp đặt - Thay mới)"
        elif code.startswith("E6C.3.01") or code.startswith("E6C.3.02") or code.startswith("E6C.3.03") or code.startswith("E6C.3.04") or code.startswith("E6C.3.05"):
            return "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy"
        elif code.startswith("E6C.3.06") or code.startswith("E6B.") or code.startswith("E6C.1."):
            return "Bảo dưỡng van tràn Deluge, hệ thống ắc quy và thiết bị đo khí"
        return "Các công tác lắp đặt và bảo dưỡng khác"
