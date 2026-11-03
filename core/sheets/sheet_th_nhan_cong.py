# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_th_nhan_cong
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet TH_NhanCong (Bảng Tổng Hợp Nhân Công chuẩn F1).

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, List, Any
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import FORMAT_CURRENCY, FORMAT_QUANTITY, SHEET_TH_NHAN_CONG


class SheetTHNhanCongBuilder(BaseSheetBuilder):
    """Builder cho Sheet TH_NhanCong theo chuẩn biểu mẫu F1."""

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        labor: List[Dict[str, Any]]
    ) -> None:
        ws = wb.create_sheet(title=SHEET_TH_NHAN_CONG)

        # 1. Cụm Tiêu Đề Biểu Mẫu Chuẩn F1 (14pt Bold Uppercase)
        self.setup_header_banner(ws, "BẢNG TỔNG HỢP NHÂN CÔNG", project_name, category_name, max_col=8)
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["H5"] = "Đơn vị: đồng"
        ws["H5"].font = self.font_unit
        ws["H5"].alignment = self.align_right

        # 2. Hàng Tiêu Đề Bảng (Hàng 6)
        headers = ["STT", "Mã hiệu", "Tên nhóm nhân công", "Đơn vị", "Khối lượng", "Giá hiện tại", "Thành tiền", "Căn cứ tính giá"]
        ws.row_dimensions[6].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(6, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_header
            if c_idx == 2:
                cell.fill = self.fill_code

        r_start = 7
        curr_r = r_start

        # 3. Ghi dữ liệu từng nhóm nhân công
        for idx, l in enumerate(labor, start=1):
            ws.row_dimensions[curr_r].height = 20
            ws.cell(curr_r, 1, idx).alignment = self.align_center

            c_code = ws.cell(curr_r, 2, l["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code

            ws.cell(curr_r, 3, l["name"]).alignment = self.align_left
            ws.cell(curr_r, 4, l.get("unit", "công")).alignment = self.align_center

            c_kl = ws.cell(curr_r, 5, l.get("total_qty", 0.0))
            c_kl.alignment = self.align_right
            c_kl.number_format = FORMAT_QUANTITY

            c_ht = ws.cell(curr_r, 6, l.get("price_current", 0))
            c_ht.alignment = self.align_right
            c_ht.number_format = FORMAT_CURRENCY

            c_tt = ws.cell(curr_r, 7, f"=ROUND(E{curr_r}*F{curr_r}, 0)")
            c_tt.alignment = self.align_right
            c_tt.number_format = FORMAT_CURRENCY

            basis_val = l.get("legal_basis", "Văn bản hướng dẫn đơn giá nhân công tỉnh Lâm Đồng")
            c_basis = ws.cell(curr_r, 8, basis_val)
            c_basis.alignment = self.align_left

            for c in range(1, 9):
                ws.cell(curr_r, c).border = self.border_cell
                if c not in [1, 2, 7]:
                    ws.cell(curr_r, c).font = self.font_regular
                elif c == 7:
                    ws.cell(curr_r, c).font = self.font_bold

            curr_r += 1

        r_total = curr_r
        ws.row_dimensions[r_total].height = 24
        for c in range(1, 9):
            ws.cell(r_total, c).border = self.border_total

        c_lbl = ws.cell(r_total, 3, "Cộng nhân công:")
        c_lbl.font = self.font_total
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        c_tot = ws.cell(r_total, 7, f"=SUM(G{r_start}:G{r_total-1})")
        c_tot.font = self.font_total
        c_tot.alignment = self.align_right
        c_tot.number_format = FORMAT_CURRENCY

        # 4. Thêm Khối chữ ký thẩm định 3 bên chuẩn EVN
        self.add_signature_block(ws, start_row=r_total, col_left=2, col_mid=4, col_right=8)

        # 5. Cấu hình Freeze Panes, AutoFilter và Trang in A4 Landscape
        self.setup_table_view(ws, freeze_cell="C7", filter_ref=f"A6:H{r_total-1}")
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="6:6")

        # 6. Thiết lập độ rộng cột chuẩn F1
        col_w = {1: 6, 2: 14, 3: 45, 4: 12, 5: 15, 6: 16, 7: 18, 8: 38}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
