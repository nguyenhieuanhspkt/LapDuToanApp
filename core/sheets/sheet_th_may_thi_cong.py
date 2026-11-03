# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_th_may_thi_cong
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet TH_MayThiCong (Bảng Tổng Hợp Máy Thi Công chuẩn F1).

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, List, Any
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import FORMAT_CURRENCY, FORMAT_QUANTITY, SHEET_TH_MAY_THI_CONG


class SheetTHMayThiCongBuilder(BaseSheetBuilder):
    """Builder cho Sheet TH_MayThiCong theo chuẩn biểu mẫu F1."""

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        machines: List[Dict[str, Any]]
    ) -> None:
        ws = wb.create_sheet(title=SHEET_TH_MAY_THI_CONG)

        # 1. Cụm Tiêu Đề Biểu Mẫu Chuẩn F1 (14pt Bold Uppercase)
        self.setup_header_banner(ws, "BẢNG TỔNG HỢP MÁY THI CÔNG", project_name, category_name, max_col=8)
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["H5"] = "Đơn vị: đồng"
        ws["H5"].font = self.font_unit
        ws["H5"].alignment = self.align_right

        # 2. Hàng Tiêu Đề Bảng (Hàng 6)
        headers = ["STT", "Mã hiệu", "Tên ca máy thi công", "Đơn vị", "Khối lượng", "Giá hiện tại", "Thành tiền", "Căn cứ tính giá"]
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

        # 3. Ghi dữ liệu từng loại máy thi công
        for idx, mac in enumerate(machines, start=1):
            ws.row_dimensions[curr_r].height = 20
            ws.cell(curr_r, 1, idx).alignment = self.align_center

            c_code = ws.cell(curr_r, 2, mac["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code

            ws.cell(curr_r, 3, mac["name"]).alignment = self.align_left
            ws.cell(curr_r, 4, mac.get("unit", "ca")).alignment = self.align_center

            c_kl = ws.cell(curr_r, 5, mac.get("total_qty", 0.0))
            c_kl.alignment = self.align_right
            c_kl.number_format = FORMAT_QUANTITY

            c_ht_val = mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0))
            c_ht = ws.cell(curr_r, 6, c_ht_val)
            c_ht.alignment = self.align_right
            c_ht.number_format = FORMAT_CURRENCY

            c_tt = ws.cell(curr_r, 7, f"=ROUND(E{curr_r}*F{curr_r}, 0)")
            c_tt.alignment = self.align_right
            c_tt.number_format = FORMAT_CURRENCY

            basis_val = mac.get("legal_basis", "Thông tư 37/2026/TT-BXD của Bộ Xây dựng")
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

        c_lbl = ws.cell(r_total, 3, "Cộng máy thi công:")
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
        col_w = {1: 6, 2: 14, 3: 48, 4: 12, 5: 15, 6: 16, 7: 18, 8: 38}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
