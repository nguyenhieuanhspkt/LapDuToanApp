# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_thong_ke_dinh_muc
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet 00_ThongKe_DinhMuc (Thống kê định mức chuẩn F1).

Tác giả: Tổ Thẩm Định EVN
"""

from typing import List
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import SHEET_THONG_KE_DINH_MUC


class SheetThongKeDinhMucBuilder(BaseSheetBuilder):
    """Builder cho Sheet 00_ThongKe_DinhMuc theo chuẩn biểu mẫu F1."""

    def build(self, wb: openpyxl.Workbook, norm_codes: List[str]) -> None:
        ws = wb.create_sheet(title=SHEET_THONG_KE_DINH_MUC)

        ws["A1"] = "BẢNG THỐNG KÊ TOÀN BỘ ĐỊNH MỨC ÁP DỤNG TRONG DỰ TOÁN PCCC 2026"
        ws["A1"].font = self.font_title
        ws["A2"] = "(Quy chuẩn phân loại theo ĐM SCL EVN QĐ 45/QĐ-HĐTV và ĐM BXD TT 38/2026/TT-BXD)"
        ws["A2"].font = self.font_subtitle

        headers = ["STT", "Bộ định mức", "Mã hiệu chuẩn", "Tên công tác chuẩn", "ĐVT gốc", "Căn cứ pháp lý", "Link sang sheet minh chứng"]
        ws.row_dimensions[4].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(4, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_header
            if c_idx == 3:
                cell.fill = self.fill_code

        for idx, code in enumerate(norm_codes, start=1):
            norm = self.norm_manager.get_norm(code) or {}
            r = 4 + idx
            ws.row_dimensions[r].height = 22
            ws.cell(r, 1, idx).alignment = self.align_center

            bo_dm = "ĐM SCL EVN (QĐ 45)" if code.startswith("E") else "ĐM BXD (TT 38/2026)"
            ws.cell(r, 2, bo_dm).alignment = self.align_center

            c_code = ws.cell(r, 3, code)
            c_code.alignment = self.align_center
            c_code.font = self.font_code
            c_code.fill = self.fill_code

            ws.cell(r, 4, norm.get("name", "")).alignment = self.align_left
            ws.cell(r, 5, norm.get("unit", "")).alignment = self.align_center
            ws.cell(r, 6, norm.get("legal_basis", "")).alignment = self.align_left

            # Hyperlink to evidence sheet
            c_link = ws.cell(r, 7, f"Xem sheet {code}")
            c_link.hyperlink = f"#'{code}'!A1"
            c_link.font = Font(name="Times New Roman", size=11, color="0563C1", underline="single")
            c_link.alignment = self.align_center

            for c in range(1, 8):
                ws.cell(r, c).border = self.border_cell

        # Freeze Panes, AutoFilter và Page Setup
        last_r = 4 + len(norm_codes)
        self.setup_table_view(ws, freeze_cell="C5", filter_ref=f"A4:G{last_r}")
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="4:4")

        col_w = {1: 6, 2: 24, 3: 16, 4: 55, 5: 12, 6: 45, 7: 25}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
