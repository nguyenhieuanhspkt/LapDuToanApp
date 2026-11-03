# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_th_vat_lieu
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet TH_VatLieu (Bảng Tổng Hợp Vật Liệu chuẩn F1).

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, List, Any
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import FORMAT_CURRENCY, FORMAT_QUANTITY, SHEET_TH_VAT_LIEU


class SheetTHVatLieuBuilder(BaseSheetBuilder):
    """Builder cho Sheet TH_VatLieu theo chuẩn biểu mẫu F1."""

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        materials: List[Dict[str, Any]]
    ) -> None:
        ws = wb.create_sheet(title=SHEET_TH_VAT_LIEU)

        # 1. Cụm Tiêu Đề Biểu Mẫu Chuẩn F1 (14pt Bold Uppercase)
        self.setup_header_banner(ws, "BẢNG TỔNG HỢP VẬT LIỆU", project_name, category_name, max_col=9)
        ws["A2"] = f"CÔNG TRÌNH: {project_name}"
        ws["A3"] = f"HẠNG MỤC: {category_name}"
        ws["I5"] = "Đơn vị: đồng"
        ws["I5"].font = self.font_unit
        ws["I5"].alignment = self.align_right

        # 2. Hàng Tiêu Đề Bảng (Hàng 6)
        headers = ["STT", "Mã hiệu", "Tên vật tư", "Đơn vị", "Nguồn mua", "Khối lượng", "Giá hiện tại", "Thành tiền", "Căn cứ tính giá"]
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

        # 3. Ghi dữ liệu từng dòng vật tư
        for idx, m in enumerate(materials, start=1):
            ws.row_dimensions[curr_r].height = 20
            ws.cell(curr_r, 1, idx).alignment = self.align_center

            c_code = ws.cell(curr_r, 2, m["code"])
            c_code.font = self.font_code
            c_code.alignment = self.align_center
            c_code.fill = self.fill_code

            ws.cell(curr_r, 3, m["name"]).alignment = self.align_left
            ws.cell(curr_r, 4, m["unit"]).alignment = self.align_center
            ws.cell(curr_r, 5, m.get("source", "")).alignment = self.align_left

            is_other = (m.get("code") == "V00750" or m.get("unit") == "%")
            if is_other:
                ws.cell(curr_r, 6, "")
                ws.cell(curr_r, 7, "")
                pct = float(m.get("rate", 2.0))
                c_tt = ws.cell(curr_r, 8, f"=ROUND(SUM(H{r_start}:H{curr_r-1})*{pct}%, 0)")
                c_tt.alignment = self.align_right
                c_tt.number_format = FORMAT_CURRENCY
                c_basis = ws.cell(curr_r, 9, m.get("legal_basis", f"{pct}% vật liệu chính theo định mức"))
                c_basis.alignment = self.align_left
            else:
                c_kl = ws.cell(curr_r, 6, m.get("total_qty", 0.0))
                c_kl.alignment = self.align_right
                c_kl.number_format = FORMAT_QUANTITY

                c_ht = ws.cell(curr_r, 7, m.get("price_notice", 0))
                c_ht.alignment = self.align_right
                c_ht.number_format = FORMAT_CURRENCY

                c_tt = ws.cell(curr_r, 8, f"=ROUND(F{curr_r}*G{curr_r}, 0)")
                c_tt.alignment = self.align_right
                c_tt.number_format = FORMAT_CURRENCY

                basis_val = m.get("legal_basis", m.get("notes", "Công bố giá LS / Báo giá NSX"))
                c_basis = ws.cell(curr_r, 9, basis_val)
                c_basis.alignment = self.align_left

            for c in range(1, 10):
                ws.cell(curr_r, c).border = self.border_cell
                if c not in [1, 2, 8]:
                    ws.cell(curr_r, c).font = self.font_regular
                elif c == 8:
                    ws.cell(curr_r, c).font = self.font_bold

            curr_r += 1

        r_total = curr_r
        ws.row_dimensions[r_total].height = 24
        for c in range(1, 10):
            ws.cell(r_total, c).border = self.border_total

        c_lbl = ws.cell(r_total, 3, "Tổng vật liệu:")
        c_lbl.font = self.font_total
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        c_tot = ws.cell(r_total, 8, f"=SUM(H{r_start}:H{r_total-1})")
        c_tot.font = self.font_total
        c_tot.alignment = self.align_right
        c_tot.number_format = FORMAT_CURRENCY

        # 4. Thêm Khối chữ ký thẩm định 3 bên chuẩn EVN
        self.add_signature_block(ws, start_row=r_total, col_left=2, col_mid=5, col_right=9)

        # 5. Cấu hình Freeze Panes, AutoFilter và Trang in A4 Landscape
        self.setup_table_view(ws, freeze_cell="C7", filter_ref=f"A6:I{r_total-1}")
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="6:6")

        # 6. Thiết lập độ rộng cột chuẩn F1
        col_w = {1: 6, 2: 14, 3: 45, 4: 10, 5: 20, 6: 15, 7: 16, 8: 18, 9: 35}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
