# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_th_kinh_phi
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet TH_KinhPhi:
- Bảng tổng hợp dự toán chi phí xây dựng theo chuẩn F1 & Thông tư 11/2021/TT-BXD, Nghị định 72/2024/NĐ-CP
- Liên kết 100% công thức động trực tiếp đến bảng khối lượng và các bảng tổng hợp tài nguyên
- Tự động chuyển đổi số tiền thành chữ tiếng Việt (num2words_vi)
- Khối chữ ký thẩm định 3 bên chuẩn EVN

Tác giả: Tổ Thẩm Định EVN
"""

from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import FORMAT_CURRENCY, SHEET_TH_KINH_PHI
from utils.num2words_vi import num2words_vi


class SheetTHKinhPhiBuilder(BaseSheetBuilder):
    """Builder xây dựng bảng Tổng Hợp Kinh Phí Dự Toán Chi Phí Xây Dựng."""

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        coefficients: Dict[str, Any],
        r_grand_total_kl: Optional[int] = None,
        r_total_vl: int = 7,
        r_total_nc: int = 7,
        r_total_mac: int = 7,
        materials: Optional[List[Dict[str, Any]]] = None,
        labor: Optional[List[Dict[str, Any]]] = None,
        machines: Optional[List[Dict[str, Any]]] = None,
        quote_sum: float = 0.0,
        quote_count: int = 0
    ) -> None:
        ws = wb.create_sheet(title=SHEET_TH_KINH_PHI)

        # 1. Cụm Tiêu Đề Biểu Mẫu Chuẩn F1 (A1:F1 merge, 14pt Bold Uppercase)
        self.setup_header_banner(ws, "BẢNG TỔNG HỢP DỰ TOÁN CHI PHÍ XÂY DỰNG", project_name, category_name, max_col=6)
        ws["F4"] = "Đơn vị: đồng"
        ws["F4"].font = self.font_unit
        ws["F4"].alignment = self.align_right

        # 2. Hàng Tiêu Đề Bảng (Hàng 5)
        headers = ["STT", "NỘI DUNG CHI PHÍ", "CÁCH TÍNH", "GIÁ TRỊ", "KÝ HIỆU", "CƠ SỞ HỆ SỐ / GHI CHÚ"]
        ws.row_dimensions[5].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, c_idx, h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = self.align_center
            cell.border = self.border_cell

        # 3. Thông số Hệ số Chi phí theo quy định hiện hành
        c_rate = float(coefficients.get("c", {}).get("rate", 7.3))
        c_basis = coefficients.get("c", {}).get("basis", "Thông tư 11/2021/TT-BXD")
        tt_rate = float(coefficients.get("tt", {}).get("rate", 2.5))
        tt_basis = coefficients.get("tt", {}).get("basis", "Thông tư 11/2021/TT-BXD")
        tl_rate = float(coefficients.get("tl", {}).get("rate", 5.5))
        tl_basis = coefficients.get("tl", {}).get("basis", "Thông tư 11/2021/TT-BXD")
        vat_rate = float(coefficients.get("vat", {}).get("rate", 8.0))
        vat_basis = coefficients.get("vat", {}).get("basis", "Nghị định 72/2024/NĐ-CP")
        lt_rate = float(coefficients.get("lt", {}).get("rate", 1.1))
        lt_basis = coefficients.get("lt", {}).get("basis", "Thông tư 11/2021/TT-BXD")

        # 4. Tham chiếu nguồn giá
        if r_grand_total_kl:
            ref_vl = f"='Khối lượng-Định mức-ĐVT'!$K${r_grand_total_kl}"
            ref_nc = f"='Khối lượng-Định mức-ĐVT'!$L${r_grand_total_kl}"
            ref_mac = f"='Khối lượng-Định mức-ĐVT'!$M${r_grand_total_kl}"
            ref_bg = f"='Khối lượng-Định mức-ĐVT'!$P${r_grand_total_kl}"
            note_src = "Bảng tiên lượng khối lượng"
        else:
            ref_vl = f"='TH_VatLieu'!H{r_total_vl}"
            ref_nc = f"='TH_NhanCong'!G{r_total_nc}"
            ref_mac = f"='TH_MayThiCong'!G{r_total_mac}"
            ref_bg = f"{quote_sum:.2f}"
            note_src = "Bảng tổng hợp tài nguyên"

        # 5. Xây dựng cấu trúc các dòng chi phí (Chuẩn F1)
        rows_spec = []
        r_T = 6
        if quote_sum > 0:
            formula_T = "=D7+D9+D11+D13"
            calc_T = "T = VL + NC + M + BG"
            r_GT = 15
            r_C = 16
            r_TT = 17
            r_TL = 18
            r_G = 19
            r_VAT = 20
            r_GXD = 21
            r_LT = 22
            r_TOT = 23

            rows_spec.append(("I", "Chi phí trực tiếp", formula_T, formula_T, "T", calc_T, True, True))
            rows_spec.append(("1", "Chi phí vật liệu", "=D8", "=D8", "VL", "VL = VLHT", True, False))
            rows_spec.append(("", " - Đơn giá vật liệu", f"Theo {note_src}", ref_vl, "VLHT", note_src, False, False))
            rows_spec.append(("2", "Chi phí nhân công", "=D10", "=D10", "NC", "NC = NCHT", True, False))
            rows_spec.append(("", " - Đơn giá nhân công", f"Theo {note_src}", ref_nc, "NCHT", note_src, False, False))
            rows_spec.append(("3", "Chi phí máy và thiết bị thi công", "=D12", "=D12", "M", "M = MHT", True, False))
            rows_spec.append(("", " - Đơn giá máy thi công", f"Theo {note_src}", ref_mac, "MHT", note_src, False, False))
            rows_spec.append(("4", "Chi phí công tác báo giá / tạm tính", "=D14", "=D14", "BG", f"{quote_count} công việc báo giá trọn gói", True, False))
            rows_spec.append(("", " - Chi phí báo giá trước thuế", "Theo báo giá cạnh tranh / PAKT", ref_bg, "BGHT", "Hồ sơ PAKT & Báo giá", False, False))
        else:
            formula_T = "=D7+D9+D11"
            calc_T = "T = VL + NC + M"
            r_GT = 13
            r_C = 14
            r_TT = 15
            r_TL = 16
            r_G = 17
            r_VAT = 18
            r_GXD = 19
            r_LT = 20
            r_TOT = 21

            rows_spec.append(("I", "Chi phí trực tiếp", formula_T, formula_T, "T", calc_T, True, True))
            rows_spec.append(("1", "Chi phí vật liệu", "=D8", "=D8", "VL", "VL = VLHT", True, False))
            rows_spec.append(("", " - Đơn giá vật liệu", f"Theo {note_src}", ref_vl, "VLHT", note_src, False, False))
            rows_spec.append(("2", "Chi phí nhân công", "=D10", "=D10", "NC", "NC = NCHT", True, False))
            rows_spec.append(("", " - Đơn giá nhân công", f"Theo {note_src}", ref_nc, "NCHT", note_src, False, False))
            rows_spec.append(("3", "Chi phí máy và thiết bị thi công", "=D12", "=D12", "M", "M = MHT", True, False))
            rows_spec.append(("", " - Đơn giá máy thi công", f"Theo {note_src}", ref_mac, "MHT", note_src, False, False))

        rows_spec.extend([
            ("II", "Chi phí gián tiếp", f"=D{r_C}+D{r_TT}", f"=D{r_C}+D{r_TT}", "GT", "GT = C + TT", True, True),
            ("1", "Chi phí chung", f"=D{r_T}*{c_rate}%", f"=D{r_T}*{c_rate/100:.6f}", "C", c_basis, False, False),
            ("2", "Chi phí một số công việc không xác định được KL từ thiết kế", f"=D{r_T}*{tt_rate}%", f"=D{r_T}*{tt_rate/100:.6f}", "TT", tt_basis, False, False),
            ("III", "Thu nhập chịu thuế tính trước", f"=(D{r_T}+D{r_GT})*{tl_rate}%", f"=(D{r_T}+D{r_GT})*{tl_rate/100:.6f}", "TL", tl_basis, True, True),
            ("", "Chi phí xây dựng trước thuế", f"=D{r_T}+D{r_GT}+D{r_TL}", f"=D{r_T}+D{r_GT}+D{r_TL}", "G", "G = T + GT + TL", True, False),
            ("IV", "Thuế giá trị gia tăng", f"=D{r_G}*{vat_rate}%", f"=D{r_G}*{vat_rate/100:.6f}", "GTGT", vat_basis, True, True),
            ("", "CHI PHÍ XÂY DỰNG SAU THUẾ", f"=D{r_G}+D{r_VAT}", f"=D{r_G}+D{r_VAT}", "Gxd", "Gxd = G + GTGT", True, True),
            ("V", "Chi phí nhà tạm để ở và điều hành thi công", f"=D{r_G}*{lt_rate}%*(1+{vat_rate}%)", f"=D{r_G}*{lt_rate/100:.6f}*(1+{vat_rate/100:.6f})", "LT", lt_basis, True, False),
            ("", "Tổng cộng", f"=D{r_GXD}+D{r_LT}", f"=D{r_GXD}+D{r_LT}", "", "", True, True),
            ("", "Làm tròn", f"=ROUND(D{r_TOT},-3)", f"=ROUND(D{r_TOT},-3)", "", "", True, True)
        ])

        curr_r = 6
        for stt, name, calc, formula, sym, note, is_b, is_hl in rows_spec:
            ws.row_dimensions[curr_r].height = 22
            c_stt = ws.cell(curr_r, 1, stt)
            c_stt.alignment = self.align_center
            if is_b:
                c_stt.font = self.font_bold

            c_name = ws.cell(curr_r, 2, name)
            c_name.alignment = self.align_left
            c_name.font = self.font_bold if is_b else self.font_regular

            c_calc = ws.cell(curr_r, 3, calc)
            c_calc.alignment = self.align_left
            c_calc.font = self.font_regular

            c_val = ws.cell(curr_r, 4, formula)
            c_val.alignment = self.align_right
            c_val.font = self.font_bold if is_b else self.font_regular
            c_val.number_format = FORMAT_CURRENCY

            c_sym = ws.cell(curr_r, 5, sym)
            c_sym.alignment = self.align_center
            c_sym.font = self.font_bold if sym else self.font_regular

            c_note = ws.cell(curr_r, 6, note)
            c_note.alignment = self.align_left
            c_note.font = self.font_regular

            for c in range(1, 7):
                ws.cell(curr_r, c).border = self.border_cell
                if is_hl and stt in ["I", "II", "III", "IV", "V"]:
                    ws.cell(curr_r, c).fill = self.fill_yellow
                elif name in ["Tổng cộng", "Làm tròn", "CHI PHÍ XÂY DỰNG SAU THUẾ"]:
                    ws.cell(curr_r, c).fill = self.fill_code

            curr_r += 1

        # 6. Dòng Số tiền bằng chữ
        m_list = materials or []
        l_list = labor or []
        mac_list = machines or []
        sum_vl_chinh = sum(m.get("total_qty", 0.0) * m.get("price_notice", 0) for m in m_list if m.get("code") != "V00750" and m.get("unit") != "%")
        sum_vl = sum_vl_chinh * 1.02
        sum_nc = sum(l.get("total_qty", 0.0) * l.get("price_current", 0) for l in l_list)
        sum_mac = sum(mac.get("total_qty", 0.0) * mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0)) for mac in mac_list)
        t_est = sum_vl + sum_nc + sum_mac + quote_sum
        c_est = t_est * (c_rate / 100.0)
        tt_est = t_est * (tt_rate / 100.0)
        gt_est = c_est + tt_est
        tl_est = (t_est + gt_est) * (tl_rate / 100.0)
        g_est = t_est + gt_est + tl_est
        vat_est = g_est * (vat_rate / 100.0)
        gxd_est = g_est + vat_est
        lt_est = g_est * (lt_rate / 100.0) * (1.0 + vat_rate / 100.0)
        tot_est = gxd_est + lt_est
        round_est = round(tot_est, -3)
        words_vi = num2words_vi(round_est)

        ws.row_dimensions[curr_r].height = 24
        ws.merge_cells(start_row=curr_r, start_column=2, end_row=curr_r, end_column=6)
        c_words = ws.cell(curr_r, 2, f"Bằng chữ: {words_vi}")
        c_words.font = Font(name="Times New Roman", size=11, bold=True, italic=True, color="002060")
        c_words.alignment = self.align_left

        # 7. Khối Chữ Ký Chuẩn 3 Bên EVN
        self.add_signature_block(
            ws=ws,
            start_row=curr_r,
            roles=["NGƯỜI LẬP", "NGƯỜI KIỂM TRA", "NGƯỜI CHỦ TRÌ (Kỹ sư định giá)"],
            col_left=2,
            col_mid=3,
            col_right=6
        )

        # 8. Cấu hình Freeze Panes, AutoFilter và Thiết lập in ấn A4 Portrait
        self.setup_table_view(ws, freeze_cell="C6", filter_ref=None)
        self.setup_page_printing(ws, orientation="portrait", repeat_rows="5:5")

        # 9. Độ rộng cột chuẩn
        col_w = {1: 8, 2: 45, 3: 28, 4: 18, 5: 12, 6: 42}
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w
