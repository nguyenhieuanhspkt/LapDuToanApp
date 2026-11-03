# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_don_gia_chi_tiet
Tầng xuất dữ liệu Excel (Presentation Layer) cho Sheet 01_ChiTiet_HaoPhi_DinhMuc.

Tuân thủ nghiêm ngặt chỉ đạo kiến trúc và tối ưu hóa mã nguồn:
1. Nguyên tắc Đơn nhiệm (Single Responsibility):
   - Tách biệt rõ ràng từng khâu: Khởi tạo Header, Ghi công tác Báo giá, Ghi Đầu mục công việc,
     và Ghi các phân đoạn tài nguyên (_write_materials, _write_labors, _write_machines).
2. Mẫu thiết kế Chiến lược (Strategy Pattern):
   - Hợp nhất quy trình ghi dữ liệu Vật liệu, Nhân công, Máy thi công qua cấu hình ResourceCategoryConfig,
     loại bỏ hoàn toàn sự trùng lặp mã nguồn giữa các nhóm tài nguyên.
3. Tập trung hằng số & Công thức động:
   - Toàn bộ tên sheet, tên cột, mã đặc thù được import từ core.constants.
   - Toàn bộ công thức Excel được sinh từ core.formulas.ExcelFormulaBuilder.
4. Tầng nghiệp vụ độc lập:
   - Toàn bộ biểu thức điều kiện phức tạp được chuẩn hóa qua các hàm nghiệp vụ có tên tường minh
     trong core.business (is_quote_item, is_replacement_material, is_percentage_material, should_link_evidence_sheet).

Tác giả: Tổ Thẩm Định EVN
"""

from dataclasses import dataclass
from typing import Dict, List, Any, Optional, Tuple
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from core.sheets.base import BaseSheetBuilder
from core.constants import (
    SHEET_CHI_TIET_HAO_PHI,
    SHEET_TH_VAT_LIEU,
    SHEET_TH_NHAN_CONG,
    SHEET_TH_MAY_THI_CONG,
    FORMAT_CURRENCY,
    FORMAT_RATE,
    FORMAT_QUANTITY,
    COL_STT,
    COL_CODE,
    COL_NAME,
    COL_UNIT,
    COL_RATE,
    COL_PRICE,
    COL_AMOUNT,
    LET_COL_RATE,
    LET_COL_AMOUNT,
    LET_COL_PRICE_TH_VAT_LIEU,
    LET_COL_PRICE_TH_NHAN_CONG,
    LET_COL_PRICE_TH_MAY,
)
from core.business import (
    is_quote_item,
    is_replacement_material,
    is_percentage_material,
    get_unit_divisor,
    should_link_evidence_sheet,
    calculate_effective_rate,
    format_job_display_name,
)
from core.formulas import ExcelFormulaBuilder


@dataclass(frozen=True)
class ResourceCategoryConfig:
    """Cấu hình chiến lược ghi dữ liệu cho từng nhóm tài nguyên hao phí."""
    category_key: str          # Key trong từ điển resources ('materials', 'labor', 'machines')
    category_title: str        # Tiêu đề nhóm hiển thị trên bảng ('Vật liệu', 'Nhân công', 'Máy thi công')
    header_exclusion: str      # Tên dòng tiêu đề cần lọc bỏ ('1. Vật liệu', '2. Nhân công', '3. Máy thi công')
    th_sheet_name: str         # Tên sheet Tổng hợp tham chiếu đơn giá
    th_price_col: str          # Tên cột đơn giá trong sheet Tổng hợp ('G' hoặc 'F')
    default_unit: str          # Đơn vị mặc định nếu thiếu dữ liệu
    apply_k_step: bool         # Cờ áp dụng hệ số biến thể k_step vào liên kết định mức


# Khởi tạo sẵn cấu hình chuẩn cho 3 nhóm tài nguyên
CONFIG_MATERIALS = ResourceCategoryConfig(
    category_key="materials",
    category_title="Vật liệu",
    header_exclusion="1. Vật liệu",
    th_sheet_name=SHEET_TH_VAT_LIEU,
    th_price_col=LET_COL_PRICE_TH_VAT_LIEU,
    default_unit="",
    apply_k_step=False
)

CONFIG_LABOR = ResourceCategoryConfig(
    category_key="labor",
    category_title="Nhân công",
    header_exclusion="2. Nhân công",
    th_sheet_name=SHEET_TH_NHAN_CONG,
    th_price_col=LET_COL_PRICE_TH_NHAN_CONG,
    default_unit="công",
    apply_k_step=True
)

CONFIG_MACHINES = ResourceCategoryConfig(
    category_key="machines",
    category_title="Máy thi công",
    header_exclusion="3. Máy thi công",
    th_sheet_name=SHEET_TH_MAY_THI_CONG,
    th_price_col=LET_COL_PRICE_TH_MAY,
    default_unit="ca",
    apply_k_step=True
)


class SheetDonGiaChiTietBuilder(BaseSheetBuilder):
    """
    Builder xây dựng Sheet 01_ChiTiet_HaoPhi_DinhMuc theo kiến trúc phân tầng,
    hỗ trợ 100% công thức động liên kết sang Sheet Minh Chứng và các Sheet Tổng Hợp.
    """

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        work_items: List[Dict[str, Any]],
        norm_cell_map: Dict[str, Any],
        mat_row_map: Dict[str, int],
        lab_row_map: Dict[str, int],
        mac_row_map: Dict[str, int]
    ) -> Dict[int, Dict[str, Any]]:
        """
        Phương thức điều phối chính: Khởi tạo sheet và duyệt ghi từng công tác dự toán.
        """
        ws = wb.create_sheet(title=SHEET_CHI_TIET_HAO_PHI)
        ws.views.sheetView[0].showGridLines = True

        # 1. Khởi tạo tiêu đề bảng biểu
        self._setup_sheet_headers(ws, project_name, category_name)

        curr_r = 6
        item_subtotal_map: Dict[int, Dict[str, Any]] = {}

        # 2. Xử lý từng công tác trong danh mục dự án
        for item in work_items:
            stt = item.get("stt", 1)
            code = (item.get("norm_code") or "").strip()
            norm = self.norm_manager.get_norm(code)

            # Phân nhánh: Công tác Báo giá / Tạm tính vs Công tác Định mức Nhà nước
            if is_quote_item(item, norm):
                curr_r, subtotal_info = self._write_quote_item(ws, curr_r, item, stt, code)
                item_subtotal_map[stt] = subtotal_info
            else:
                curr_r, subtotal_info = self._write_norm_item(
                    ws=ws,
                    start_row=curr_r,
                    item=item,
                    stt=stt,
                    code=code,
                    norm=norm,
                    norm_cell_map=norm_cell_map,
                    mat_row_map=mat_row_map,
                    lab_row_map=lab_row_map,
                    mac_row_map=mac_row_map
                )
                item_subtotal_map[stt] = subtotal_info

        # 3. Kẻ khung viền toàn bộ dữ liệu bảng tính
        self._apply_table_borders(ws, start_row=5, end_row=curr_r)

        # 4. Thêm khối chữ ký thẩm định 3 bên chuẩn EVN
        self.add_signature_block(ws, start_row=curr_r, col_left=2, col_mid=5, col_right=10)

        # 5. Cấu hình Freeze Panes (C6), AutoFilter (hàng 5) và Thiết lập trang in A4 chuẩn F1
        self.setup_table_view(ws, freeze_cell="C6", filter_ref=f"A5:J{curr_r-1}")
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="5:5")

        return item_subtotal_map

    # ==========================================================================
    # TIỂU PHÂN HỆ 1: KHỞI TẠO TIÊU ĐỀ & ĐỊNH DẠNG BẢNG BIỂU
    # ==========================================================================
    def _setup_sheet_headers(self, ws: Worksheet, project_name: str, category_name: str) -> None:
        """Thiết lập thông tin dự án và tiêu đề cột bảng đơn giá chi tiết chuẩn F1."""
        self.setup_header_banner(ws, "BẢNG ĐƠN GIÁ CHI TIẾT", project_name, category_name, max_col=10)

        headers = ["STT", "Mã hiệu\nđơn giá", "", "", "Tên công tác", "Đơn vị", "Định mức", "Đơn giá", "Hệ số", "Thành tiền"]
        ws.row_dimensions[5].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, c_idx, h)
            cell.font = self.font_header
            cell.fill = self.fill_header
            cell.alignment = self.align_center
            cell.border = self.border_cell

    def _apply_table_borders(self, ws: Worksheet, start_row: int, end_row: int) -> None:
        """Kẻ khung viền mỏng chuẩn EVN cho toàn bộ phạm vi bảng tính."""
        for r in range(start_row, end_row):
            for c in range(1, 11):
                ws.cell(r, c).border = self.border_cell

    # ==========================================================================
    # TIỂU PHÂN HỆ 2: GHI CÔNG TÁC BÁO GIÁ / TẠM TÍNH
    # ==========================================================================
    def _write_quote_item(
        self,
        ws: Worksheet,
        row: int,
        item: Dict[str, Any],
        stt: int,
        code: str
    ) -> Tuple[int, Dict[str, Any]]:
        """Ghi dữ liệu cho công tác báo giá thị trường (1 dòng duy nhất, định mức = 1.0)."""
        ws.row_dimensions[row].height = 22
        
        # Cột STT & Mã hiệu
        ws.cell(row, COL_STT, stt).font = self.font_bold
        ws.cell(row, COL_STT).alignment = self.align_center
        ws.cell(row, COL_CODE, code).font = self.font_code
        ws.cell(row, COL_CODE).alignment = self.align_center

        # Tên & ĐVT
        ws.cell(row, COL_NAME, item.get("job_name", "")).font = self.font_bold
        ws.cell(row, COL_UNIT, item.get("unit", "")).alignment = self.align_center

        # Định mức (1.0) & Đơn giá báo giá
        ws.cell(row, COL_RATE, 1.0).alignment = self.align_right
        ws.cell(row, COL_RATE).number_format = FORMAT_QUANTITY

        u_quote = float(item.get("unit_price_quote", 0.0))
        c_h = ws.cell(row, COL_PRICE, u_quote)
        c_h.font = self.font_bold
        c_h.alignment = self.align_right
        c_h.number_format = FORMAT_CURRENCY

        # Thành tiền: =G*H
        c_j = ws.cell(row, COL_AMOUNT, ExcelFormulaBuilder.item_amount(row))
        c_j.font = self.font_bold
        c_j.alignment = self.align_right
        c_j.number_format = FORMAT_CURRENCY

        subtotal_info = {
            "header_row": row,
            "mat_row": None,
            "lab_row": row,
            "mac_row": None,
            "is_quote": True
        }
        return row + 1, subtotal_info

    # ==========================================================================
    # TIỂU PHÂN HỆ 3: GHI CÔNG TÁC ĐỊNH MỨC TIÊU CHUẨN
    # ==========================================================================
    def _write_norm_item(
        self,
        ws: Worksheet,
        start_row: int,
        item: Dict[str, Any],
        stt: int,
        code: str,
        norm: Dict[str, Any],
        norm_cell_map: Dict[str, Any],
        mat_row_map: Dict[str, int],
        lab_row_map: Dict[str, int],
        mac_row_map: Dict[str, int]
    ) -> Tuple[int, Dict[str, Any]]:
        """Ghi trọn vẹn một công tác định mức bao gồm dòng tiêu đề và 3 nhóm hao phí."""
        eff = self.norm_manager.get_effective_item_resources(item)
        divisor = get_unit_divisor(norm.get("unit", ""), item.get("unit", norm.get("unit", "")))
        k_step = eff.get("k_step", 1.0)
        exclude_rep = eff.get("exclude_replacement", False)

        # 1. Ghi dòng tiêu đề công tác
        job_hdr_row = start_row
        self._write_job_header(ws, job_hdr_row, item, stt, code, norm, eff)
        curr_r = job_hdr_row + 1

        # Lấy tọa độ tra cứu ô định mức trên Sheet Minh chứng của mã hiệu này
        ev_coords = norm_cell_map.get(code, {})
        ev_mat_map = ev_coords.get("materials", {})
        ev_lab_map = ev_coords.get("labor", {})
        ev_mac_map = ev_coords.get("machines", {})

        # 2. Ghi 3 nhóm tài nguyên hao phí qua các hàm chuyên biệt
        curr_r, row_mat = self._write_materials(
            ws=ws, start_row=curr_r, eff=eff, norm_code=code, divisor=divisor,
            exclude_rep=exclude_rep, ev_mat_map=ev_mat_map, mat_row_map=mat_row_map
        )
        curr_r, row_lab = self._write_labors(
            ws=ws, start_row=curr_r, eff=eff, norm_code=code, divisor=divisor,
            k_step=k_step, ev_lab_map=ev_lab_map, lab_row_map=lab_row_map
        )
        curr_r, row_mac = self._write_machines(
            ws=ws, start_row=curr_r, eff=eff, norm_code=code, divisor=divisor,
            k_step=k_step, ev_mac_map=ev_mac_map, mac_row_map=mac_row_map
        )

        # 3. Tổng hợp thành tiền tại dòng tiêu đề công tác
        subtotal_rows = [r for r in [row_mat, row_lab, row_mac] if r is not None]
        if subtotal_rows:
            sum_cells = [f"{LET_COL_AMOUNT}{r}" for r in subtotal_rows]
            ws.cell(job_hdr_row, COL_AMOUNT, ExcelFormulaBuilder.sum_discrete_cells(sum_cells)).font = self.font_bold
            ws.cell(job_hdr_row, COL_AMOUNT).number_format = FORMAT_CURRENCY

        subtotal_info = {
            "header_row": job_hdr_row,
            "mat_row": row_mat,
            "lab_row": row_lab,
            "mac_row": row_mac,
            "is_quote": False
        }
        return curr_r, subtotal_info

    def _write_job_header(
        self,
        ws: Worksheet,
        row: int,
        item: Dict[str, Any],
        stt: int,
        code: str,
        norm: Dict[str, Any],
        eff: Dict[str, Any]
    ) -> None:
        """Ghi thông tin dòng đầu mục công tác (Mã hiệu nền xanh, tên có biến thể)."""
        ws.row_dimensions[row].height = 22
        ws.cell(row, COL_STT, stt).font = self.font_bold
        ws.cell(row, COL_STT).alignment = self.align_center

        c_code = ws.cell(row, COL_CODE, code)
        c_code.font = self.font_code
        c_code.alignment = self.align_center
        c_code.fill = self.fill_code

        job_name_txt = format_job_display_name(item, norm, eff)
        ws.cell(row, COL_NAME, job_name_txt).font = self.font_bold
        ws.cell(row, COL_UNIT, item.get("unit", norm.get("unit", ""))).alignment = self.align_center

    # ==========================================================================
    # TIỂU PHÂN HỆ 4: CÁC HÀM XỬ LÝ CHUYÊN BIỆT TỪNG NHÓM TÀI NGUYÊN
    # ==========================================================================
    def _write_materials(
        self,
        ws: Worksheet,
        start_row: int,
        eff: Dict[str, Any],
        norm_code: str,
        divisor: float,
        exclude_rep: bool,
        ev_mat_map: Dict[str, int],
        mat_row_map: Dict[str, int]
    ) -> Tuple[int, Optional[int]]:
        """Xử lý chuyên biệt nhóm Vật liệu (hỗ trợ vật liệu phụ % và vật tư thay thế)."""
        return self._write_resource_section(
            ws=ws,
            start_row=start_row,
            config=CONFIG_MATERIALS,
            resources=eff.get("materials", []),
            norm_code=norm_code,
            divisor=divisor,
            k_step=1.0,
            exclude_replacement=exclude_rep,
            evidence_row_map=ev_mat_map,
            th_row_map=mat_row_map
        )

    def _write_labors(
        self,
        ws: Worksheet,
        start_row: int,
        eff: Dict[str, Any],
        norm_code: str,
        divisor: float,
        k_step: float,
        ev_lab_map: Dict[str, int],
        lab_row_map: Dict[str, int]
    ) -> Tuple[int, Optional[int]]:
        """Xử lý chuyên biệt nhóm Nhân công (hỗ trợ hệ số biến thể k_step)."""
        return self._write_resource_section(
            ws=ws,
            start_row=start_row,
            config=CONFIG_LABOR,
            resources=eff.get("labor", []),
            norm_code=norm_code,
            divisor=divisor,
            k_step=k_step,
            exclude_replacement=False,
            evidence_row_map=ev_lab_map,
            th_row_map=lab_row_map
        )

    def _write_machines(
        self,
        ws: Worksheet,
        start_row: int,
        eff: Dict[str, Any],
        norm_code: str,
        divisor: float,
        k_step: float,
        ev_mac_map: Dict[str, int],
        mac_row_map: Dict[str, int]
    ) -> Tuple[int, Optional[int]]:
        """Xử lý chuyên biệt nhóm Máy thi công (hỗ trợ hệ số biến thể k_step)."""
        return self._write_resource_section(
            ws=ws,
            start_row=start_row,
            config=CONFIG_MACHINES,
            resources=eff.get("machines", []),
            norm_code=norm_code,
            divisor=divisor,
            k_step=k_step,
            exclude_replacement=False,
            evidence_row_map=ev_mac_map,
            th_row_map=mac_row_map
        )

    # ==========================================================================
    # TIỂU PHÂN HỆ 5: QUY TRÌNH GHI TÀI NGUYÊN TỔNG QUÁT (STRATEGY PATTERN)
    # ==========================================================================
    def _write_resource_section(
        self,
        ws: Worksheet,
        start_row: int,
        config: ResourceCategoryConfig,
        resources: List[Dict[str, Any]],
        norm_code: str,
        divisor: float,
        k_step: float,
        exclude_replacement: bool,
        evidence_row_map: Dict[str, int],
        th_row_map: Dict[str, int]
    ) -> Tuple[int, Optional[int]]:
        """
        Quy trình ghi tài nguyên tổng quát dùng chung cho Vật liệu, Nhân công và Máy thi công.
        
        Quy tắc:
        1. Lọc bỏ các dòng tiêu đề trùng lặp trong cơ sở dữ liệu.
        2. Tạo dòng phân nhóm (Category Header) in đậm nghiêng màu xanh navy.
        3. Ghi từng dòng tài nguyên chi tiết với đầy đủ liên kết công thức.
        4. Tạo dòng công thức tổng nhóm: =SUM(J_start:J_end).
        """
        valid_items = [
            r for r in resources 
            if r.get("code") and r.get("name", "").strip() != config.header_exclusion
        ]
        if not valid_items:
            return start_row, None

        # Dòng tiêu đề phân nhóm (Ví dụ: "Vật liệu", "Nhân công", "Máy thi công")
        row_cat = start_row
        ws.cell(row_cat, COL_NAME, config.category_title).font = self.font_cat

        curr_r = row_cat + 1
        data_start_r = curr_r

        # Ghi từng tài nguyên thành phần
        for item in valid_items:
            self._write_single_resource_row(
                ws=ws,
                row=curr_r,
                config=config,
                res_item=item,
                norm_code=norm_code,
                divisor=divisor,
                k_step=k_step,
                exclude_replacement=exclude_replacement,
                evidence_row_map=evidence_row_map,
                th_row_map=th_row_map,
                data_start_r=data_start_r
            )
            curr_r += 1

        data_end_r = curr_r - 1

        # Công thức tính tổng nhóm tại Cột J của dòng tiêu đề phân nhóm
        ws.cell(
            row_cat,
            COL_AMOUNT,
            ExcelFormulaBuilder.sum_range(LET_COL_AMOUNT, data_start_r, data_end_r)
        ).font = self.font_cat
        ws.cell(row_cat, COL_AMOUNT).number_format = FORMAT_CURRENCY

        return curr_r, row_cat

    def _write_single_resource_row(
        self,
        ws: Worksheet,
        row: int,
        config: ResourceCategoryConfig,
        res_item: Dict[str, Any],
        norm_code: str,
        divisor: float,
        k_step: float,
        exclude_replacement: bool,
        evidence_row_map: Dict[str, int],
        th_row_map: Dict[str, int],
        data_start_r: int
    ) -> None:
        """Ghi một dòng tài nguyên chi tiết (Mã, Tên, ĐVT, Định mức, Đơn giá, Thành tiền)."""
        res_code = res_item.get("code", "")
        res_name = res_item.get("name", "")
        res_unit = res_item.get("unit") or config.default_unit
        base_rate = float(res_item.get("rate", 0.0))

        # Đánh giá các quy tắc nghiệp vụ đặc thù
        is_rep = is_replacement_material(res_code, exclude_replacement)
        is_pct = is_percentage_material(res_item)
        r_ev = evidence_row_map.get(res_code)
        effective_k_step = k_step if config.apply_k_step else 1.0
        can_link_ev = should_link_evidence_sheet(
            has_evidence_row=(r_ev is not None),
            is_replacement=is_rep,
            is_percentage=is_pct
        )

        # 1. Mã tài nguyên (Cột B)
        ws.cell(row, COL_CODE, res_code).font = self.font_regular
        ws.cell(row, COL_CODE).alignment = self.align_center

        # 2. Tên tài nguyên (Cột E)
        disp_name = f" - {res_name} (Không tính vật tư thay thế - ĐG = 0)" if is_rep else f" - {res_name}"
        ws.cell(row, COL_NAME, disp_name).font = self.font_regular

        # 3. Đơn vị tính (Cột F)
        ws.cell(row, COL_UNIT, res_unit).alignment = self.align_center

        # 4. Định mức hao phí (Cột G)
        if can_link_ev:
            rate_formula = ExcelFormulaBuilder.evidence_rate_link(
                norm_code=norm_code,
                evidence_row=r_ev,
                divisor=divisor,
                k_step=effective_k_step
            )
            ws.cell(row, COL_RATE, rate_formula).alignment = self.align_right
        else:
            eff_rate = calculate_effective_rate(base_rate, divisor=divisor, k_step=effective_k_step)
            ws.cell(row, COL_RATE, eff_rate).alignment = self.align_right
        ws.cell(row, COL_RATE).number_format = FORMAT_RATE

        # 5. Đơn giá tài nguyên (Cột H) & Thành tiền (Cột J)
        if is_rep:
            # Vật tư thay thế: Đơn giá = 0, Thành tiền = 0
            c_price = ws.cell(row, COL_PRICE, 0)
            c_price.alignment = self.align_right
            c_price.number_format = FORMAT_CURRENCY
            c_amt = ws.cell(row, COL_AMOUNT, 0)
            c_amt.alignment = self.align_right
            c_amt.number_format = FORMAT_CURRENCY

        elif is_pct:
            # Vật liệu phụ %: Tính theo % trên tổng vật liệu chính trước nó
            pct_formula = ExcelFormulaBuilder.percentage_material_amount(
                amount_col=LET_COL_AMOUNT,
                mat_start_row=data_start_r,
                prev_row=row - 1,
                rate_col=LET_COL_RATE,
                curr_row=row
            )
            c_amt = ws.cell(row, COL_AMOUNT, pct_formula)
            c_amt.alignment = self.align_right
            c_amt.number_format = FORMAT_CURRENCY

        else:
            # Tài nguyên thông thường: Link đơn giá từ sheet TH tương ứng
            r_th = th_row_map.get(res_code)
            if r_th:
                price_formula = ExcelFormulaBuilder.th_price_link(
                    th_sheet_name=config.th_sheet_name,
                    price_col=config.th_price_col,
                    th_row=r_th
                )
                c_price = ws.cell(row, COL_PRICE, price_formula)
            else:
                c_price = ws.cell(row, COL_PRICE, 0)

            c_price.fill = self.fill_yellow
            c_price.alignment = self.align_right
            c_price.number_format = FORMAT_CURRENCY

            # Thành tiền = G * H
            c_amt = ws.cell(row, COL_AMOUNT, ExcelFormulaBuilder.item_amount(row))
            c_amt.alignment = self.align_right
            c_amt.number_format = FORMAT_CURRENCY
