# -*- coding: utf-8 -*-
"""
Module: core.sheets.sheet_khoi_luong
Tầng xuất dữ liệu Excel (Presentation Layer) cho:
1. Sheet 'Khối lượng-Định mức-ĐVT': Bảng 16 cột chuẩn theo PHƯƠNG ÁN ĐỀ XUẤT TỐI ƯU.
2. Sheet 'Doi_chieu_dinh_muc': Bảng đối chiếu 3 chiều (31 cột, 883 dòng) sao chép nguyên bản
   từ file master 'DANH MUC CONG TAC_da tra ma.xlsx'.

Tác giả: Tổ Thẩm Định EVN
"""

from pathlib import Path
from typing import Dict, List, Any, Optional
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from core.sheets.base import BaseSheetBuilder
from core.constants import (
    SHEET_KHOI_LUONG,
    SHEET_DOI_CHIEU_DINH_MUC,
    FORMAT_QUANTITY
)

# Đường dẫn file master đối chiếu
MASTER_FILE = Path(__file__).resolve().parent.parent / "templates" / "DANH_MUC_CONG_TAC_da_tra_ma.xlsx"
FALLBACK_DOWNLOAD_FILE = Path(r"C:\Users\Admin\Downloads\DANH MUC CONG TAC_da tra ma.xlsx")


class SheetKhoiLuongBuilder(BaseSheetBuilder):
    """Builder xây dựng Sheet Khối lượng (Đề xuất tối ưu) và Sheet Đối chiếu định mức (31 cột)."""

    def build(
        self,
        wb: openpyxl.Workbook,
        project_name: str,
        category_name: str,
        work_items: List[Dict[str, Any]],
        item_subtotal_map: Optional[Dict[Any, Dict[str, Any]]] = None,
        coefficients: Optional[Dict[str, Any]] = None
    ) -> Optional[int]:
        src_path = MASTER_FILE if MASTER_FILE.exists() else FALLBACK_DOWNLOAD_FILE
        if not src_path.exists():
            print(f"Warning: Master comparison file not found at {src_path}")
            return None

        # Load master workbook
        wb_src = openpyxl.load_workbook(src_path, data_only=True)
        if "Doi_chieu_dinh_muc" not in wb_src.sheetnames:
            print("Warning: Sheet 'Doi_chieu_dinh_muc' not found in master template.")
            return None

        ws_src_dc = wb_src["Doi_chieu_dinh_muc"]

        # ======================================================================
        # 1. TẠO SHEET 1: Khối lượng-Định mức-ĐVT (16 cột theo Phương án Tối ưu)
        # ======================================================================
        ws_kl = wb.create_sheet(title=SHEET_KHOI_LUONG)
        self._build_optimal_khoi_luong_sheet(ws_kl, ws_src_dc, project_name, category_name)

        # ======================================================================
        # 2. TẠO SHEET 2: Doi_chieu_dinh_muc (31 cột đối chiếu 3 chiều đầy đủ)
        # ======================================================================
        ws_dc = wb.create_sheet(title=SHEET_DOI_CHIEU_DINH_MUC)
        self._clone_full_doi_chieu_sheet(ws_dc, ws_src_dc)

        return None

    def _build_optimal_khoi_luong_sheet(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        ws_src: openpyxl.worksheet.worksheet.Worksheet,
        project_name: str,
        category_name: str
    ) -> None:
        """Xây dựng bảng khối lượng 16 cột theo phương án đề xuất tối ưu."""
        # 1. Tiêu đề biểu mẫu chuẩn F1
        self.setup_header_banner(
            ws=ws,
            title="BẢNG TIÊN LƯỢNG KHỐI LƯỢNG VÀ ÁP MÃ ĐỊNH MỨC THEO PHƯƠNG ÁN ĐỀ XUẤT TỐI ƯU",
            project_name=project_name,
            category_name=category_name,
            max_col=16
        )

        # 2. Hàng tiêu đề bảng (Hàng 5)
        headers = [
            "STT", "Danh mục hàng hóa/dịch vụ (THEO QUYẾT ĐỊNH 1246/QĐ-NĐVT4)",
            "Khu vực / Vị trí", "Mô tả hàng hoá/dịch vụ", "Khối lượng", "Đơn vị tính",
            "Mã tham khảo", "Hướng dẫn", "Mã định mức đề xuất tối ưu", "Tên công việc theo định mức",
            "ĐVT theo định mức", "STT thành phần công việc áp dụng", "Loại mã là SCL hay TT38",
            "Tên công việc sau chuyển đổi theo định mức", "Khối lượng sau chuyển đổi",
            "Căn cứ & Lý do chọn tối ưu"
        ]
        ws.row_dimensions[5].height = 28
        for c_idx, h in enumerate(headers, start=1):
            cell = ws.cell(5, c_idx, h)
            cell.font = self.font_header
            cell.alignment = self.align_center
            cell.border = self.border_cell
            cell.fill = self.fill_header
            if c_idx == 9:
                cell.fill = self.fill_code

        # 3. Ghi dữ liệu 876 công tác từ hàng 6 của sheet Doi_chieu_dinh_muc
        curr_r = 6
        for r_src in range(5, ws_src.max_row + 1):
            val_b = ws_src.cell(r_src, 2).value
            stt = ws_src.cell(r_src, 1).value

            # Tiêu đề nhóm phân đoạn (PHẦN DỊCH VỤ / PHẦN VẬT TƯ)
            if val_b and ("PHẦN DỊCH VỤ" in str(val_b) or "PHẦN VẬT TƯ" in str(val_b)):
                ws.row_dimensions[curr_r].height = 24
                ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=16)
                c_sec = ws.cell(curr_r, 1, str(val_b).strip())
                c_sec.font = Font(name="Times New Roman", size=11, bold=True, color="002060")
                c_sec.fill = self.fill_header
                c_sec.alignment = self.align_left
                for c in range(1, 17):
                    ws.cell(curr_r, c).border = self.border_cell
                curr_r += 1
                continue

            if stt is None or not isinstance(stt, int):
                continue

            ws.row_dimensions[curr_r].height = 20

            # Lấy thông tin nguồn
            name_1246 = ws_src.cell(r_src, 2).value
            loc = ws_src.cell(r_src, 3).value
            desc = ws_src.cell(r_src, 4).value
            vol_goc = ws_src.cell(r_src, 5).value
            dvt_goc = ws_src.cell(r_src, 6).value
            ref_code = ws_src.cell(r_src, 7).value
            guide = ws_src.cell(r_src, 8).value

            opt_code = ws_src.cell(r_src, 27).value   # Cột AA: Đề xuất tối ưu
            opt_reason = ws_src.cell(r_src, 28).value # Cột AB: Lý do chọn

            # Nhận diện thông tin tương ứng giữa KHVT rà và KTAT
            khvt_code = str(ws_src.cell(r_src, 9).value or "").strip()
            if khvt_code and khvt_code in str(opt_code or ""):
                dm_name = ws_src.cell(r_src, 10).value
                dm_dvt = ws_src.cell(r_src, 11).value
                dm_step = ws_src.cell(r_src, 12).value
                dm_type = ws_src.cell(r_src, 13).value
                conv_name = ws_src.cell(r_src, 14).value
                conv_vol = ws_src.cell(r_src, 15).value
            else:
                dm_name = ws_src.cell(r_src, 19).value
                dm_dvt = ws_src.cell(r_src, 20).value
                dm_step = ws_src.cell(r_src, 21).value
                dm_type = ws_src.cell(r_src, 22).value
                conv_name = ws_src.cell(r_src, 23).value
                conv_vol = ws_src.cell(r_src, 24).value

            # Ghi các cột
            ws.cell(curr_r, 1, stt).alignment = self.align_center
            ws.cell(curr_r, 2, name_1246).alignment = self.align_left
            ws.cell(curr_r, 3, loc).alignment = self.align_left
            ws.cell(curr_r, 4, desc).alignment = self.align_left

            # Khối lượng gốc
            c_v = ws.cell(curr_r, 5, vol_goc)
            c_v.alignment = self.align_right
            c_v.number_format = FORMAT_QUANTITY

            ws.cell(curr_r, 6, dvt_goc).alignment = self.align_center
            ws.cell(curr_r, 7, ref_code).alignment = self.align_center
            ws.cell(curr_r, 8, guide).alignment = self.align_left

            # Cột 9: Mã tối ưu
            c_opt = ws.cell(curr_r, 9, opt_code)
            c_opt.alignment = self.align_center
            c_opt.font = self.font_code
            if opt_code and "Không áp mã" not in str(opt_code):
                c_opt.fill = self.fill_code
            else:
                c_opt.font = Font(name="Times New Roman", size=10, italic=True, color="555555")

            ws.cell(curr_r, 10, dm_name).alignment = self.align_left
            ws.cell(curr_r, 11, dm_dvt).alignment = self.align_center
            ws.cell(curr_r, 12, dm_step).alignment = self.align_center
            ws.cell(curr_r, 13, dm_type).alignment = self.align_center
            ws.cell(curr_r, 14, conv_name).alignment = self.align_left

            # Khối lượng sau chuyển đổi
            c_cv = ws.cell(curr_r, 15, conv_vol if conv_vol is not None else vol_goc)
            c_cv.alignment = self.align_right
            c_cv.number_format = FORMAT_QUANTITY

            # Cột 16: Lý do chọn tối ưu
            ws.cell(curr_r, 16, opt_reason).alignment = self.align_left

            for c in range(1, 17):
                ws.cell(curr_r, c).border = self.border_cell
                if c not in [1, 5, 9, 15]:
                    ws.cell(curr_r, c).font = self.font_regular

            curr_r += 1

        # Cấu hình Freeze Panes, AutoFilter và Page Setup
        self.setup_table_view(ws, freeze_cell="C6", filter_ref=f"A5:P{curr_r-1}")
        self.setup_page_printing(ws, orientation="landscape", repeat_rows="5:5")

        # Độ rộng cột tối ưu
        col_w = {
            1: 6, 2: 40, 3: 25, 4: 30, 5: 12, 6: 10,
            7: 15, 8: 15, 9: 22, 10: 35, 11: 12, 12: 15,
            13: 10, 14: 35, 15: 14, 16: 45
        }
        for c, w in col_w.items():
            ws.column_dimensions[get_column_letter(c)].width = w

    def _clone_full_doi_chieu_sheet(
        self,
        ws_dst: openpyxl.worksheet.worksheet.Worksheet,
        ws_src: openpyxl.worksheet.worksheet.Worksheet
    ) -> None:
        """Sao chép toàn bộ 31 cột x 883 dòng từ sheet Doi_chieu_dinh_muc gốc."""
        from copy import copy
        ws_dst.views.sheetView[0].showGridLines = True

        for r in range(1, ws_src.max_row + 1):
            for c in range(1, ws_src.max_column + 1):
                c_src = ws_src.cell(r, c)
                c_dst = ws_dst.cell(r, c)
                c_dst.value = c_src.value
                if c_src.has_style:
                    c_dst.font = copy(c_src.font)
                    c_dst.fill = copy(c_src.fill)
                    c_dst.border = copy(c_src.border)
                    c_dst.alignment = copy(c_src.alignment)
                    c_dst.number_format = c_src.number_format

        # Copy column dimensions
        for col_l, col_d in ws_src.column_dimensions.items():
            if col_d.width:
                ws_dst.column_dimensions[col_l].width = col_d.width

        # Copy row heights
        for r_idx, r_dim in ws_src.row_dimensions.items():
            if r_dim.height:
                ws_dst.row_dimensions[r_idx].height = r_dim.height

        # Setup Freeze Panes và AutoFilter
        self.setup_table_view(ws_dst, freeze_cell="C6", filter_ref=f"A2:AE{ws_src.max_row}")
        self.setup_page_printing(ws_dst, orientation="landscape", repeat_rows="2:2")
