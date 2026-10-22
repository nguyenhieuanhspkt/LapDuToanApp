"""
Streamlit Web App: Hệ Thống Quản Lý Định Mức & Lập Dự Toán PCCC - SCL EVN
Author: Tổ Thẩm Định EVN
"""

import streamlit as st
import pandas as pd
import json
import io
import openpyxl
from pathlib import Path
import sys
import re

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import DATA_DIR, OUTPUT_DIR, IMAGES_DIR, check_environment
from core.norm_manager import NormManager
from core.excel_generator import ExcelGenerator, DEFAULT_COEFFICIENTS
from utils.num2words_vi import num2words_vi

def create_single_summary_sheet_excel(sheet_type: str, p_name: str, c_name: str, items_data: list, n_mgr: NormManager) -> bytes:
    wb = openpyxl.Workbook()
    ws_dummy = wb.active
    gen = ExcelGenerator(n_mgr)
    if sheet_type == "material":
        gen._build_sheet_th_vat_lieu(wb, p_name, c_name, items_data)
    elif sheet_type == "labor":
        gen._build_sheet_th_nhan_cong(wb, p_name, c_name, items_data)
    elif sheet_type == "machine":
        gen._build_sheet_th_may_thi_cong(wb, p_name, c_name, items_data)
    wb.remove(ws_dummy)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# Set Page Config
st.set_page_config(
    page_title="Quản Lý & Lập Dự Toán PCCC - EVN",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling (EVN Theme)
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 700;
        color: #002060;
        margin-bottom: 5px;
    }
    .sub-title {
        font-size: 14px;
        color: #555;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #F2F7FC;
        border-left: 5px solid #002060;
        padding: 12px 18px;
        border-radius: 6px;
        margin-bottom: 15px;
    }
    .group-header {
        background-color: #EBF1F5;
        border-left: 4px solid #002060;
        padding: 8px 14px;
        font-weight: bold;
        color: #002060;
        border-radius: 4px;
        margin-top: 10px;
        margin-bottom: 10px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 10px 20px;
        font-weight: 600;
        background-color: #F8F9FA;
        border-radius: 6px 6px 0 0;
    }
    .stTabs [aria-selected="true"] {
        background-color: #002060 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Initialize Session State
if "norm_manager" not in st.session_state or not hasattr(st.session_state.norm_manager, "load_work_items"):
    import importlib
    import core.norm_manager
    importlib.reload(core.norm_manager)
    st.session_state.norm_manager = core.norm_manager.NormManager()

if "coefficients" not in st.session_state:
    st.session_state.coefficients = {
        "c": {"rate": 7.3, "basis": "Thông tư số 11/2021/TT-BXD (Công trình công nghiệp)"},
        "tt": {"rate": 2.5, "basis": "Thông tư số 11/2021/TT-BXD (Không xác định KL từ thiết kế)"},
        "tl": {"rate": 5.5, "basis": "Thông tư số 11/2021/TT-BXD (Thu nhập chịu thuế tính trước)"},
        "vat": {"rate": 8.0, "basis": "Nghị định 72/2024/NĐ-CP của Chính phủ"},
        "lt": {"rate": 1.1, "basis": "Thông tư số 11/2021/TT-BXD (Nhà tạm ngoài đô thị)"}
    }

mgr = st.session_state.norm_manager
generator = ExcelGenerator(mgr)

if "work_items" not in st.session_state:
    st.session_state.work_items = mgr.load_work_items()

# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1d/EVN_Logo.svg/300px-EVN_Logo.svg.png", width=120)
    st.markdown("### 🏢 THÔNG TIN DỰ ÁN")
    project_name = st.text_area(
        "Tên công trình:",
        value="BẢO DƯỠNG, SỬA CHỮA VÀ THÍ NGHIỆM ĐỊNH KỲ HỆ THỐNG PCCC NĂM 2026 - NMNĐ VĨNH TÂN 4",
        height=85
    )
    category_name = st.text_area(
        "Hạng mục:",
        value="Hệ thống báo cháy tự động & Chữa cháy (Định mức Thông tư 38/2026/TT-BXD & QĐ 45/QĐ-HĐTV)",
        height=75
    )
    st.divider()

    st.markdown("### 🗺️ ĐỊA BÀN & ĐƠN GIÁ ĐỊA PHƯƠNG")
    selected_province = st.selectbox(
        "Tỉnh / Địa bàn thực hiện:",
        options=["Lâm Đồng (Mặc định)", "Khánh Hòa"],
        index=0,
        help="Mặc định áp dụng tỉnh Lâm Đồng khi chưa rõ địa điểm cụ thể."
    )
    is_lam_dong = "Lâm Đồng" in selected_province
    
    if is_lam_dong:
        region_options = [
            "Vùng IV (Đơn vị xã, huyện còn lại - MẶC ĐỊNH)",
            "Vùng II (Đà Lạt, Bảo Lộc)",
            "Vùng III (Đức Trọng, Đơn Dương...)"
        ]
        default_idx = 0
    else:
        region_options = [
            "Vùng III (Địa bàn Khánh Hòa - MẶC ĐỊNH)",
            "Vùng II (Nha Trang, Cam Ranh)",
            "Vùng IV"
        ]
        default_idx = 0
        
    selected_region = st.selectbox(
        "Vùng áp dụng đơn giá Nhân công & Ca máy:",
        options=region_options,
        index=default_idx
    )
    
    if "Vùng II" in selected_region:
        current_region_key = "vung_2"
    elif "Vùng III" in selected_region:
        current_region_key = "vung_3"
    else:
        current_region_key = "vung_4"
        
    st.caption(f"📌 Đang áp dụng: **{'Lâm Đồng' if is_lam_dong else 'Khánh Hòa'}** | **{selected_region.split('(')[0].strip()}**")
    st.divider()

    st.markdown("### 📈 THỐNG KÊ NHANH")
    st.metric("Tổng số công tác", f"{len(st.session_state.work_items)} công việc")
    st.metric("Định mức sẵn sàng", f"{len(mgr.get_all_norm_codes())} mã")
    st.divider()

    st.markdown("### ⚙️ HỆ THỐNG")
    env = check_environment()
    if env["status"]:
        st.success("🟢 CSDL & Thư mục OneDrive: Sẵn sàng")
    else:
        st.error("🔴 Phát hiện lỗi kết nối dữ liệu")
        for iss in env["issues"]:
            st.caption(iss)

# =============================================================================
# MAIN HEADER
# =============================================================================
st.markdown('<div class="main-title">📊 HỆ THỐNG QUẢN LÝ ĐỊNH MỨC & LẬP DỰ TOÁN PCCC - EVN</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Tự động hóa lập dự toán, quản lý đơn giá trực quan và trích xuất minh chứng định mức 100% công thức động khép kín</div>', unsafe_allow_html=True)

# =============================================================================
# TABS
# =============================================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 1. Khối Lượng Tiên Lượng",
    "💰 2. Quản Lý Đơn Giá & Pháp Lý",
    "📊 3. Bảng Tổng Hợp Kinh Phí",
    "🔍 4. Chi Tiết Đơn Giá & Scan Gốc",
    "🚀 5. Xuất File Excel Thẩm Định"
])

# Dictionary of PAKT Job Name Suggestions mapped by norm code
PAKT_JOB_SUGGESTIONS = {
    "BD.41111": [
        "Thay mới lắp đặt thiết bị đầu báo khói địa chỉ",
        "Lắp đặt đầu báo khói quang điện tử địa chỉ",
        "Thay thế đầu báo nhiệt gia tăng khu vực gian máy",
        "Lắp đặt thiết bị đầu báo và đầu báo cháy (Đầu báo khói, báo nhiệt)"
    ],
    "BD.41121": [
        "Lắp đặt đèn báo cháy",
        "Lắp mới / thay thế đèn báo cháy vị trí khu vực phụ trợ",
        "Lắp đặt đèn báo cháy phòng điều khiển trung tâm"
    ],
    "BD.41131": [
        "Lắp đặt lại nút nhấn khẩn báo cháy",
        "Lắp đặt lại nút nhấn sau khi bảo trì, sửa chữa",
        "Lắp đặt nút ấn báo cháy khẩn cấp gắn tường"
    ],
    "BD.41141": [
        "Lắp đặt chuông báo cháy",
        "Thay thế, lắp đặt chuông báo cháy tự động",
        "Lắp đặt chuông báo cháy âm thanh cao"
    ],
    "BD.42201": [
        "Lắp đặt bộ điều khiển, bộ chuyển mạch PCCC",
        "Lắp đặt bộ chuyển mạch tín hiệu tủ trung gian",
        "Lắp đặt module điều khiển giám sát PCCC"
    ],
    "E6C.3.01": [
        "Bảo trì, sửa chữa, thử nghiệm giả lập tác động nút nhấn khẩn báo cháy",
        "Kiểm tra, bảo dưỡng nút nhấn báo cháy chuyên dụng",
        "Thử nghiệm chức năng tác động nút nhấn báo cháy"
    ],
    "E6C.3.02": [
        "Bảo trì, sửa chữa, hiệu chỉnh, thử nghiệm độ nhạy đầu báo khói/nhiệt",
        "Dịch vụ Bảo trì đầu báo chuyên dụng",
        "Kiểm tra độ nhạy và vệ sinh buồng khói đầu báo địa chỉ"
    ],
    "E6C.3.03": [
        "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng dây báo cháy",
        "Kiểm tra module dây giám sát nhiệt độ cao"
    ],
    "E6C.3.04": [
        "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng cảm biến beam quang",
        "Bảo dưỡng beam khói khoảng cách xa",
        "Hiệu chuẩn căn chỉnh trục thu phát đầu báo beam quang học"
    ],
    "E6C.3.05": [
        "Sửa chữa, bảo dưỡng trung tâm báo cháy FMZ 5000",
        "Bảo dưỡng tủ điều khiển chữa cháy tại chỗ và tủ trung gian",
        "Kiểm tra toàn diện phần cứng và phần mềm tủ trung tâm báo cháy"
    ],
    "E6C.3.06": [
        "Kiểm tra và bảo dưỡng phần điện van Deluge, Wet hệ PCCC",
        "Bảo dưỡng cuộn hút, van xả tràn Deluge và đồng hồ áp",
        "Thí nghiệm tác động cuộn solenoid van Deluge hệ thống foam/nước"
    ],
    "E6B.1.13.01": [
        "Kiểm tra bình accu chuyên dụng phục vụ hệ thống PCCC",
        "Đo nội trở, dung lượng Ah, vô mỡ cọc bình và châm nước cất",
        "Bảo dưỡng hệ thống ắc quy nguồn dự phòng 24VDC tủ PCCC"
    ],
    "E6C.1.31": [
        "Kiểm tra và bảo dưỡng thiết bị đo khí NH3, Clo, H2",
        "Bơm khí chuẩn hiệu chuẩn điểm Zero & Span cho đầu đo khí",
        "Hiệu chuẩn cảm biến phát hiện rò rỉ khí phòng hóa chất"
    ]
}

# -----------------------------------------------------------------------------
# =============================================================================
# MODAL DIALOGS FOR TAB 1 (STREAMLIT @st.dialog - KHÔNG DÙNG EXPANDER)
# =============================================================================

@st.dialog("➕ THÊM CÔNG TÁC MỚI VÀO DỰ ÁN", width="large")
def dialog_add_work_item():
    st.caption("Thêm công việc từ CSDL Định mức hoặc nhập Báo giá trọn gói.")
    work_type = st.radio(
        "Chọn loại hình công tác:",
        ["📌 1. Áp dụng Định mức CSDL (Có hao phí VL, NC, Máy)", "📋 2. Báo giá trọn gói (Không có định mức / Tạm tính)"],
        horizontal=True,
        key="dlg_add_work_type"
    )

    existing_groups = list(dict.fromkeys([it.get("group", "Nhóm công việc chung") for it in st.session_state.work_items if it.get("group")]))
    if not existing_groups:
        existing_groups = [
            "Hệ thống báo cháy tự động và thiết bị ngoại vi (Lắp đặt - Thay mới)",
            "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy",
            "Bảo dưỡng van tràn Deluge, hệ thống ắc quy và thiết bị đo khí"
        ]

    col_grp_sel, col_grp_new = st.columns([6, 6])
    with col_grp_sel:
        selected_group_choice = st.selectbox("Chọn nhóm công tác:", options=existing_groups + ["➕ [Tạo nhóm mới...]"], key="dlg_grp_sel")
    with col_grp_new:
        if selected_group_choice == "➕ [Tạo nhóm mới...]":
            new_group_name = st.text_input("Nhập tên nhóm mới:", value="Nhóm công việc mới", key="dlg_grp_new")
            final_group = new_group_name.strip()
        else:
            final_group = selected_group_choice

    if "1. Áp dụng Định mức" in work_type:
        all_codes = mgr.get_all_norm_codes()
        norm_options = []
        for c in all_codes:
            nm = mgr.get_norm(c)
            norm_options.append({
                "code": c,
                "label": f"{c} - {nm.get('name', '')} (ĐVT gốc: {nm.get('unit', '')})",
                "name": nm.get("name", ""),
                "unit": nm.get("unit", ""),
                "basis": nm.get("legal_basis", ""),
                "tech_profile": mgr.get_default_tech_profile(c)
            })

        col_search, col_tech = st.columns([7, 5])
        with col_search:
            selected_option = st.selectbox(
                "Tra cứu định mức (Gõ mã hoặc tên):",
                options=norm_options,
                format_func=lambda x: x["label"],
                key="dlg_norm_search"
            )
        with col_tech:
            st.info(f"🎯 **Bản chất kỹ thuật:**\n*{selected_option['tech_profile']}*")

        sel_code = selected_option["code"]
        raw_suggestions = PAKT_JOB_SUGGESTIONS.get(sel_code, [])
        if selected_option["name"] not in raw_suggestions:
            raw_suggestions = raw_suggestions + [selected_option["name"]]

        selected_sugg = st.selectbox(
            "💡 Gợi ý tên danh mục công việc theo PAKT (chọn để tự động điền):",
            options=raw_suggestions,
            index=0,
            key=f"dlg_pakt_sugg_{sel_code}"
        )

        col_f1, col_f2 = st.columns([7, 5])
        with col_f1:
            inp_job_name = st.text_input("Tên danh mục công việc (Theo PAKT):", value=selected_sugg, key=f"dlg_job_name_{sel_code}")
        with col_f2:
            inp_desc = st.text_input("Mô tả công việc thực tế:", value=selected_sugg, key=f"dlg_desc_{sel_code}")

        # Unit & Quantity conversion
        norm_base_unit = selected_option["unit"]
        col_u1, col_u2, col_u3, col_u4 = st.columns([3, 3, 3, 3])
        with col_u1:
            inp_pakt_unit = st.text_input("ĐVT theo PAKT:", value="Bộ" if "bộ" in norm_base_unit.lower() else ("Cái" if "đầu" in norm_base_unit.lower() or "cái" in norm_base_unit.lower() else norm_base_unit), key=f"dlg_pakt_unit_{sel_code}")
        with col_u2:
            inp_vol_pakt = st.number_input("Khối lượng (PAKT):", min_value=0.01, value=1.0, step=1.0, key=f"dlg_vol_pakt_{sel_code}")
        with col_u3:
            default_conv = 1.0
            if "10" in norm_base_unit and "10" not in inp_pakt_unit:
                default_conv = 0.1
            elif "5" in norm_base_unit and "5" not in inp_pakt_unit:
                default_conv = 0.2
            inp_conv = st.number_input(f"Hệ số quy đổi sang ĐM ({norm_base_unit}):", value=default_conv, step=0.1, format="%.4f", key=f"dlg_conv_{sel_code}")
        with col_u4:
            calc_vol_norm = float(inp_vol_pakt) * float(inp_conv)
            st.metric("Khối lượng tính ĐM", f"{calc_vol_norm:.4f}".rstrip("0").rstrip("."))

        # Variants choice
        avail_vars = mgr.get_variants_for_norm(sel_code)
        var_keys = list(avail_vars.keys())
        col_v1, col_v2 = st.columns([6, 6])
        with col_v1:
            sel_var_id = st.selectbox(
                "Biến thể định mức áp dụng:",
                options=var_keys,
                format_func=lambda x: f"[{x}] {avail_vars[x].get('variant_name', '')}",
                key=f"dlg_var_sel_{sel_code}"
            )
        with col_v2:
            st.caption(f"📌 Căn cứ pháp lý: *{selected_option['basis']}*")

        col_extra1, col_extra2 = st.columns([6, 6])
        with col_extra1:
            inp_extra_quote = st.number_input("Đơn giá báo giá vật tư kèm theo trước thuế (đ - nếu có):", min_value=0, value=0, step=500000, key=f"dlg_extra_quote_{sel_code}")
        with col_extra2:
            inp_extra_basis = st.text_input("Căn cứ chào giá kèm theo (nếu có):", value="", key=f"dlg_extra_basis_{sel_code}", placeholder="Ví dụ: Theo Báo giá số...")

        if st.button("➕ Xác nhận thêm công tác định mức vào dự án", type="primary", use_container_width=True):
            new_stt = len(st.session_state.work_items) + 1
            st.session_state.work_items.append({
                "stt": new_stt,
                "group": final_group,
                "is_quote": False,
                "norm_code": sel_code,
                "tech_profile": selected_option["tech_profile"],
                "job_name": inp_job_name,
                "description": inp_desc,
                "volume": float(calc_vol_norm),
                "unit": norm_base_unit,
                "pakt_unit": inp_pakt_unit,
                "pakt_volume": float(inp_vol_pakt),
                "unit_price_quote": float(inp_extra_quote),
                "quote_basis": inp_extra_basis,
                "variant_id": sel_var_id,
                "exclude_replacement": avail_vars[sel_var_id].get("exclude_replacement", False)
            })
            mgr.save_work_items(st.session_state.work_items)
            st.success(f"Đã thêm công tác '{inp_job_name}' vào nhóm '{final_group}' thành công!")
            st.rerun()

    else:
        # Quote Item
        q_existing_bg = [it.get("norm_code") for it in st.session_state.work_items if str(it.get("norm_code", "")).startswith("BG")]
        next_bg_num = len(q_existing_bg) + 1
        default_bg_code = f"BG.{next_bg_num:02d}"

        col_q1, col_q2, col_q3 = st.columns([3, 5, 4])
        with col_q1:
            inp_q_code = st.text_input("Mã hiệu tạm tính / báo giá:", value=default_bg_code, key="dlg_q_code")
        with col_q2:
            inp_q_tech = st.text_input("Bản chất kỹ thuật:", value="Dịch vụ bảo dưỡng chuyên sâu / OEM", key="dlg_q_tech")
        with col_q3:
            inp_q_basis = st.text_input("Căn cứ báo giá / Văn bản chào giá:", value="Theo Báo giá số ... ngày ... của NSX", key="dlg_q_basis")

        col_q4, col_q5, col_q6, col_q7 = st.columns([5, 4, 2, 2])
        with col_q4:
            inp_q_job = st.text_input("Tên danh mục công việc (Theo PAKT):", value="Dịch vụ kiểm tra và hiệu chuẩn chuyên sâu", key="dlg_q_job")
        with col_q5:
            inp_q_desc = st.text_input("Mô tả công việc thực tế:", value="Thực hiện theo quy trình kiểm tra của nhà sản xuất", key="dlg_q_desc")
        with col_q6:
            inp_q_vol = st.number_input("Khối lượng:", min_value=0.01, value=1.0, step=1.0, key="dlg_q_vol")
        with col_q7:
            inp_q_unit = st.text_input("ĐVT:", value="Gói", key="dlg_q_unit")

        col_p1, col_p2 = st.columns([5, 7])
        with col_p1:
            inp_q_price = st.number_input("Đơn giá báo giá trước thuế (đồng/ĐVT):", min_value=0, value=10000000, step=500000, format="%d", key="dlg_q_price")
        with col_p2:
            st.write("")
            apply_oh = st.checkbox("Áp dụng hệ số chi phí chung và thu nhập tính trước (C=7,3%, TT=2,5%, TL=5,5%)", value=False, key="dlg_q_apply_oh")
            st.caption("*(Lưu ý: Nếu báo giá đã là giá trọn gói trước thuế thì không cần tích chọn)*")

        if st.button("➕ Xác nhận thêm công tác báo giá vào dự án", type="primary", use_container_width=True):
            new_stt = len(st.session_state.work_items) + 1
            st.session_state.work_items.append({
                "stt": new_stt,
                "group": final_group,
                "is_quote": True,
                "norm_code": inp_q_code,
                "tech_profile": inp_q_tech,
                "job_name": inp_q_job,
                "description": inp_q_desc,
                "volume": float(inp_q_vol),
                "unit": inp_q_unit,
                "unit_price_quote": float(inp_q_price),
                "quote_basis": inp_q_basis,
                "apply_overhead": apply_oh
            })
            mgr.save_work_items(st.session_state.work_items)
            st.success(f"Đã thêm công tác báo giá '{inp_q_job}' vào nhóm '{final_group}' thành công!")
            st.rerun()


@st.dialog("📥 NHẬP NHANH HÀNG LOẠT TỪ BẢNG PAKT / EXCEL", width="large")
def dialog_quick_intake_pakt():
    st.markdown("""
    **Hướng dẫn dán (Paste) dữ liệu từ Excel / PAKT:**
    - Sao chép các cột từ file Excel dự toán / PAKT và dán vào ô bên dưới.
    - Định dạng ưu tiên: `STT [Tab] Tên công tác [Tab] ĐVT [Tab] Khối lượng [Tab] [Mã ĐM tùy chọn]`
    - Hệ thống sẽ **tự động phân tích từ khóa**, **gợi ý mã định mức chuẩn EVN**, và **tự động quy đổi ĐVT** (Ví dụ `10 đầu` $\leftrightarrow$ `Bộ`/`Cái`).
    """)

    existing_groups = list(dict.fromkeys([it.get("group", "Nhóm công việc chung") for it in st.session_state.work_items if it.get("group")]))
    if not existing_groups:
        existing_groups = ["Hệ thống báo cháy tự động và thiết bị ngoại vi (Lắp đặt - Thay mới)"]

    col_target_g1, col_target_g2 = st.columns([6, 6])
    with col_target_g1:
        target_group = st.selectbox("Chọn nhóm công việc mục tiêu để nạp vào:", options=existing_groups + ["➕ [Tạo nhóm mới...]"], key="pakt_target_grp")
    with col_target_g2:
        if target_group == "➕ [Tạo nhóm mới...]":
            pakt_new_g = st.text_input("Tên nhóm mới:", value="Nhóm công tác nạp từ PAKT", key="pakt_new_g_inp")
            final_target_group = pakt_new_g.strip()
        else:
            final_target_group = target_group

    paste_data = st.text_area(
        "Dán dữ liệu từ bảng Excel / PAKT vào đây:",
        height=180,
        placeholder="1\tThay mới lắp đặt thiết bị đầu báo khói địa chỉ\tBộ\t10\n2\tLắp đặt đèn báo cháy\tCái\t5\n3\tBảo dưỡng van xả tràn Deluge\tCái\t12\tE6C.3.06",
        key="pakt_paste_text"
    )

    if st.button("🔍 Phân tích dữ liệu & Nhận diện tự động mã định mức", type="primary", use_container_width=True):
        if not paste_data.strip():
            st.warning("Vui lòng dán dữ liệu vào ô trên trước khi phân tích.")
        else:
            parsed_rows = []
            all_codes = set(mgr.get_all_norm_codes())

            def is_num(val):
                s = re.sub(r'[^0-9.,]', '', str(val)).replace(',', '.')
                if not s: return False
                try:
                    float(s)
                    return True
                except:
                    return False

            def to_num(val):
                s = re.sub(r'[^0-9.,]', '', str(val)).replace(',', '.')
                return float(s)

            known_units = {'bộ', 'cái', 'đầu', 'bình', 'trung tâm', 'hệ thống', 'tủ', 'gói', 'm', 'mét', 'kg', 'tấm', 'cuộn', 'điểm', 'trạm'}

            # Keyword mapping for smart recognition
            kw_map = {
                "đầu báo khói": "BD.41111", "đầu báo nhiệt": "BD.41111", "báo khói": "BD.41111", "báo nhiệt": "BD.41111",
                "đèn báo cháy": "BD.41121", "đèn báo": "BD.41121",
                "nút nhấn": "BD.41131", "nút ấn": "BD.41131",
                "chuông": "BD.41141",
                "chuyển mạch": "BD.42201", "module điều khiển": "BD.42201", "module loop": "BD.42201", "bộ nguồn": "BD.42201", "module": "BD.42201",
                "nút nhấn khẩn": "E6C.3.01",
                "hiệu chỉnh đầu báo": "E6C.3.02", "độ nhạy đầu báo": "E6C.3.02", "bảo trì đầu báo": "E6C.3.02",
                "dây báo cháy": "E6C.3.03", "dây nhiệt": "E6C.3.03",
                "beam": "E6C.3.04", "quang học": "E6C.3.04",
                "fmz 5000": "E6C.3.05", "trung tâm báo cháy": "E6C.3.05", "tủ trung tâm": "E6C.3.05",
                "deluge": "E6C.3.06", "xả tràn": "E6C.3.06", "van tràn": "E6C.3.06",
                "accu": "E6B.1.13.01", "ắc quy": "E6B.1.13.01",
                "đo khí": "E6C.1.31", "nh3": "E6C.1.31", "clo": "E6C.1.31"
            }

            for line in paste_data.strip().split("\n"):
                line = line.strip()
                if not line:
                    continue
                # Split by tab or 2+ spaces
                parts = [p.strip() for p in re.split(r"\t+| {2,}", line) if p.strip()]
                if not parts:
                    continue

                # 1. Detect STT
                stt_val = len(parsed_rows) + 1
                if parts[0].isdigit():
                    stt_val = int(parts[0])
                    rem = parts[1:]
                else:
                    rem = parts[:]

                # 2. Check norm code in parts
                found_code = ""
                for p in list(rem):
                    if p in all_codes or p.startswith("BD.") or p.startswith("E6") or p.startswith("BG."):
                        found_code = p
                        rem.remove(p)
                        break

                # 3. Detect Unit
                found_unit = ""
                for p in list(rem):
                    if p.lower() in known_units:
                        found_unit = p
                        rem.remove(p)
                        break
                if not found_unit:
                    for p in list(rem):
                        if len(p) <= 12 and not is_num(p) and any(u in p.lower() for u in known_units):
                            found_unit = p
                            rem.remove(p)
                            break
                if not found_unit:
                    found_unit = "Bộ"

                # 4. Detect Volume (Khối lượng)
                found_vol = 1.0
                for p in reversed(list(rem)):
                    if is_num(p):
                        found_vol = to_num(p)
                        rem.remove(p)
                        break

                # 5. Job name and Description
                job_name = rem[0] if rem else f"Công tác {stt_val}"
                desc = rem[1] if len(rem) > 1 else job_name

                # Auto detect code by keyword if not provided
                if not found_code:
                    j_lower = (job_name + " " + desc).lower()
                    for kw, cd in kw_map.items():
                        if kw in j_lower:
                            found_code = cd
                            break
                if not found_code:
                    found_code = "BD.41111"

                nm = mgr.get_norm(found_code)
                norm_unit = nm.get("unit", "")

                # Auto compute conversion factor
                conv = 1.0
                if "10" in norm_unit and "10" not in found_unit:
                    conv = 0.1
                elif "5" in norm_unit and "5" not in found_unit:
                    conv = 0.2

                parsed_rows.append({
                    "stt": stt_val,
                    "job_name": job_name,
                    "pakt_unit": found_unit,
                    "pakt_vol": found_vol,
                    "norm_code": found_code,
                    "norm_unit": norm_unit,
                    "conv_factor": conv,
                    "calc_vol": round(found_vol * conv, 4),
                    "variant_id": "A1" if found_code == "BD.41111" else "A0"
                })

            st.session_state["pakt_preview_rows"] = parsed_rows
            st.session_state["pakt_final_group"] = final_target_group

    if "pakt_preview_rows" in st.session_state and st.session_state["pakt_preview_rows"]:
        st.markdown(f"##### 📋 Kết quả phân tích ({len(st.session_state['pakt_preview_rows'])} công tác) -> Nhóm: `{st.session_state.get('pakt_final_group', '')}`")
        st.caption("Kiểm tra lại mã ĐM, ĐVT và hệ số quy đổi trước khi xác nhận nạp vào bảng dự toán:")

        df_pakt_prev = pd.DataFrame(st.session_state["pakt_preview_rows"])
        all_norm_codes_list = mgr.get_all_norm_codes()

        edited_pakt_df = st.data_editor(
            df_pakt_prev,
            column_config={
                "stt": st.column_config.NumberColumn("STT", width="small"),
                "job_name": st.column_config.TextColumn("Tên công việc theo PAKT", width="large", required=True),
                "pakt_unit": st.column_config.TextColumn("ĐVT PAKT", width="small"),
                "pakt_vol": st.column_config.NumberColumn("KL PAKT", format="%.2f", width="small"),
                "norm_code": st.column_config.SelectboxColumn("Mã Định mức", options=all_norm_codes_list, width="medium"),
                "norm_unit": st.column_config.TextColumn("ĐVT Định mức", width="small", disabled=True),
                "conv_factor": st.column_config.NumberColumn("Hệ số quy đổi", format="%.4f", width="small"),
                "calc_vol": st.column_config.NumberColumn("KL Tính ĐM", format="%.4f", width="small"),
                "variant_id": st.column_config.TextColumn("Biến thể", width="small")
            },
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="pakt_editor_preview"
        )

        col_n1, col_n2 = st.columns([8, 4])
        with col_n1:
            if st.button(f"🚀 XÁC NHẬN NẠP {len(edited_pakt_df)} CÔNG TÁC VÀO DỰ TOÁN", type="primary", use_container_width=True):
                target_g = st.session_state.get("pakt_final_group", "Nhóm công tác PAKT")
                records = edited_pakt_df.to_dict("records")
                cur_len = len(st.session_state.work_items)
                for r in records:
                    cur_len += 1
                    cd = r.get("norm_code", "BD.41111")
                    nm_info = mgr.get_norm(cd)
                    c_fac = float(r.get("conv_factor", 1.0) or 1.0)
                    p_vol = float(r.get("pakt_vol", 1.0) or 1.0)
                    vol_calc = round(p_vol * c_fac, 4)
                    v_id = r.get("variant_id", "A0")
                    avail_v = mgr.get_variants_for_norm(cd)
                    
                    st.session_state.work_items.append({
                        "stt": cur_len,
                        "group": target_g,
                        "is_quote": False,
                        "norm_code": cd,
                        "tech_profile": mgr.get_default_tech_profile(cd),
                        "job_name": r.get("job_name", ""),
                        "description": r.get("job_name", ""),
                        "volume": vol_calc,
                        "unit": nm_info.get("unit", r.get("norm_unit", "")),
                        "pakt_unit": r.get("pakt_unit", ""),
                        "pakt_volume": p_vol,
                        "unit_price_quote": 0.0,
                        "quote_basis": "",
                        "variant_id": v_id,
                        "exclude_replacement": avail_v.get(v_id, {}).get("exclude_replacement", False)
                    })
                mgr.save_work_items(st.session_state.work_items)
                del st.session_state["pakt_preview_rows"]
                st.success(f"Đã nạp thành công {len(records)} công tác vào nhóm '{target_g}'!")
                st.rerun()
        with col_n2:
            if st.button("❌ Hủy bỏ kết quả phân tích", use_container_width=True):
                del st.session_state["pakt_preview_rows"]
                st.rerun()


@st.dialog("⚙️ QUẢN LÝ BIẾN THỂ ĐỊNH MỨC & ÁP DỤNG HÀNG LOẠT", width="large")
def dialog_manage_variants():
    st.markdown("""
    > **Ý nghĩa biến thể định mức (A0, A1, A2...):**
    > - **A0 (Chuẩn)**: Thực hiện đầy đủ 100% các bước theo định mức gốc.
    > - **A1, A2...**: Giảm trừ các bước không thực hiện (hệ số $k$), đưa giá vật tư thay thế về **0 đ**.
    > - **Gán hàng loạt theo từ khóa**: Phân loại tự động hàng trăm công tác trong 1 cú nhấp.
    """)

    tab_v1, tab_v2 = st.tabs(["🧩 1. Thiết lập Hồ sơ Biến thể", "⚡ 2. Lọc Từ Khóa & Gán Hàng Loạt"])

    with tab_v1:
        all_norm_codes_v = mgr.get_all_norm_codes()
        col_sel_n, col_sel_v = st.columns([6, 6])
        with col_sel_n:
            var_norm_code = st.selectbox("Chọn mã định mức:", options=all_norm_codes_v, key="dlg_var_norm_code")

        norm_data = mgr.get_norm(var_norm_code)
        all_vars_for_norm = mgr.get_variants_for_norm(var_norm_code)
        var_keys = list(all_vars_for_norm.keys())

        with col_sel_v:
            var_choice = st.selectbox(
                "Chọn biến thể hiện có hoặc tạo mới:",
                options=var_keys + ["➕ [Tạo biến thể mới...]"],
                format_func=lambda x: f"Biến thể {x}: {all_vars_for_norm.get(x, {}).get('variant_name', 'Mặc định')}" if x in all_vars_for_norm else x,
                key="dlg_var_choice"
            )

        if var_choice == "➕ [Tạo biến thể mới...]":
            existing_nums = [int(k[1:]) for k in var_keys if k.startswith("A") and k[1:].isdigit()]
            next_num = max(existing_nums, default=0) + 1
            default_new_id = f"A{next_num}"
            cur_var_id = st.text_input("Mã biến thể mới:", value=default_new_id, key="dlg_var_new_id")
            cur_var_name = st.text_input("Tên mô tả:", value=f"Biến thể điều chỉnh {cur_var_id} cho {var_norm_code}", key="dlg_var_new_name")
            cur_exclude_rep = st.checkbox("❌ Không bao gồm vật tư thay thế (Đơn giá vật tư thay thế = 0 đ)", value=False, key="dlg_var_new_rep")
            base_steps = norm_data.get("work_steps", [])
        else:
            cur_var_id = var_choice
            v_obj = all_vars_for_norm[var_choice]
            cur_var_name = st.text_input("Tên mô tả biến thể:", value=v_obj.get("variant_name", ""), key=f"dlg_var_name_{var_choice}")
            cur_exclude_rep = st.checkbox("❌ Không bao gồm vật tư thay thế (Đơn giá vật tư thay thế = 0 đ)", value=bool(v_obj.get("exclude_replacement", False)), key=f"dlg_var_rep_{var_choice}")
            base_steps = v_obj.get("work_steps", norm_data.get("work_steps", []))

        st.markdown("##### 📝 Chọn các bước thành phần công việc thực hiện:")
        s_edit_data = []
        for idx_s, s in enumerate(base_steps, start=1):
            s_edit_data.append({
                "step_no": s.get("step") or s.get("step_no") or idx_s,
                "step_name": s.get("name") or s.get("step_name") or "",
                "rate_base": float(s.get("rate_base", 0.0)),
                "apply_flag": bool(int(s.get("apply_flag", 1)) == 1)
            })
        df_v_steps = pd.DataFrame(s_edit_data)

        edited_v_steps = st.data_editor(
            df_v_steps,
            column_config={
                "step_no": st.column_config.NumberColumn("Bước", width="small", disabled=True),
                "step_name": st.column_config.TextColumn("Nội dung thành phần công việc", width="large", disabled=True),
                "rate_base": st.column_config.NumberColumn("Hao phí chuẩn (công)", format="%.4f", width="small", disabled=True),
                "apply_flag": st.column_config.CheckboxColumn("Thực hiện", width="small")
            },
            use_container_width=True,
            hide_index=True,
            key=f"dlg_var_steps_editor_{var_norm_code}_{cur_var_id}"
        )

        if st.button(f"💾 Lưu hồ sơ biến thể '{cur_var_id}' vào Dự án", type="primary", use_container_width=True, key=f"dlg_btn_save_var_{var_norm_code}_{cur_var_id}"):
            final_steps = []
            for idx_s, row_s in edited_v_steps.iterrows():
                orig_s = dict(base_steps[idx_s])
                orig_s["apply_flag"] = 1 if row_s["apply_flag"] else 0
                final_steps.append(orig_s)
            mgr.create_or_update_variant(
                norm_code=var_norm_code,
                variant_id=cur_var_id.strip(),
                variant_name=cur_var_name.strip(),
                work_steps=final_steps,
                exclude_replacement=cur_exclude_rep
            )
            st.success(f"✅ Đã lưu biến thể '{cur_var_id}' cho định mức {var_norm_code} thành công!")
            st.rerun()

    with tab_v2:
        st.markdown("##### ⚡ Lọc công tác theo từ khóa & Gán hàng loạt Biến thể")
        f_col1, f_col2, f_col3 = st.columns([4, 4, 4])
        with f_col1:
            batch_norm_code = st.selectbox("Chọn mã định mức cần lọc:", options=all_norm_codes_v, key="dlg_batch_norm_code")
        with f_col2:
            batch_keyword = st.text_input("Từ khóa lọc trong tên công việc (để trống để chọn tất cả):", value="", placeholder="Ví dụ: thay thế, bảo dưỡng...", key="dlg_batch_kw")

        avail_variants = mgr.get_variants_for_norm(batch_norm_code)
        avail_var_keys = list(avail_variants.keys())
        with f_col3:
            chosen_var_to_assign = st.selectbox(
                "Biến thể gán hàng loạt:",
                options=avail_var_keys,
                format_func=lambda x: f"[{x}] {avail_variants[x].get('variant_name', '')}",
                key="dlg_batch_chosen_var"
            )

        matching_items = []
        for it in st.session_state.work_items:
            if it.get("norm_code") == batch_norm_code:
                j_name = str(it.get("job_name", "")).lower()
                kw = batch_keyword.strip().lower()
                if not kw or kw in j_name:
                    matching_items.append(it)

        st.caption(f"Tìm thấy **{len(matching_items)}** công tác khớp với định mức `{batch_norm_code}`" + (f" và từ khóa '{batch_keyword}'" if batch_keyword.strip() else "") + ".")

        if matching_items:
            preview_rows = []
            for it in matching_items:
                preview_rows.append({
                    "STT": it.get("stt"),
                    "Nhóm": it.get("group"),
                    "Tên công tác": it.get("job_name"),
                    "Khối lượng": it.get("volume"),
                    "ĐVT": it.get("unit"),
                    "Biến thể hiện tại": it.get("variant_id", "A0")
                })
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)

            col_btn_apply, col_btn_revert = st.columns([6, 6])
            with col_btn_apply:
                if st.button(f"🚀 Áp dụng Biến thể [{chosen_var_to_assign}] cho TẤT CẢ {len(matching_items)} công tác trên", type="primary", use_container_width=True, key="dlg_btn_batch_apply_var"):
                    for it in matching_items:
                        it["variant_id"] = chosen_var_to_assign
                        it["exclude_replacement"] = avail_variants[chosen_var_to_assign].get("exclude_replacement", False)
                        if "work_steps" in it:
                            del it["work_steps"]
                    mgr.save_work_items(st.session_state.work_items)
                    st.success(f"✅ Đã gán biến thể [{chosen_var_to_assign}] cho {len(matching_items)} công việc!")
                    st.rerun()
            with col_btn_revert:
                if st.button(f"🔄 Khôi phục Biến thể [A0 - Gốc 100%] cho các công tác trên", use_container_width=True, key="dlg_btn_batch_revert_var"):
                    for it in matching_items:
                        it["variant_id"] = "A0"
                        it["exclude_replacement"] = False
                        if "work_steps" in it:
                            del it["work_steps"]
                    mgr.save_work_items(st.session_state.work_items)
                    st.success(f"✅ Đã khôi phục biến thể gốc A0 cho {len(matching_items)} công việc!")
                    st.rerun()


# -----------------------------------------------------------------------------
# TAB 1: KHỐI LƯỢNG TIÊN LƯỢNG (PHẲNG CHUẨN F1/EXCEL - KHÔNG EXPANDER)
# -----------------------------------------------------------------------------
with tab1:
    # 1. ACTION TOOLBAR (THANH CÔNG CỤ NỔI - POPUP DIALOGS)
    col_t1, col_t2, col_t3, col_t4 = st.columns([3, 3, 3, 2])
    with col_t1:
        if st.button("➕ THÊM CÔNG TÁC MỚI", type="primary", use_container_width=True):
            dialog_add_work_item()
    with col_t2:
        if st.button("📥 NHẬP TỪ PAKT / EXCEL", use_container_width=True):
            dialog_quick_intake_pakt()
    with col_t3:
        if st.button("⚙️ QUẢN LÝ BIẾN THỂ (A0/A1...)", use_container_width=True):
            dialog_manage_variants()
    with col_t4:
        if st.button("🔄 LÀM MỚI BẢNG", use_container_width=True):
            st.rerun()

    st.markdown("---")

    # 2. GROUP MANAGEMENT & FILTER BAR
    existing_groups = list(dict.fromkeys([it.get("group", "Nhóm công việc chung").strip() for it in st.session_state.work_items if it.get("group")]))
    if not existing_groups:
        existing_groups = ["Nhóm công tác chung"]

    col_flt_grp, col_flt_mode = st.columns([7, 5])
    with col_flt_grp:
        group_filter_options = ["🌐 [Tất cả các nhóm công tác]"] + existing_groups
        selected_filter_group = st.selectbox(
            "📂 Chọn nhóm công tác hiển thị:",
            options=group_filter_options,
            index=0,
            key="tab1_group_filter"
        )
    with col_flt_mode:
        tab1_view_mode = st.radio(
            "Chế độ hiển thị Master Grid:",
            ["📊 Bảng Đơn Giá Hàng Ngang (F1 Master)", "✏️ Chỉnh Sửa Trực Tiếp Dữ Liệu (F1 Editor)"],
            horizontal=True,
            key="tab1_view_mode_choice"
        )

    # Filter items
    if selected_filter_group == "🌐 [Tất cả các nhóm công tác]":
        current_display_items = st.session_state.work_items
        is_all_groups = True
    else:
        current_display_items = [it for it in st.session_state.work_items if it.get("group", "").strip() == selected_filter_group]
        is_all_groups = False

    # Group Rename & Delete toolbar if single group selected
    if not is_all_groups:
        col_g_ren_txt, col_g_ren_btn, col_g_del_btn = st.columns([6, 3, 3])
        with col_g_ren_txt:
            new_grp_label = st.text_input("Đổi tên nhóm này thành:", value=selected_filter_group, key=f"inp_rename_{selected_filter_group}")
        with col_g_ren_btn:
            st.write("")
            if st.button("✏️ Cập nhật tên nhóm", use_container_width=True, key=f"btn_rename_{selected_filter_group}"):
                if new_grp_label.strip() and new_grp_label.strip() != selected_filter_group:
                    for it in st.session_state.work_items:
                        if it.get("group", "").strip() == selected_filter_group:
                            it["group"] = new_grp_label.strip()
                    mgr.save_work_items(st.session_state.work_items)
                    st.success(f"Đã đổi tên nhóm thành: {new_grp_label.strip()}")
                    st.rerun()
        with col_g_del_btn:
            st.write("")
            if st.button("🗑️ Xóa cả nhóm này", use_container_width=True, key=f"btn_del_grp_{selected_filter_group}"):
                st.session_state.work_items = [it for it in st.session_state.work_items if it.get("group", "").strip() != selected_filter_group]
                for i, it in enumerate(st.session_state.work_items, start=1):
                    it["stt"] = i
                mgr.save_work_items(st.session_state.work_items)
                st.success(f"Đã xóa toàn bộ công tác thuộc nhóm '{selected_filter_group}'!")
                st.rerun()

    # 3. TẦNG 1: MASTER GRID VIEW
    st.markdown(f"#### 📋 TẦNG 1: BẢNG TIÊN LƯỢNG MASTER ({len(current_display_items)} công tác)")

    if not current_display_items:
        st.info("ℹ️ Hiện tại chưa có công tác nào. Hãy nhấn nút **'➕ THÊM CÔNG TÁC MỚI'** hoặc **'📥 NHẬP TỪ PAKT / EXCEL'** ở thanh công cụ phía trên.")
    else:
        if "Bảng Đơn Giá Hàng Ngang" in tab1_view_mode:
            h_rows = []
            total_grid_amount = 0.0
            for item in current_display_items:
                rates = mgr.calculate_item_unit_rate(item, st.session_state.coefficients)
                vol = float(item.get("volume", 0.0))
                var_id = rates.get("variant_id", "")
                var_label = f"[{var_id}]" if var_id and var_id != "A0" else ("Gốc [A0]" if not rates.get("is_quote") else "-")
                tt_vl = round(vol * rates["dg_vl"])
                tt_nc = round(vol * rates["dg_nc"])
                tt_m = round(vol * rates["dg_m"])
                h_rows.append({
                    "STT": item.get("stt"),
                    "Nhóm": item.get("group", ""),
                    "Mã hiệu": item.get("norm_code"),
                    "Biến thể": var_label,
                    "Bản chất kỹ thuật": item.get("tech_profile") or mgr.get_default_tech_profile(item.get("norm_code", "")),
                    "Tên công tác (PAKT)": item.get("job_name"),
                    "ĐVT": item.get("unit"),
                    "Khối lượng": vol,
                    "ĐG Vật liệu (đ)": rates["dg_vl"],
                    "ĐG Nhân công (đ)": rates["dg_nc"],
                    "ĐG Ca máy (đ)": rates["dg_m"],
                    "ĐG Trực tiếp T (đ)": rates["dg_tt"],
                    "Thành tiền VL (đ)": tt_vl,
                    "Thành tiền NC (đ)": tt_nc,
                    "Thành tiền Máy (đ)": tt_m,
                    "ĐG Báo giá (đ)": rates["unit_price_quote"],
                    "ĐG Đầy đủ trước thuế (đ)": rates["dg_truoc_thue"],
                    "Thành tiền trước thuế (đ)": rates["thanh_tien_truoc_thue"]
                })
                total_grid_amount += rates["thanh_tien_truoc_thue"]

            df_h = pd.DataFrame(h_rows)
            st.dataframe(
                df_h.style.format({
                    "Khối lượng": "{:,.2f}",
                    "ĐG Vật liệu (đ)": "{:,.0f}",
                    "ĐG Nhân công (đ)": "{:,.0f}",
                    "ĐG Ca máy (đ)": "{:,.0f}",
                    "ĐG Trực tiếp T (đ)": "{:,.0f}",
                    "Thành tiền VL (đ)": "{:,.0f}",
                    "Thành tiền NC (đ)": "{:,.0f}",
                    "Thành tiền Máy (đ)": "{:,.0f}",
                    "ĐG Báo giá (đ)": "{:,.0f}",
                    "ĐG Đầy đủ trước thuế (đ)": "{:,.0f}",
                    "Thành tiền trước thuế (đ)": "{:,.0f}"
                }),
                use_container_width=True,
                hide_index=True
            )

            st.markdown(f"""
            <div class="metric-card">
                <b>TỔNG CỘNG GIÁ TRỊ TRƯỚC THUẾ:</b> 
                <span style="font-size:20px; color:#002060; font-weight:bold;">{total_grid_amount:,.0f} đồng</span> 
                <i>(Bao gồm Đầy đủ Chi phí chung {st.session_state.coefficients['c']['rate']}%, TT {st.session_state.coefficients['tt']['rate']}%, Thu nhập chịu thuế tính trước {st.session_state.coefficients['tl']['rate']}%)</i>
            </div>
            """, unsafe_allow_html=True)

        else:
            # Editor Mode
            df_edit = pd.DataFrame(current_display_items)
            if "tech_profile" not in df_edit.columns:
                df_edit["tech_profile"] = df_edit["norm_code"].apply(lambda c: mgr.get_default_tech_profile(str(c)))
            if "is_quote" not in df_edit.columns:
                df_edit["is_quote"] = False
            if "unit_price_quote" not in df_edit.columns:
                df_edit["unit_price_quote"] = 0.0

            disp_cols = ["stt", "group", "norm_code", "variant_id", "tech_profile", "job_name", "description", "volume", "unit", "unit_price_quote"]
            for c_req in disp_cols:
                if c_req not in df_edit.columns:
                    df_edit[c_req] = ""

            edited_df = st.data_editor(
                df_edit[disp_cols],
                column_config={
                    "stt": st.column_config.NumberColumn("STT", width="small", disabled=True),
                    "group": st.column_config.TextColumn("Nhóm công tác", width="medium", required=True),
                    "norm_code": st.column_config.TextColumn("Mã ĐM / BG", width="small", required=True),
                    "variant_id": st.column_config.TextColumn("Biến thể", width="small"),
                    "tech_profile": st.column_config.TextColumn("Bản chất kỹ thuật", width="medium"),
                    "job_name": st.column_config.TextColumn("Danh mục công việc (Theo PAKT)", width="large", required=True),
                    "description": st.column_config.TextColumn("Mô tả công việc thực tế", width="medium"),
                    "volume": st.column_config.NumberColumn("Khối lượng", format="%.2f", min_value=0.0, step=1.0, required=True),
                    "unit": st.column_config.TextColumn("Đơn vị tính", width="small", required=True),
                    "unit_price_quote": st.column_config.NumberColumn("ĐG Báo giá trước thuế (đ)", format="%d", min_value=0.0, step=500000.0)
                },
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key="tab1_master_editor"
            )

            col_s1, col_s2 = st.columns([6, 6])
            with col_s1:
                if st.button("💾 LƯU TẤT CẢ THAY ĐỔI VÀO DỰ ÁN", key="btn_save_master_edit", type="primary", use_container_width=True):
                    updated_records = edited_df.to_dict("records")
                    old_items_map = {it.get("stt"): it for it in current_display_items}
                    
                    merged_records = []
                    for r in updated_records:
                        orig = old_items_map.get(r.get("stt"))
                        if orig:
                            merged_r = dict(orig)
                            merged_r.update(r)
                        else:
                            merged_r = dict(r)
                            c_str = str(r.get("norm_code", "")).strip()
                            if c_str.startswith("BG") or c_str.startswith("TT"):
                                merged_r["is_quote"] = True
                            else:
                                merged_r["is_quote"] = False
                        merged_r["volume"] = float(merged_r.get("volume", 0.0) or 0.0)
                        merged_r["unit_price_quote"] = float(merged_r.get("unit_price_quote", 0.0) or 0.0)
                        merged_records.append(merged_r)

                    if is_all_groups:
                        st.session_state.work_items = merged_records
                    else:
                        other_items = [it for it in st.session_state.work_items if it.get("group", "").strip() != selected_filter_group]
                        st.session_state.work_items = other_items + merged_records

                    for i, it in enumerate(st.session_state.work_items, start=1):
                        it["stt"] = i

                    mgr.save_work_items(st.session_state.work_items)
                    st.success("✅ Đã lưu toàn bộ các thay đổi vào CSDL Dự án thành công!")
                    st.rerun()

            with col_s2:
                # Delete single item directly
                del_opts = [it["stt"] for it in current_display_items]
                col_d_sel, col_d_btn = st.columns([8, 4])
                with col_d_sel:
                    item_to_del_stt = st.selectbox(
                        "Xóa một công tác:",
                        options=del_opts,
                        format_func=lambda x: f"STT {x}: {[it.get('job_name') for it in current_display_items if it.get('stt') == x][0]}",
                        key="tab1_del_single_stt"
                    )
                with col_d_btn:
                    st.write("")
                    if st.button("🗑️ Xác nhận xóa", key="tab1_btn_del_single", use_container_width=True):
                        st.session_state.work_items = [it for it in st.session_state.work_items if it.get("stt") != item_to_del_stt]
                        for i, it in enumerate(st.session_state.work_items, start=1):
                            it["stt"] = i
                        mgr.save_work_items(st.session_state.work_items)
                        st.success(f"Đã xóa công tác STT {item_to_del_stt}!")
                        st.rerun()

    # 4. TẦNG 2: F1 DETAIL INSPECTOR (PHÂN RÃ HAO PHÍ & BƯỚC CÔNG VIỆC)
    st.markdown("---")
    st.markdown("#### 🔍 TẦNG 2: F1 DETAIL INSPECTOR - PHÂN RÃ CHI TIẾT HAO PHÍ & BƯỚC CÔNG VIỆC")
    st.caption("Bảng dưới tự động đồng bộ hiển thị chi tiết Vật liệu, Nhân công, Máy thi công và các bước quy trình của công tác đang chọn mà không cần mở expander.")

    if current_display_items:
        avail_stt_inspect = [it["stt"] for it in current_display_items]
        sel_insp_stt = st.selectbox(
            "👉 Chọn công tác để soi chi tiết hao phí (Chuẩn F1/G8):",
            options=avail_stt_inspect,
            format_func=lambda x: f"STT {x}: [{next(it.get('norm_code') for it in current_display_items if it['stt'] == x)}] {next(it.get('job_name') for it in current_display_items if it['stt'] == x)} (KL: {next(it.get('volume') for it in current_display_items if it['stt'] == x)} {next(it.get('unit') for it in current_display_items if it['stt'] == x)})",
            key="tab1_inspect_stt_sel"
        )

        insp_item = next(it for it in current_display_items if it["stt"] == sel_insp_stt)
        insp_code = insp_item.get("norm_code", "")
        insp_is_quote = insp_item.get("is_quote", False)
        insp_rates = mgr.calculate_item_unit_rate(insp_item, st.session_state.coefficients)

        # Overview Header Card for Inspected Item
        st.markdown(f"""
        <div style="background-color: #EBF3FB; border-left: 5px solid #002060; padding: 12px 18px; border-radius: 6px; margin-bottom: 12px;">
            <b>CÔNG TÁC ĐANG SOI:</b> <span style="font-size:16px; color:#002060; font-weight:bold;">STT {insp_item.get('stt')} - {insp_item.get('job_name')}</span><br/>
            📌 <b>Mã hiệu:</b> <code>{insp_code}</code> | 
            <b>Biến thể:</b> <span style="background-color:#002060; color:white; padding:2px 8px; border-radius:4px;">{insp_rates.get('variant_id', 'A0')}</span> - {insp_rates.get('variant_name', 'Mặc định')} | 
            <b>Khối lượng:</b> <b>{insp_item.get('volume', 0.0):,.2f} {insp_item.get('unit', '')}</b> | 
            <b>Bản chất kỹ thuật:</b> <i>{insp_item.get('tech_profile') or mgr.get_default_tech_profile(insp_code)}</i>
        </div>
        """, unsafe_allow_html=True)

        if not insp_is_quote and mgr.get_norm(insp_code):
            # Case Norm Item: 2 Columns (Steps & Resource Breakdown)
            col_insp_steps, col_insp_res = st.columns([5, 7])

            eff_info = mgr.get_effective_item_resources(insp_item)
            curr_steps = eff_info["raw_steps"]
            avail_vars_for_insp = mgr.get_variants_for_norm(insp_code)

            with col_insp_steps:
                st.markdown("##### ⚙️ Các bước quy trình & Tỷ lệ hao phí $k$")
                
                # Switch variant for this item directly
                col_iv1, col_iv2 = st.columns([8, 4])
                with col_iv1:
                    cur_v_key = insp_item.get("variant_id", "A0")
                    sel_v_for_item = st.selectbox(
                        "Gán biến thể cho công tác này:",
                        options=list(avail_vars_for_insp.keys()),
                        index=list(avail_vars_for_insp.keys()).index(cur_v_key) if cur_v_key in avail_vars_for_insp else 0,
                        format_func=lambda x: f"[{x}] {avail_vars_for_insp[x].get('variant_name', '')}",
                        key=f"insp_pick_var_{sel_insp_stt}"
                    )
                with col_iv2:
                    st.write("")
                    if st.button("Áp dụng", key=f"btn_apply_v_{sel_insp_stt}", use_container_width=True):
                        insp_item["variant_id"] = sel_v_for_item
                        insp_item["exclude_replacement"] = avail_vars_for_insp[sel_v_for_item].get("exclude_replacement", False)
                        if "work_steps" in insp_item:
                            del insp_item["work_steps"]
                        mgr.save_work_items(st.session_state.work_items)
                        st.success(f"Đã gán biến thể [{sel_v_for_item}]!")
                        st.rerun()

                # Step Table
                step_editor_data = []
                for idx_s, s in enumerate(curr_steps, start=1):
                    step_editor_data.append({
                        "step_no": s.get("step") or s.get("step_no") or idx_s,
                        "step_name": s.get("name") or s.get("step_name") or "",
                        "rate_base": float(s.get("rate_base", 0.0)),
                        "apply_flag": bool(int(s.get("apply_flag", 1)) == 1)
                    })
                df_step_edit = pd.DataFrame(step_editor_data)

                edited_steps = st.data_editor(
                    df_step_edit,
                    column_config={
                        "step_no": st.column_config.NumberColumn("Bước", width="small", disabled=True),
                        "step_name": st.column_config.TextColumn("Tên bước công việc", width="large", disabled=True),
                        "rate_base": st.column_config.NumberColumn("Hao phí chuẩn", format="%.4f", width="small", disabled=True),
                        "apply_flag": st.column_config.CheckboxColumn("Làm", width="small")
                    },
                    use_container_width=True,
                    hide_index=True,
                    key=f"insp_step_editor_{sel_insp_stt}"
                )

                col_st_save, col_st_reset = st.columns([6, 6])
                with col_st_save:
                    if st.button(f"💾 Lưu bước cho STT {sel_insp_stt}", key=f"btn_save_steps_{sel_insp_stt}", type="primary", use_container_width=True):
                        new_steps_list = []
                        for idx_s, row_s in edited_steps.iterrows():
                            orig_s = dict(curr_steps[idx_s])
                            orig_s["apply_flag"] = 1 if row_s["apply_flag"] else 0
                            new_steps_list.append(orig_s)
                        insp_item["work_steps"] = new_steps_list
                        mgr.save_work_items(st.session_state.work_items)
                        st.success(f"Đã lưu các bước cho STT {sel_insp_stt}!")
                        st.rerun()
                with col_st_reset:
                    if st.button(f"🔄 Khôi phục 100% gốc", key=f"btn_reset_steps_{sel_insp_stt}", use_container_width=True):
                        if "work_steps" in insp_item:
                            del insp_item["work_steps"]
                        mgr.save_work_items(st.session_state.work_items)
                        st.success("Đã khôi phục định mức gốc!")
                        st.rerun()

            with col_insp_res:
                st.markdown(f"##### 📊 Phân rã hao phí tài nguyên (Hệ số $k$ = {eff_info.get('k_step', 1.0):.1%})")
                
                # Fetch price dicts for lookup
                all_prices = mgr.get_resource_prices()
                mat_price_map = {m["code"]: m.get("price_notice", m.get("price_current", 0)) for m in all_prices.get("materials", [])}
                lab_price_map = {l["code"]: l.get("price_current", 0) for l in all_prices.get("labor", [])}
                mac_price_map = {m["code"]: m.get("price_current", 0) for m in all_prices.get("machines", [])}

                res_rows = []
                for m in eff_info.get("materials", []):
                    p = mat_price_map.get(m["code"], 0)
                    r_val = float(m.get("rate", 0.0))
                    res_rows.append({
                        "Loại": "Vật liệu",
                        "Mã hiệu": m["code"],
                        "Tên tài nguyên": m["name"] + (" (Đã loại trừ thay thế)" if m.get("excluded") else ""),
                        "ĐVT": m["unit"],
                        "Định mức": r_val,
                        "Đơn giá (đ)": p,
                        "Thành tiền (đ)": round(r_val * p)
                    })
                for l in eff_info.get("labor", []):
                    p = lab_price_map.get(l["code"], 0)
                    r_val = float(l.get("rate", 0.0))
                    res_rows.append({
                        "Loại": "Nhân công",
                        "Mã hiệu": l["code"],
                        "Tên tài nguyên": l["name"],
                        "ĐVT": l["unit"],
                        "Định mức": r_val,
                        "Đơn giá (đ)": p,
                        "Thành tiền (đ)": round(r_val * p)
                    })
                for mc in eff_info.get("machines", []):
                    p = mac_price_map.get(mc["code"], 0)
                    r_val = float(mc.get("rate", 0.0))
                    res_rows.append({
                        "Loại": "Ca máy",
                        "Mã hiệu": mc["code"],
                        "Tên tài nguyên": mc["name"],
                        "ĐVT": mc["unit"],
                        "Định mức": r_val,
                        "Đơn giá (đ)": p,
                        "Thành tiền (đ)": round(r_val * p)
                    })

                df_res = pd.DataFrame(res_rows)
                st.dataframe(
                    df_res.style.format({
                        "Định mức": "{:,.4f}",
                        "Đơn giá (đ)": "{:,.0f}",
                        "Thành tiền (đ)": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                # Rate summary pills
                st.markdown(f"""
                <div style="display:flex; gap:10px; flex-wrap:wrap; margin-top:8px;">
                    <div style="background:#FFF3E0; padding:6px 12px; border-radius:4px; border:1px solid #FFE0B2;">
                        <b>VL:</b> {insp_rates['dg_vl']:,.0f} đ
                    </div>
                    <div style="background:#E8F5E9; padding:6px 12px; border-radius:4px; border:1px solid #C8E6C9;">
                        <b>NC:</b> {insp_rates['dg_nc']:,.0f} đ
                    </div>
                    <div style="background:#E1F5FE; padding:6px 12px; border-radius:4px; border:1px solid #B3E5FC;">
                        <b>Máy:</b> {insp_rates['dg_m']:,.0f} đ
                    </div>
                    <div style="background:#EDE7F6; padding:6px 12px; border-radius:4px; border:1px solid #D1C4E9;">
                        <b>ĐG Trực tiếp T:</b> {insp_rates['dg_tt']:,.0f} đ
                    </div>
                    <div style="background:#002060; color:white; padding:6px 14px; border-radius:4px;">
                        <b>ĐG Trước thuế:</b> {insp_rates['dg_truoc_thue']:,.0f} đ/{insp_item.get('unit')}
                    </div>
                </div>
                """, unsafe_allow_html=True)

        else:
            # Case Quote Item
            st.markdown("##### 📝 Chi tiết Báo giá trọn gói / Tạm tính")
            q_cols = st.columns([4, 4, 4])
            with q_cols[0]:
                st.metric("Đơn giá báo giá trước thuế", f"{insp_rates['unit_price_quote']:,.0f} đ/{insp_item.get('unit')}")
            with q_cols[1]:
                st.metric("Khối lượng dự án", f"{insp_item.get('volume', 0.0):,.2f} {insp_item.get('unit')}")
            with q_cols[2]:
                st.metric("Thành tiền trước thuế", f"{insp_rates['thanh_tien_truoc_thue']:,.0f} đ")
            st.info(f"📌 **Căn cứ báo giá:** {insp_item.get('quote_basis') or 'Theo báo giá nhà sản xuất / nhà cung cấp'}")

# TAB 2: QUẢN LÝ ĐƠN GIÁ (VẬT LIỆU, NHÂN CÔNG, MÁY THI CÔNG & PHÁP LÝ)
# -----------------------------------------------------------------------------
with tab2:
    st.markdown("#### 💰 QUẢN LÝ ĐƠN GIÁ TÀI NGUYÊN & CƠ SỞ PHÁP LÝ")
    st.caption("Cột Khối lượng được tính toán tự động từ bảng Tiên lượng. Người dùng tùy chỉnh Đơn giá và Căn cứ pháp lý.")

    # Calculate real-time resource quantities from work_items taking into account item-level step adjustments
    mat_qty = {}
    lab_qty = {}
    mac_qty = {}
    for item in st.session_state.work_items:
        c = item.get("norm_code")
        v = float(item.get("volume", 0.0))
        norm = mgr.get_norm(c)
        if not norm: continue
        u_str = norm.get("unit", "")
        div = mgr.get_unit_divisor(u_str, item.get("unit", ""))
        eff = mgr.get_effective_item_resources(item)
        for m in eff.get("materials", []):
            if m.get("code"): mat_qty[m["code"]] = mat_qty.get(m["code"], 0.0) + (v * float(m.get("rate", 0.0)) / div)
        for l in eff.get("labor", []):
            if l.get("code"): lab_qty[l["code"]] = lab_qty.get(l["code"], 0.0) + (v * float(l.get("rate", 0.0)) / div)
        for mac in eff.get("machines", []):
            if mac.get("code"): mac_qty[mac["code"]] = mac_qty.get(mac["code"], 0.0) + (v * float(mac.get("rate", 0.0)) / div)

    sub_t1, sub_t2, sub_t3, sub_t4 = st.tabs([
        "🧱 1. Vật Liệu",
        "👷 2. Nhân Công",
        "🚜 3. Máy Thi Công",
        "🎯 4. Bản Chất Kỹ Thuật & Đơn Giá Thống Nhất"
    ])
    prices = mgr.get_resource_prices()

    with sub_t1:
        st.markdown("##### 1. BẢNG TỔNG HỢP VẬT LIỆU")
        vl_mode = st.radio(
            "Chế độ hiển thị Vật liệu:",
            ["👁️ 1. Bảng giá hiện tại (Theo chuẩn PAKT & EVN)", "✏️ 2. Chỉnh sửa đơn giá & Căn cứ pháp lý"],
            horizontal=True,
            key="vl_view_mode_choice"
        )
        df_vl = pd.DataFrame(prices.get("materials", []))
        if not df_vl.empty:
            df_vl["qty_calc"] = df_vl["code"].apply(lambda x: round(mat_qty.get(x, 0.0), 3))
            if "legal_basis" not in df_vl.columns: df_vl["legal_basis"] = df_vl.get("notes", "Công bố giá LS / Báo giá NSX")
            if "source" not in df_vl.columns: df_vl["source"] = ""

            if "1. Bảng giá hiện tại" in vl_mode:
                # View Mode matching user's image:
                # STT | Mã hiệu | Tên vật tư | Đơn vị | Nguồn mua | Khối lượng | Giá hiện tại | Thành tiền | Căn cứ tính giá
                rows_current = []
                main_mat_total = 0.0
                for idx, r in df_vl.iterrows():
                    code = r["code"]
                    is_other = (code == "V00750" or r["unit"] == "%")
                    if not is_other:
                        q = float(r["qty_calc"])
                        if q > 0:
                            p = float(r.get("price_notice", 0))
                            main_mat_total += round(q * p)
                
                total_vl_all = 0.0
                stt_counter = 1
                for idx, r in df_vl.iterrows():
                    code = r["code"]
                    is_other = (code == "V00750" or r["unit"] == "%")
                    if is_other:
                        if main_mat_total > 0:
                            pct = float(r.get("rate", 2.0))
                            tt = round(main_mat_total * (pct / 100.0))
                            total_vl_all += tt
                            rows_current.append({
                                "STT": stt_counter,
                                "Mã hiệu": code,
                                "Tên vật tư": r["name"],
                                "Đơn vị": r["unit"],
                                "Nguồn mua": r.get("source", ""),
                                "Khối lượng": None,
                                "Giá hiện tại (đ)": None,
                                "Thành tiền (đ)": tt,
                                "Căn cứ tính giá": r.get("legal_basis", f"{pct}% vật liệu chính theo định mức")
                            })
                            stt_counter += 1
                    else:
                        q = float(r["qty_calc"])
                        if q > 0:
                            p = float(r.get("price_notice", 0))
                            tt = round(q * p)
                            total_vl_all += tt
                            rows_current.append({
                                "STT": stt_counter,
                                "Mã hiệu": code,
                                "Tên vật tư": r["name"],
                                "Đơn vị": r["unit"],
                                "Nguồn mua": r.get("source", ""),
                                "Khối lượng": q,
                                "Giá hiện tại (đ)": p,
                                "Thành tiền (đ)": tt,
                                "Căn cứ tính giá": r.get("legal_basis", "Công bố giá LS / Báo giá NSX")
                            })
                            stt_counter += 1
                
                df_vl_disp = pd.DataFrame(rows_current)
                st.dataframe(
                    df_vl_disp.style.format({
                        "Khối lượng": lambda x: f"{x:,.2f}" if pd.notnull(x) else "",
                        "Giá hiện tại (đ)": lambda x: f"{x:,.0f}" if pd.notnull(x) else "",
                        "Thành tiền (đ)": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                col_sum, col_dl = st.columns([6, 4])
                with col_sum:
                    st.markdown(f"""
                    <div class="metric-card" style="margin-top:5px;">
                        <b>Tổng vật liệu:</b> 
                        <span style="font-size:18px; color:#002060; font-weight:bold; margin-left:10px;">{total_vl_all:,.0f} đồng</span>
                    </div>
                    """, unsafe_allow_html=True)
                with col_dl:
                    materials_payload = []
                    for idx, r in df_vl.iterrows():
                        q_val = float(r["qty_calc"])
                        is_other = (r["code"] == "V00750" or r["unit"] == "%")
                        if q_val > 0 or (is_other and main_mat_total > 0):
                            m_copy = r.to_dict()
                            m_copy["total_qty"] = q_val
                            materials_payload.append(m_copy)
                    excel_bytes_vl = create_single_summary_sheet_excel("material", project_name, category_name, materials_payload, mgr)
                    st.download_button(
                        "📥 Tải Bảng Tổng Hợp Vật Liệu (.xlsx)",
                        data=excel_bytes_vl,
                        file_name="TH_VatLieu_GiaHienTai.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                edited_vl = df_vl
            else:
                df_vl_edit_source = df_vl[df_vl["qty_calc"] > 0] if not df_vl.empty else df_vl
                edited_vl = st.data_editor(
                    df_vl_edit_source[["code", "name", "unit", "qty_calc", "source", "price_base", "price_notice", "legal_basis"]],
                    column_config={
                        "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                        "name": st.column_config.TextColumn("Tên vật tư", disabled=True),
                        "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                        "qty_calc": st.column_config.NumberColumn("Khối lượng (Tiên lượng)", format="%.3f", disabled=True),
                        "source": st.column_config.TextColumn("Nguồn mua"),
                        "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                        "price_notice": st.column_config.NumberColumn("Giá hiện tại / Thông báo (đ)", format="%d"),
                        "legal_basis": st.column_config.TextColumn("Căn cứ tính giá / Pháp lý", width="large")
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="editor_vl"
                )
                if st.button("💾 Lưu Cập Nhật Bảng Giá Vật Liệu", key="save_vl_btn", type="primary"):
                    mgr.update_resource_prices({"materials": edited_vl.to_dict("records")})
                    st.success("Đã cập nhật bảng giá vật liệu thành công!")
                    st.rerun()

    with sub_t2:
        st.markdown("##### 2. BẢNG TỔNG HỢP NHÂN CÔNG")
        nc_mode = st.radio(
            "Chế độ hiển thị Nhân công:",
            ["👁️ 1. Bảng giá hiện tại (Theo chuẩn PAKT & EVN)", "✏️ 2. Chỉnh sửa đơn giá & Căn cứ pháp lý"],
            horizontal=True,
            key="nc_view_mode_choice"
        )
        df_nc = pd.DataFrame(prices.get("labor", []))
        if not df_nc.empty:
            df_nc["qty_calc"] = df_nc["code"].apply(lambda x: round(lab_qty.get(x, 0.0), 3))
            if "legal_basis" not in df_nc.columns: df_nc["legal_basis"] = "5898/SXD-QLXD ngày 02/7/2026 của SXD tỉnh Lâm Đồng"

            if "1. Bảng giá hiện tại" in nc_mode:
                # View Mode matching user's image:
                # STT | Mã hiệu | Tên vật tư (Tên nhân công) | Đơn vị | Khối lượng | Giá hiện tại | Thành tiền | Căn cứ tính giá
                rows_nc_current = []
                total_nc_all = 0.0
                stt_nc_counter = 1
                for idx, r in df_nc.iterrows():
                    q = float(r["qty_calc"])
                    if q <= 0:
                        continue
                    p = float(r.get("price_current", 0))
                    tt = round(q * p)
                    total_nc_all += tt
                    rows_nc_current.append({
                        "STT": stt_nc_counter,
                        "Mã hiệu": r["code"],
                        "Tên nhân công": r["name"],
                        "Đơn vị": r["unit"],
                        "Khối lượng": q,
                        "Giá hiện tại (đ)": p,
                        "Thành tiền (đ)": tt,
                        "Căn cứ tính giá": r.get("legal_basis", "5898/SXD-QLXD")
                    })
                    stt_nc_counter += 1
                
                df_nc_disp = pd.DataFrame(rows_nc_current)
                st.dataframe(
                    df_nc_disp.style.format({
                        "Khối lượng": "{:,.2f}",
                        "Giá hiện tại (đ)": "{:,.0f}",
                        "Thành tiền (đ)": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                col_sum_nc, col_dl_nc = st.columns([6, 4])
                with col_sum_nc:
                    st.markdown(f"""
                    <div class="metric-card" style="margin-top:5px;">
                        <b>Cộng nhân công:</b> 
                        <span style="font-size:18px; color:#002060; font-weight:bold; margin-left:10px;">{total_nc_all:,.0f} đồng</span>
                    </div>
                    """, unsafe_allow_html=True)
                with col_dl_nc:
                    labor_payload = []
                    for idx, r in df_nc.iterrows():
                        q_val = float(r["qty_calc"])
                        if q_val > 0:
                            l_copy = r.to_dict()
                            l_copy["total_qty"] = q_val
                            labor_payload.append(l_copy)
                    excel_bytes_nc = create_single_summary_sheet_excel("labor", project_name, category_name, labor_payload, mgr)
                    st.download_button(
                        "📥 Tải Bảng Tổng Hợp Nhân Công (.xlsx)",
                        data=excel_bytes_nc,
                        file_name="TH_NhanCong_GiaHienTai.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                edited_nc = df_nc
            else:
                df_nc_edit_source = df_nc[df_nc["qty_calc"] > 0] if not df_nc.empty else df_nc
                edited_nc = st.data_editor(
                    df_nc_edit_source[["code", "name", "unit", "qty_calc", "price_base", "price_current", "legal_basis"]],
                    column_config={
                        "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                        "name": st.column_config.TextColumn("Tên nhân công", disabled=True),
                        "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                        "qty_calc": st.column_config.NumberColumn("Khối lượng (Tiên lượng)", format="%.3f", disabled=True),
                        "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                        "price_current": st.column_config.NumberColumn("Giá hiện tại (đ)", format="%d"),
                        "legal_basis": st.column_config.TextColumn("Căn cứ tính giá / Pháp lý", width="large")
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="editor_nc"
                )
                if st.button("💾 Lưu Cập Nhật Bảng Giá Nhân Công", key="save_nc_btn", type="primary"):
                    mgr.update_resource_prices({"labor": edited_nc.to_dict("records")})
                    st.success("Đã cập nhật bảng giá nhân công thành công!")
                    st.rerun()

    with sub_t3:
        st.markdown("##### 3. BẢNG TỔNG HỢP MÁY THI CÔNG")
        m_mode = st.radio(
            "Chế độ hiển thị Máy thi công:",
            ["👁️ 1. Bảng giá hiện tại (Theo chuẩn PAKT & EVN)", "✏️ 2. Chỉnh sửa đơn giá & Căn cứ pháp lý"],
            horizontal=True,
            key="m_view_mode_choice"
        )
        df_m = pd.DataFrame(prices.get("machines", []))
        if not df_m.empty:
            df_m["qty_calc"] = df_m["code"].apply(lambda x: round(mac_qty.get(x, 0.0), 3))
            if "legal_basis" not in df_m.columns: df_m["legal_basis"] = "Thông tư 37/2026/TT-BXD của Bộ Xây dựng"

            if "1. Bảng giá hiện tại" in m_mode:
                # View Mode matching user's image:
                # STT | Mã hiệu | Tên vật tư (Tên máy) | Đơn vị | Khối lượng | Giá hiện tại | Thành tiền | Căn cứ tính giá
                rows_m_current = []
                total_m_all = 0.0
                stt_m_counter = 1
                for idx, r in df_m.iterrows():
                    q = float(r["qty_calc"])
                    if q <= 0:
                        continue
                    p = float(r.get("price_current", r.get("price_notice", 0) + r.get("price_diff", 0)))
                    tt = round(q * p)
                    total_m_all += tt
                    rows_m_current.append({
                        "STT": stt_m_counter,
                        "Mã hiệu": r["code"],
                        "Tên máy thi công": r["name"],
                        "Đơn vị": r["unit"],
                        "Khối lượng": q,
                        "Giá hiện tại (đ)": p,
                        "Thành tiền (đ)": tt,
                        "Căn cứ tính giá": r.get("legal_basis", "TT 37/2026/TT-BXD")
                    })
                    stt_m_counter += 1
                
                df_m_disp = pd.DataFrame(rows_m_current)
                st.dataframe(
                    df_m_disp.style.format({
                        "Khối lượng": "{:,.2f}",
                        "Giá hiện tại (đ)": "{:,.0f}",
                        "Thành tiền (đ)": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                col_sum_m, col_dl_m = st.columns([6, 4])
                with col_sum_m:
                    st.markdown(f"""
                    <div class="metric-card" style="margin-top:5px;">
                        <b>Cộng Máy:</b> 
                        <span style="font-size:18px; color:#002060; font-weight:bold; margin-left:10px;">{total_m_all:,.0f} đồng</span>
                    </div>
                    """, unsafe_allow_html=True)
                with col_dl_m:
                    machines_payload = []
                    for idx, r in df_m.iterrows():
                        q_val = float(r["qty_calc"])
                        if q_val > 0:
                            mac_copy = r.to_dict()
                            mac_copy["total_qty"] = q_val
                            machines_payload.append(mac_copy)
                    excel_bytes_m = create_single_summary_sheet_excel("machine", project_name, category_name, machines_payload, mgr)
                    st.download_button(
                        "📥 Tải Bảng Tổng Hợp Máy Thi Công (.xlsx)",
                        data=excel_bytes_m,
                        file_name="TH_MayThiCong_GiaHienTai.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                edited_m = df_m
            else:
                edited_m = st.data_editor(
                    df_m[["code", "name", "unit", "qty_calc", "price_base", "price_notice", "price_diff", "price_current", "legal_basis"]],
                    column_config={
                        "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                        "name": st.column_config.TextColumn("Tên máy thi công", disabled=True),
                        "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                        "qty_calc": st.column_config.NumberColumn("Khối lượng (Tiên lượng)", format="%.3f", disabled=True),
                        "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                        "price_notice": st.column_config.NumberColumn("Giá thông báo (đ)", format="%d"),
                        "price_diff": st.column_config.NumberColumn("Bù giá (đ)", format="%d"),
                        "price_current": st.column_config.NumberColumn("Giá hiện tại (đ)", format="%d"),
                        "legal_basis": st.column_config.TextColumn("Căn cứ tính giá / Pháp lý", width="large")
                    },
                    use_container_width=True,
                    hide_index=True,
                    key="editor_m"
                )
                if st.button("💾 Lưu Cập Nhật Bảng Giá Máy Thi Công", key="save_m_btn", type="primary"):
                    mgr.update_resource_prices({"machines": edited_m.to_dict("records")})
                    st.success("Đã cập nhật bảng giá máy thi công thành công!")
                    st.rerun()

        # Tra cứu Bảng giá ca máy Lâm Đồng năm 2026 (QĐ 21/QĐ-SXD)
        st.divider()
        with st.expander("🏛️ TRA CỨU BẢNG GIÁ CA MÁY LÂM ĐỒNG NĂM 2026 (QĐ 21/QĐ-SXD)", expanded=False):
            st.markdown("##### 📖 Cơ sở dữ liệu 698 máy thi công tỉnh Lâm Đồng năm 2026")
            st.caption("Căn cứ Quyết định số 21/QĐ-SXD ngày 14/01/2026 của Sở Xây dựng Lâm Đồng.")
            
            lam_dong_machines = mgr.get_province_machines("lam_dong")
            if lam_dong_machines:
                search_kw = st.text_input("🔍 Tìm kiếm theo mã máy hoặc tên máy:", "", key="search_ld_mac")
                filtered_mac = lam_dong_machines
                if search_kw:
                    kw = search_kw.lower()
                    filtered_mac = [m for m in lam_dong_machines if kw in m.get("code", "").lower() or kw in m.get("name", "").lower()]
                
                rows_show = []
                for m in filtered_mac:
                    rows_show.append({
                        "Mã hiệu": m.get("code"),
                        "Tên máy & Quy cách": m.get("name"),
                        "ĐVT": m.get("unit", "ca"),
                        "Thợ ĐK": m.get("driver", ""),
                        "Nguyên giá (1000đ)": m.get("original_price_1k", 0),
                        "CP Nhiên liệu (đ)": m.get("fuel_cost", 0),
                        "Lương thợ Vùng IV (đ)": m.get("labor_cost", {}).get("vung_4", 0),
                        "Giá ca máy Vùng II (đ)": m.get("price", {}).get("vung_2", 0),
                        "Giá ca máy Vùng III (đ)": m.get("price", {}).get("vung_3", 0),
                        "Giá ca máy Vùng IV (đ - MẶC ĐỊNH)": m.get("price", {}).get("vung_4", 0)
                    })
                
                df_show = pd.DataFrame(rows_show)
                num_cols = [
                    "Nguyên giá (1000đ)", "CP Nhiên liệu (đ)", "Lương thợ Vùng IV (đ)",
                    "Giá ca máy Vùng II (đ)", "Giá ca máy Vùng III (đ)", "Giá ca máy Vùng IV (đ - MẶC ĐỊNH)"
                ]
                for col in num_cols:
                    if col in df_show.columns:
                        df_show[col] = pd.to_numeric(df_show[col], errors="coerce").fillna(0).astype("int64")

                st.dataframe(
                    df_show.style.format({
                        "Nguyên giá (1000đ)": "{:,.0f}",
                        "CP Nhiên liệu (đ)": "{:,.0f}",
                        "Lương thợ Vùng IV (đ)": "{:,.0f}",
                        "Giá ca máy Vùng II (đ)": "{:,.0f}",
                        "Giá ca máy Vùng III (đ)": "{:,.0f}",
                        "Giá ca máy Vùng IV (đ - MẶC ĐỊNH)": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                
                col_exp1, col_exp2 = st.columns([5, 5])
                with col_exp1:
                    excel_path_ld = Path(r"D:\TaskApp_kiet\LapDuToanApp\gia_ca_may_lam_dong_2026.xlsx")
                    if excel_path_ld.exists():
                        with open(excel_path_ld, "rb") as f_ex:
                            st.download_button(
                                "📥 Tải File Excel Bảng Giá Ca Máy Lâm Đồng 2026 (.xlsx)",
                                data=f_ex.read(),
                                file_name="gia_ca_may_lam_dong_2026.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                with col_exp2:
                    if st.button("⚡ Tự Động Áp Giá Ca Máy Lâm Đồng Vào Dự Toán Hiện Tại", key="btn_apply_ld_prices"):
                        curr_prices = mgr.get_resource_prices()
                        updated_count = 0
                        for mac_row in curr_prices.get("machines", []):
                            code = mac_row.get("code")
                            lookup = mgr.lookup_machine_price(code, "lam_dong", current_region_key)
                            if lookup and lookup["price_current"] > 0:
                                mac_row["price_notice"] = lookup["price_current"]
                                mac_row["price_current"] = lookup["price_current"]
                                mac_row["legal_basis"] = f"QĐ 21/QĐ-SXD Lâm Đồng 2026 ({selected_region.split('(')[0].strip()})"
                                updated_count += 1
                        if updated_count > 0:
                            mgr.update_resource_prices(curr_prices)
                            st.success(f"Đã cập nhật giá cho {updated_count} máy thi công theo {selected_region}!")
                            st.rerun()
                        else:
                            st.info("Các mã máy trong dự toán hiện tại không trùng khớp với danh mục máy chuẩn xây dựng của QĐ 21.")

    with sub_t4:
        st.markdown("##### 4. QUẢN LÝ ĐƠN GIÁ THEO BẢN CHẤT KỸ THUẬT (CHUẨN HÓA ĐƠN GIÁ)")
        st.caption("Đảm bảo nguyên tắc thẩm định dự toán EVN: Các công việc có CÙNG BẢN CHẤT KỸ THUẬT thì phải áp dụng CÙNG 1 ĐƠN GIÁ, không có sự sai lệch giá giữa các vị trí/nhóm công tác.")

        tech_map_items = {}
        for it in st.session_state.work_items:
            tp = it.get("tech_profile") or mgr.get_default_tech_profile(it.get("norm_code", ""))
            if tp not in tech_map_items: tech_map_items[tp] = []
            tech_map_items[tp].append(it)

        tech_rows = []
        for tp, items in tech_map_items.items():
            first_it = items[0]
            rates = mgr.calculate_item_unit_rate(first_it, st.session_state.coefficients)
            tot_vol = sum(float(x.get("volume", 0.0)) for x in items)
            unique_codes = list(dict.fromkeys([str(x.get("norm_code", "")) for x in items]))
            tech_rows.append({
                "Bản chất kỹ thuật": tp,
                "Mã hiệu áp dụng": ", ".join(unique_codes),
                "ĐVT": first_it.get("unit", ""),
                "Số công tác áp dụng": len(items),
                "Tổng khối lượng toàn DA": tot_vol,
                "ĐG Vật liệu (đ)": rates["dg_vl"],
                "ĐG Nhân công (đ)": rates["dg_nc"],
                "ĐG Ca máy (đ)": rates["dg_m"],
                "ĐG Trực tiếp T (đ)": rates["dg_tt"],
                "ĐG Đầy đủ trước thuế (đ)": rates["dg_truoc_thue"],
                "Tổng tiền trước thuế (đ)": tot_vol * rates["dg_truoc_thue"],
                "Trạng thái": "🟢 100% Đồng nhất đơn giá"
            })

        if tech_rows:
            st.dataframe(
                pd.DataFrame(tech_rows).style.format({
                    "Tổng khối lượng toàn DA": "{:,.2f}",
                    "ĐG Vật liệu (đ)": "{:,.0f}",
                    "ĐG Nhân công (đ)": "{:,.0f}",
                    "ĐG Ca máy (đ)": "{:,.0f}",
                    "ĐG Trực tiếp T (đ)": "{:,.0f}",
                    "ĐG Đầy đủ trước thuế (đ)": "{:,.0f}",
                    "Tổng tiền trước thuế (đ)": "{:,.0f}"
                }),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("ℹ️ Hiện tại dự toán chưa có công tác nào để hiển thị bảng đơn giá theo bản chất kỹ thuật.")

        st.info("💡 **Hệ thống tự động thực thi:** Đơn giá trực tiếp và trước thuế được tính toán từ CSDL giá tài nguyên tập trung. Khi có nhiều công việc nằm ở các khu vực/hạng mục khác nhau nhưng cùng bản chất kỹ thuật, hệ thống luôn đảm bảo đơn giá hoàn toàn nhất quán.")

    col_b1, col_b2 = st.columns([3, 7])
    with col_b1:
        if st.button("💾 Lưu Cập Nhật Đơn Giá & Pháp Lý", type="primary"):
            updated = {
                "materials": edited_vl.to_dict("records"),
                "labor": edited_nc.to_dict("records"),
                "machines": edited_m.to_dict("records")
            }
            mgr.update_resource_prices(updated)
            st.success("Đã lưu bảng giá và căn cứ pháp lý thành công vào hệ thống!")
            st.rerun()

# -----------------------------------------------------------------------------
# TAB 3: BẢNG TỔNG HỢP KINH PHÍ & HỆ SỐ
# -----------------------------------------------------------------------------
with tab3:
    st.markdown("#### 📊 BẢNG TỔNG HỢP DỰ TOÁN CHI PHÍ XÂY DỰNG")
    st.caption("Bảng tổng hợp chi phí chuẩn mực theo mẫu hình chụp EVN. Cho phép nhập tay các hệ số và cơ sở hệ số.")

    # 1. Khung cấu hình các hệ số tỷ lệ
    with st.expander("⚙️ **CẤU HÌNH CÁC HỆ SỐ DỰ TOÁN & CƠ SỞ PHÁP LÝ**", expanded=False):
        c_cols = st.columns(5)
        with c_cols[0]:
            c_val = st.number_input("Hệ số C (%):", min_value=0.0, max_value=50.0, value=float(st.session_state.coefficients["c"]["rate"]), step=0.1)
            c_basis = st.text_area("Căn cứ hệ số C:", value=st.session_state.coefficients["c"]["basis"], height=68)
        with c_cols[1]:
            tt_val = st.number_input("Hệ số TT (%):", min_value=0.0, max_value=50.0, value=float(st.session_state.coefficients["tt"]["rate"]), step=0.1)
            tt_basis = st.text_area("Căn cứ hệ số TT:", value=st.session_state.coefficients["tt"]["basis"], height=68)
        with c_cols[2]:
            tl_val = st.number_input("Hệ số TL (%):", min_value=0.0, max_value=50.0, value=float(st.session_state.coefficients["tl"]["rate"]), step=0.1)
            tl_basis = st.text_area("Căn cứ hệ số TL:", value=st.session_state.coefficients["tl"]["basis"], height=68)
        with c_cols[3]:
            vat_val = st.number_input("Thuế GTGT (%):", min_value=0.0, max_value=20.0, value=float(st.session_state.coefficients["vat"]["rate"]), step=0.5)
            vat_basis = st.text_area("Căn cứ thuế GTGT:", value=st.session_state.coefficients["vat"]["basis"], height=68)
        with c_cols[4]:
            lt_val = st.number_input("Nhà tạm LT (%):", min_value=0.0, max_value=10.0, value=float(st.session_state.coefficients["lt"]["rate"]), step=0.1)
            lt_basis = st.text_area("Căn cứ nhà tạm LT:", value=st.session_state.coefficients["lt"]["basis"], height=68)

        if st.button("💾 Cập nhật Hệ số Dự toán"):
            st.session_state.coefficients = {
                "c": {"rate": c_val, "basis": c_basis},
                "tt": {"rate": tt_val, "basis": tt_basis},
                "tl": {"rate": tl_val, "basis": tl_basis},
                "vat": {"rate": vat_val, "basis": vat_basis},
                "lt": {"rate": lt_val, "basis": lt_basis}
            }
            st.success("Đã cập nhật các hệ số dự toán thành công!")

    # 2. Tính toán tổng hợp chi phí thời gian thực
    prices_data = mgr.get_resource_prices()
    vl_sum = sum(mat_qty.get(m["code"], 0.0) * (m.get("price_notice", 0) - m.get("price_base", 0)) for m in prices_data.get("materials", []))
    nc_sum = sum(lab_qty.get(l["code"], 0.0) * (l.get("price_current", 0) - l.get("price_base", 0)) for l in prices_data.get("labor", []))
    mac_sum = sum(mac_qty.get(mac["code"], 0.0) * (mac.get("price_current", mac.get("price_notice", 0) + mac.get("price_diff", 0)) - mac.get("price_base", 0)) for mac in prices_data.get("machines", []))

    # Calculate quote items sum
    quote_items = [it for it in st.session_state.work_items if it.get("is_quote")]
    quote_sum = sum(float(it.get("unit_price_quote", 0.0)) * float(it.get("volume", 0.0)) for it in quote_items)

    T_sum = vl_sum + nc_sum + mac_sum + quote_sum
    c_r = st.session_state.coefficients["c"]["rate"] / 100.0
    tt_r = st.session_state.coefficients["tt"]["rate"] / 100.0
    tl_r = st.session_state.coefficients["tl"]["rate"] / 100.0
    vat_r = st.session_state.coefficients["vat"]["rate"] / 100.0
    lt_r = st.session_state.coefficients["lt"]["rate"] / 100.0

    C_val = T_sum * c_r
    TT_val = T_sum * tt_r
    GT_val = C_val + TT_val
    TL_val = (T_sum + GT_val) * tl_r
    G_val = T_sum + GT_val + TL_val
    GTGT_val = G_val * vat_r
    Gxd_val = G_val + GTGT_val
    LT_val = G_val * lt_r * (1.0 + vat_r)
    TongCong = Gxd_val + LT_val
    LamTron = round(TongCong, -3)
    bang_chu_text = num2words_vi(LamTron)

    # Hiển thị bảng mô phỏng chuẩn xác
    summary_table_data = [
        {"STT": "I", "Nội dung chi phí": "Chi phí trực tiếp", "Cách tính": "=VL+NC+M+BG" if quote_sum > 0 else "=D7+D9+D11", "Giá trị (đồng)": T_sum, "Ký hiệu": "T", "Cơ sở hệ số / Ghi chú": "T = VL + NC + M" + (" + BG" if quote_sum > 0 else "")},
        {"STT": "1", "Nội dung chi phí": "Chi phí vật liệu", "Cách tính": "=D8", "Giá trị (đồng)": vl_sum, "Ký hiệu": "VL", "Cơ sở hệ số / Ghi chú": "VL = VLHT"},
        {"STT": "", "Nội dung chi phí": " - Đơn giá vật liệu", "Cách tính": "Theo bảng tổng hợp vật liệu", "Giá trị (đồng)": vl_sum, "Ký hiệu": "VLHT", "Cơ sở hệ số / Ghi chú": "Theo bảng TH_VatLieu"},
        {"STT": "2", "Nội dung chi phí": "Chi phí nhân công", "Cách tính": "=D10", "Giá trị (đồng)": nc_sum, "Ký hiệu": "NC", "Cơ sở hệ số / Ghi chú": "NC = NCHT"},
        {"STT": "", "Nội dung chi phí": " - Đơn giá nhân công", "Cách tính": "Theo bảng tổng hợp nhân công", "Giá trị (đồng)": nc_sum, "Ký hiệu": "NCHT", "Cơ sở hệ số / Ghi chú": "Theo bảng TH_NhanCong"},
        {"STT": "3", "Nội dung chi phí": "Chi phí máy và thiết bị thi công", "Cách tính": "=D12", "Giá trị (đồng)": mac_sum, "Ký hiệu": "M", "Cơ sở hệ số / Ghi chú": "M = MHT"},
        {"STT": "", "Nội dung chi phí": " - Đơn giá máy thi công", "Cách tính": "Theo bảng tổng hợp máy thi công", "Giá trị (đồng)": mac_sum, "Ký hiệu": "MHT", "Cơ sở hệ số / Ghi chú": "Theo bảng TH_MayThiCong"}
    ]

    if quote_sum > 0:
        summary_table_data.append({
            "STT": "4", "Nội dung chi phí": "Chi phí công tác báo giá / tạm tính", "Cách tính": "Theo báo giá trọn gói", "Giá trị (đồng)": quote_sum, "Ký hiệu": "BG", "Cơ sở hệ số / Ghi chú": f"{len(quote_items)} công tác không có định mức"
        })

    summary_table_data.extend([
        {"STT": "II", "Nội dung chi phí": "Chi phí gián tiếp", "Cách tính": "=C+TT", "Giá trị (đồng)": GT_val, "Ký hiệu": "GT", "Cơ sở hệ số / Ghi chú": "GT = C + TT"},
        {"STT": "1", "Nội dung chi phí": "Chi phí chung", "Cách tính": f"=T * {st.session_state.coefficients['c']['rate']}%", "Giá trị (đồng)": C_val, "Ký hiệu": "C", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['c']['basis']},
        {"STT": "2", "Nội dung chi phí": "Chi phí công việc không XĐ từ thiết kế", "Cách tính": f"=T * {st.session_state.coefficients['tt']['rate']}%", "Giá trị (đồng)": TT_val, "Ký hiệu": "TT", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['tt']['basis']},
        {"STT": "III", "Nội dung chi phí": "Thu nhập chịu thuế tính trước", "Cách tính": f"=(T+GT) * {st.session_state.coefficients['tl']['rate']}%", "Giá trị (đồng)": TL_val, "Ký hiệu": "TL", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['tl']['basis']},
        {"STT": "", "Nội dung chi phí": "Chi phí xây dựng trước thuế", "Cách tính": "=T+GT+TL", "Giá trị (đồng)": G_val, "Ký hiệu": "G", "Cơ sở hệ số / Ghi chú": "G = T + GT + TL"},
        {"STT": "IV", "Nội dung chi phí": "Thuế giá trị gia tăng", "Cách tính": f"=G * {st.session_state.coefficients['vat']['rate']}%", "Giá trị (đồng)": GTGT_val, "Ký hiệu": "GTGT", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['vat']['basis']},
        {"STT": "", "Nội dung chi phí": "CHI PHÍ XÂY DỰNG SAU THUẾ", "Cách tính": "=G+GTGT", "Giá trị (đồng)": Gxd_val, "Ký hiệu": "Gxd", "Cơ sở hệ số / Ghi chú": "Gxd = G + GTGT"},
        {"STT": "V", "Nội dung chi phí": "Chi phí nhà tạm để ở và điều hành thi công", "Cách tính": f"=G * {st.session_state.coefficients['lt']['rate']}% * (1+{st.session_state.coefficients['vat']['rate']}%)", "Giá trị (đồng)": LT_val, "Ký hiệu": "LT", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['lt']['basis']},
        {"STT": "", "Nội dung chi phí": "Tổng cộng", "Cách tính": "=Gxd+LT", "Giá trị (đồng)": TongCong, "Ký hiệu": "", "Cơ sở hệ số / Ghi chú": ""},
        {"STT": "", "Nội dung chi phí": "Làm tròn (đến hàng nghìn)", "Cách tính": "=ROUND(Tổng cộng, -3)", "Giá trị (đồng)": LamTron, "Ký hiệu": "", "Cơ sở hệ số / Ghi chú": ""}
    ])

    df_summary = pd.DataFrame(summary_table_data)
    st.dataframe(
        df_summary.style.format({
            "Giá trị (đồng)": "{:,.0f}"
        }),
        use_container_width=True,
        hide_index=True
    )

    st.markdown(f"""
    <div class="metric-card">
        <span style="font-size:16px; color:#002060;"><b>📝 BẰNG CHỮ:</b> <i>{bang_chu_text}</i></span>
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# TAB 4: ĐƠN GIÁ CHI TIẾT TRỰC QUAN & MINH CHỨNG SCAN
# -----------------------------------------------------------------------------
with tab4:
    st.markdown("#### 🔍 CHI TIẾT ĐƠN GIÁ CÔNG TÁC & MINH CHỨNG SCAN GỐC")
    st.caption("Xem chiết tính hao phí định mức và hồ sơ scan gốc PDF phục vụ công tác thanh kiểm tra, thẩm định.")

    sub_e1, sub_e2 = st.tabs(["📊 Chiết Tính Đơn Giá Chi Tiết", "🖼️ Hồ Sơ Minh Chứng Scan Gốc"])

    with sub_e1:
        vl_p_map = {m["code"]: m.get("price_notice", 0) for m in mgr.get_resource_prices().get("materials", [])}
        nc_p_map = {l["code"]: l.get("price_current", 0) for l in mgr.get_resource_prices().get("labor", [])}
        mac_p_map = {mac["code"]: mac.get("price_current", 0) for mac in mgr.get_resource_prices().get("machines", [])}

        if not st.session_state.work_items:
            st.info("ℹ️ Hiện tại dự toán chưa có công tác nào. Vui lòng thêm công tác mới tại Tab 1 để kiểm tra chiết tính đơn giá.")
        else:
            sel_stt = st.selectbox(
                "Chọn công việc cần kiểm tra chiết tính:",
                options=[item["stt"] for item in st.session_state.work_items],
                format_func=lambda x: f"Công tác {x}: {[it['job_name'] for it in st.session_state.work_items if it['stt'] == x][0]} ({[it['norm_code'] for it in st.session_state.work_items if it['stt'] == x][0]})"
            )
            matching_items = [it for it in st.session_state.work_items if it["stt"] == sel_stt]
            if matching_items:
                it_sel = matching_items[0]
                n_info = mgr.get_norm(it_sel.get("norm_code", ""))

                if n_info:
                    c1, c2, c3 = st.columns([5, 2, 2])
                    with c1:
                        st.markdown(f"**Tên công việc:** {it_sel['job_name']}")
                        st.caption(f"Căn cứ: {n_info.get('legal_basis', '')}")
                    with c2:
                        st.metric("Khối lượng dự án", f"{it_sel['volume']:,.2f} {it_sel['unit']}")
                    with c3:
                        st.metric("Định mức gốc tính cho", n_info.get("unit", ""))

                    rows = []
                    t_vl, t_nc, t_m = 0.0, 0.0, 0.0
                    res = n_info.get("resources", {})
                    for m in res.get("materials", []):
                        dg = vl_p_map.get(m["code"], 0)
                        tt = m["rate"] * dg if m["code"] != "V00750" else 0
                        rows.append({"Phân loại": "Vật liệu", "Mã hiệu": m["code"], "Tên tài nguyên": m["name"], "ĐVT": m["unit"], "Định mức": m["rate"], "Đơn giá (đ)": dg, "Thành tiền (đ)": tt})
                        t_vl += tt
                    for l in res.get("labor", []):
                        dg = nc_p_map.get(l["code"], 0)
                        tt = l["rate"] * dg
                        rows.append({"Phân loại": "Nhân công", "Mã hiệu": l["code"], "Tên tài nguyên": l["name"], "ĐVT": l["unit"], "Định mức": l["rate"], "Đơn giá (đ)": dg, "Thành tiền (đ)": tt})
                        t_nc += tt
                    for mac in res.get("machines", []):
                        dg = mac_p_map.get(mac["code"], 0)
                        tt = mac["rate"] * dg
                        rows.append({"Phân loại": "Máy thi công", "Mã hiệu": mac["code"], "Tên tài nguyên": mac["name"], "ĐVT": mac["unit"], "Định mức": mac["rate"], "Đơn giá (đ)": dg, "Thành tiền (đ)": tt})
                        t_m += tt

                    st.dataframe(
                        pd.DataFrame(rows).style.format({
                            "Định mức": "{:,.2f}", "Đơn giá (đ)": "{:,.0f}", "Thành tiền (đ)": "{:,.0f}"
                        }),
                        use_container_width=True, hide_index=True
                    )
                    st.info(f"Tổng đơn giá cho 1 đơn vị định mức: **{(t_vl + t_nc + t_m):,.0f} đ** (Vật liệu: {t_vl:,.0f} đ | Nhân công: {t_nc:,.0f} đ | Máy: {t_m:,.0f} đ)")
                else:
                    st.info(f"💡 Công tác '{it_sel.get('job_name', '')}' ({it_sel.get('norm_code', '')}) là công tác áp dụng báo giá trọn gói / tạm tính, không phân tích chiết tính định mức tài nguyên.")

    with sub_e2:
        code_view = st.selectbox("Chọn mã định mức xem ảnh minh chứng scan gốc:", options=mgr.get_all_norm_codes(), key="sel_scan_code")
        n_detail = mgr.get_norm(code_view)
        if n_detail:
            c_scan, c_steps = st.columns([5, 5])
            with c_scan:
                img_p = mgr.get_evidence_image_path(code_view)
                if img_p and img_p.exists():
                    st.image(str(img_p), caption=f"Minh chứng pháp lý: {n_detail.get('legal_basis', '')}", use_container_width=True)
                else:
                    st.warning("Chưa có ảnh scan minh chứng cho mã này.")
            with c_steps:
                st.markdown("**Bảng phân tích tỷ lệ các bước thực hiện:**")
                st.caption(f"Đơn vị tính gốc: {n_detail.get('unit', '')} | Căn cứ: {n_detail.get('legal_basis', '')}")
                steps = n_detail.get("work_steps", [])
                labor_split = n_detail.get("labor_split")
                if steps:
                    df_steps = pd.DataFrame(steps)
                    if labor_split:
                        # Display split labor columns for BD norms
                        disp_cols = ["step", "name", "dvt_step", "rate_nc", "rate_ks", "rate_base", "apply_flag", "rate_calculated"]
                        disp_df = df_steps[[c for c in disp_cols if c in df_steps.columns]].rename(columns={
                            "step": "Bước", "name": "Nội dung công việc", "dvt_step": "ĐVT",
                            "rate_nc": "NC nhóm 3 (công)", "rate_ks": "Kỹ sư (công)",
                            "rate_base": "Tổng ĐM gốc", "apply_flag": "Áp dụng (1/0)", "rate_calculated": "Tính toán"
                        })
                        st.dataframe(disp_df, use_container_width=True, hide_index=True)
                        lbl_nc = labor_split.get("nc_total_per_unit", 0.087)
                        lbl_ks = labor_split.get("ks_total_per_unit", 0.087)
                        lbl_tot = labor_split.get("total_per_unit", 0.174)
                        st.info(f"💡 **Tách rõ thành phần:** Tổng cộng **{lbl_tot} công** = **{lbl_nc} công NC nhóm 3** (Bậc 3,5/7) + **{lbl_ks} công Kỹ sư** (Bậc 5,0/8) theo TT 38/2026/TT-BXD.")
                    else:
                        st.dataframe(
                            df_steps[["step", "name", "dvt_step", "rate_base", "apply_flag", "rate_calculated"]].rename(columns={
                                "step": "Bước", "name": "Nội dung công việc", "dvt_step": "ĐVT",
                                "rate_base": "Hao phí gốc", "apply_flag": "Áp dụng (1/0)", "rate_calculated": "Tính toán"
                            }),
                            use_container_width=True,
                            hide_index=True
                        )
                        tot_s = sum(s.get("rate_calculated", 0) for s in steps)
                        st.info(f"Tổng hao phí nhân công các bước: **{tot_s:.4f} công**")

# -----------------------------------------------------------------------------
# TAB 5: XUẤT BẢN FILE EXCEL DỰ TOÁN
# -----------------------------------------------------------------------------
with tab5:
    st.markdown("#### 🚀 XUẤT BẢN FILE EXCEL DỰ TOÁN THẨM ĐỊNH")
    st.caption("Tạo file Excel đầy đủ 20 sheet với 100% công thức liên kết động (Dynamic Links), tự động thêm sheet minh chứng cho các định mức được gọi.")

    st.write(f"📁 **Vị trí lưu trữ mặc định:** `{OUTPUT_DIR}`")
    export_filename = st.text_input("Tên tệp Excel xuất ra:", value="DU_TOAN_PCCC_2026_THAM_DINH.xlsx")

    if st.button("🚀 BẮT ĐẦU TẠO FILE EXCEL 20 SHEET", type="primary", use_container_width=True):
        if not st.session_state.work_items:
            st.warning("⚠️ Dự toán chưa có công tác nào. Vui lòng thêm ít nhất 1 công tác tại Tab 1 trước khi xuất file Excel thẩm định.")
        else:
            out_target = OUTPUT_DIR / export_filename
            with st.spinner("Đang tổng hợp dữ liệu, tạo bảng TH_KinhPhi chuẩn mực, chèn ảnh scan và thiết lập 100% công thức liên kết động..."):
                try:
                    generator.generate(
                        project_name=project_name,
                        category_name=category_name,
                        work_items=st.session_state.work_items,
                        output_path=out_target,
                        coefficients=st.session_state.coefficients
                    )
                    st.success(f"✅ ĐÃ TẠO FILE EXCEL THÀNH CÔNG: {out_target.name}")
                    
                    with open(out_target, "rb") as f:
                        excel_bytes = f.read()
                    
                    st.download_button(
                        label="📥 TẢI FILE EXCEL VỀ MÁY TÍNH",
                        data=excel_bytes,
                        file_name=export_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Lỗi trong quá trình tạo file: {e}")
