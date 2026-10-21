"""
Streamlit Web App: Hệ Thống Quản Lý Định Mức & Lập Dự Toán PCCC - SCL EVN
Author: Tổ Thẩm Định EVN
"""

import streamlit as st
import pandas as pd
import json
from pathlib import Path
import sys

# Ensure project root is in sys.path
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import DATA_DIR, OUTPUT_DIR, IMAGES_DIR, check_environment
from core.norm_manager import NormManager
from core.excel_generator import ExcelGenerator, DEFAULT_COEFFICIENTS
from utils.num2words_vi import num2words_vi

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
if "norm_manager" not in st.session_state:
    st.session_state.norm_manager = NormManager()

if "coefficients" not in st.session_state:
    st.session_state.coefficients = {
        "c": {"rate": 7.3, "basis": "Thông tư số 11/2021/TT-BXD (Công trình công nghiệp)"},
        "tt": {"rate": 2.5, "basis": "Thông tư số 11/2021/TT-BXD (Không xác định KL từ thiết kế)"},
        "tl": {"rate": 5.5, "basis": "Thông tư số 11/2021/TT-BXD (Thu nhập chịu thuế tính trước)"},
        "vat": {"rate": 8.0, "basis": "Nghị định 72/2024/NĐ-CP của Chính phủ"},
        "lt": {"rate": 1.1, "basis": "Thông tư số 11/2021/TT-BXD (Nhà tạm ngoài đô thị)"}
    }

if "work_items" not in st.session_state:
    # Default 13 work items with groups
    st.session_state.work_items = [
        {"stt": 1, "group": "Hệ thống báo cháy tự động và thiết bị ngoại vi", "norm_code": "BD.41111", "job_name": "Thay mới lắp đặt thiết bị đầu báo khói địa chỉ", "description": "Dịch vụ Thay thiết bị đầu báo", "volume": 10.0, "unit": "Bộ"},
        {"stt": 2, "group": "Hệ thống báo cháy tự động và thiết bị ngoại vi", "norm_code": "BD.41121", "job_name": "Lắp đặt đèn báo cháy", "description": "Lắp mới / thay thế đèn báo cháy vị trí khu vực phụ trợ", "volume": 5.0, "unit": "Bộ"},
        {"stt": 3, "group": "Hệ thống báo cháy tự động và thiết bị ngoại vi", "norm_code": "BD.41131", "job_name": "Lắp đặt lại nút nhấn khẩn báo cháy", "description": "Lắp đặt lại nút nhấn sau khi bảo trì", "volume": 38.0, "unit": "Cái"},
        {"stt": 4, "group": "Hệ thống báo cháy tự động và thiết bị ngoại vi", "norm_code": "BD.41141", "job_name": "Lắp đặt chuông báo cháy", "description": "Thay thế, lắp đặt chuông báo cháy tự động", "volume": 5.0, "unit": "Bộ"},
        {"stt": 5, "group": "Hệ thống báo cháy tự động và thiết bị ngoại vi", "norm_code": "BD.42201", "job_name": "Lắp đặt bộ điều khiển, bộ chuyển mạch PCCC", "description": "Lắp đặt bộ chuyển mạch tín hiệu tủ trung gian", "volume": 1.0, "unit": "Bộ"},
        {"stt": 6, "group": "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy", "norm_code": "E6C.3.01", "job_name": "Bảo trì, sửa chữa, thử nghiệm giả lập tác động nút nhấn khẩn", "description": "Kiểm tra, bảo dưỡng nút nhấn báo cháy chuyên dụng", "volume": 38.0, "unit": "Cái"},
        {"stt": 7, "group": "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy", "norm_code": "E6C.3.02", "job_name": "Bảo trì, sửa chữa, hiệu chỉnh, thử nghiệm độ nhạy đầu báo khói/nhiệt", "description": "Dịch vụ Bảo trì đầu báo chuyên dụng", "volume": 890.0, "unit": "Bộ"},
        {"stt": 8, "group": "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy", "norm_code": "E6C.3.03", "job_name": "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng dây báo cháy", "description": "Kiểm tra module dây giám sát nhiệt độ cao", "volume": 20.0, "unit": "Bộ"},
        {"stt": 9, "group": "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy", "norm_code": "E6C.3.04", "job_name": "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng beam quang", "description": "Bảo dưỡng beam khói khoảng cách xa", "volume": 15.0, "unit": "Bộ"},
        {"stt": 10, "group": "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy", "norm_code": "E6C.3.05", "job_name": "Sửa chữa, bảo dưỡng trung tâm báo cháy FMZ 5000", "description": "Bảo dưỡng tủ điều khiển chữa cháy tại chỗ và tủ trung gian", "volume": 3.0, "unit": "trung tâm"},
        {"stt": 11, "group": "Bảo dưỡng van xả tràn Deluge, ắc quy và đo khí", "norm_code": "E6C.3.06", "job_name": "Kiểm tra và bảo dưỡng phần điện van Deluge, Wet hệ PCCC", "description": "Bảo dưỡng cuộn hút, van xả tràn Deluge và đồng hồ áp", "volume": 12.0, "unit": "Cái"},
        {"stt": 12, "group": "Bảo dưỡng van xả tràn Deluge, ắc quy và đo khí", "norm_code": "E6B.1.13.01", "job_name": "Kiểm tra bình accu chuyên dụng phục vụ hệ thống PCCC", "description": "Đo nội trở, dung lượng Ah, vô mỡ cọc bình và châm nước cất", "volume": 50.0, "unit": "Bình"},
        {"stt": 13, "group": "Bảo dưỡng van xả tràn Deluge, ắc quy và đo khí", "norm_code": "E6C.1.31", "job_name": "Kiểm tra và bảo dưỡng thiết bị đo khí NH3, Clo, H2", "description": "Bơm khí chuẩn hiệu chuẩn điểm Zero & Span cho đầu đo khí", "volume": 4.0, "unit": "Bộ"}
    ]

mgr = st.session_state.norm_manager
generator = ExcelGenerator(mgr)

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
# TAB 1: KHỐI LƯỢNG TIÊN LƯỢNG (TRA CỨU, PHÂN NHÓM & CRUD)
# -----------------------------------------------------------------------------
with tab1:
    st.markdown("#### 📋 BẢNG TIÊN LƯỢNG KHỐI LƯỢNG & QUẢN LÝ THEO NHÓM")
    st.caption("Tra cứu định mức, gợi ý tên công tác theo PAKT, quản lý chỉnh sửa và đổi tên nhóm trực tiếp trong từng nhóm.")

    # 1. Tra cứu & Thêm công tác mới
    with st.expander("➕ **TRA CỨU ĐỊNH MỨC & THÊM CÔNG TÁC MỚI**", expanded=False):
        all_codes = mgr.get_all_norm_codes()
        norm_options = []
        for c in all_codes:
            nm = mgr.get_norm(c)
            norm_options.append({
                "code": c,
                "label": f"{c} - {nm.get('name', '')} (ĐVT: {nm.get('unit', '')})",
                "name": nm.get("name", ""),
                "unit": nm.get("unit", ""),
                "basis": nm.get("legal_basis", "")
            })

        col_search, col_grp = st.columns([6, 4])
        with col_search:
            selected_option = st.selectbox(
                "Gõ từ khóa để tra cứu định mức (Mã hoặc Tên):",
                options=norm_options,
                format_func=lambda x: x["label"],
                key="add_norm_search"
            )
        with col_grp:
            # Existing groups
            existing_groups = list(dict.fromkeys([it.get("group", "Nhóm công việc chung") for it in st.session_state.work_items if it.get("group")]))
            if not existing_groups: existing_groups = ["Hệ thống báo cháy tự động và thiết bị ngoại vi (Lắp đặt - Thay mới)", "Bảo dưỡng, sửa chữa và hiệu chuẩn thiết bị báo cháy", "Bảo dưỡng van tràn Deluge, hệ thống ắc quy và thiết bị đo khí"]
            selected_group_choice = st.selectbox("Chọn nhóm công tác:", options=existing_groups + ["➕ [Tạo nhóm mới...]"], key="add_norm_group_sel")
            if selected_group_choice == "➕ [Tạo nhóm mới...]":
                new_group_name = st.text_input("Nhập tên nhóm mới:", value="Nhóm công việc mới", key="add_norm_group_new")
                final_group = new_group_name.strip()
            else:
                final_group = selected_group_choice

        # PAKT Suggestions based on selected norm
        sel_code = selected_option["code"]
        raw_suggestions = PAKT_JOB_SUGGESTIONS.get(sel_code, [])
        if selected_option["name"] not in raw_suggestions:
            raw_suggestions = raw_suggestions + [selected_option["name"]]
        
        st.markdown("**💡 Gợi ý tên danh mục công việc theo PAKT (chọn để tự động điền):**")
        selected_sugg = st.selectbox(
            "Danh sách gợi ý phổ biến cho mã này:",
            options=raw_suggestions,
            index=0,
            key=f"pakt_sugg_{sel_code}"
        )

        col_f1, col_f2, col_f3, col_f4 = st.columns([5, 4, 2, 2])
        with col_f1:
            inp_job_name = st.text_input("Tên danh mục công việc (Theo PAKT):", value=selected_sugg, key=f"inp_job_name_{sel_code}")
        with col_f2:
            inp_desc = st.text_input("Mô tả công việc thực tế:", value=selected_sugg, key=f"inp_desc_{sel_code}")
        with col_f3:
            inp_vol = st.number_input("Khối lượng:", min_value=0.01, value=1.0, step=1.0, key=f"inp_vol_{sel_code}")
        with col_f4:
            inp_unit = st.text_input("ĐVT:", value=selected_option["unit"], key=f"inp_unit_{sel_code}")

        st.caption(f"📌 Căn cứ pháp lý gốc: *{selected_option['basis']}*")

        if st.button("➕ Thêm công tác này vào nhóm", type="primary"):
            new_stt = len(st.session_state.work_items) + 1
            st.session_state.work_items.append({
                "stt": new_stt,
                "group": final_group,
                "norm_code": selected_option["code"],
                "job_name": inp_job_name,
                "description": inp_desc,
                "volume": float(inp_vol),
                "unit": inp_unit
            })
            st.success(f"Đã thêm công tác '{inp_job_name}' vào nhóm '{final_group}'!")
            st.rerun()

    # 2. Quản lý theo từng Nhóm (Xem, Chỉnh sửa, Đổi tên & Xóa trực tiếp trong nhóm)
    st.markdown("##### 🗂️ QUẢN LÝ CHI TIẾT THEO TỪNG NHÓM CÔNG VIỆC")
    st.caption("Xem, sửa đơn giá khối lượng, đổi tên nhóm hoặc xóa công việc trực tiếp trong từng nhóm bên dưới.")

    # Group items by group preserving order
    groups_dict = {}
    for it in st.session_state.work_items:
        g = (it.get("group") or "Công việc chung").strip() or "Công việc chung"
        if g not in groups_dict: groups_dict[g] = []
        groups_dict[g].append(it)

    norm_codes_list = mgr.get_all_norm_codes()

    for g_idx, (g_name, items_in_g) in enumerate(groups_dict.items(), start=1):
        with st.expander(f"📁 **NHÓM {g_idx}: {g_name.upper()}** ({len(items_in_g)} công tác)", expanded=True):
            # Toolbar: Rename group and Delete group
            col_ren_title, col_ren_btn, col_del_grp = st.columns([6, 3, 3])
            with col_ren_title:
                new_group_title = st.text_input(f"Tên nhóm {g_idx}:", value=g_name, key=f"edit_group_name_{g_idx}")
            with col_ren_btn:
                st.write("") # vertical spacing
                if st.button("✏️ Đổi tên nhóm", key=f"btn_apply_rename_{g_idx}", use_container_width=True):
                    if new_group_title.strip() and new_group_title.strip() != g_name:
                        for it in st.session_state.work_items:
                            if it.get("group", "").strip() == g_name:
                                it["group"] = new_group_title.strip()
                        st.success(f"Đã đổi tên nhóm thành: {new_group_title.strip()}")
                        st.rerun()
            with col_del_grp:
                st.write("")
                if st.button("🗑️ Xóa cả nhóm này", key=f"btn_del_entire_group_{g_idx}", use_container_width=True):
                    st.session_state.work_items = [it for it in st.session_state.work_items if it.get("group", "").strip() != g_name]
                    # Re-index STT
                    for i, it in enumerate(st.session_state.work_items, start=1):
                        it["stt"] = i
                    st.success(f"Đã xóa toàn bộ công tác thuộc nhóm '{g_name}'!")
                    st.rerun()

            # In-group Data Editor
            df_g = pd.DataFrame(items_in_g)
            disp_cols = ["stt", "norm_code", "job_name", "description", "volume", "unit"]
            edited_group_df = st.data_editor(
                df_g[disp_cols],
                column_config={
                    "stt": st.column_config.NumberColumn("STT", width="small", disabled=True),
                    "norm_code": st.column_config.SelectboxColumn("Mã Định Mức", options=norm_codes_list, width="medium", required=True),
                    "job_name": st.column_config.TextColumn("Danh mục công việc (Theo PAKT)", width="large", required=True),
                    "description": st.column_config.TextColumn("Mô tả công việc thực tế", width="large"),
                    "volume": st.column_config.NumberColumn("Khối lượng", format="%.2f", min_value=0.0, step=1.0, required=True),
                    "unit": st.column_config.TextColumn("Đơn vị tính", width="small", required=True),
                },
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                key=f"group_editor_{g_idx}"
            )

            if st.button(f"💾 Lưu cập nhật Nhóm {g_idx}", key=f"save_group_{g_idx}", type="primary"):
                updated_g_records = edited_group_df.to_dict("records")
                for r in updated_g_records:
                    r["group"] = g_name
                
                # Replace items of this group in work_items
                other_items = [it for it in st.session_state.work_items if it.get("group", "").strip() != g_name]
                new_all_items = []
                # Reconstruct full list preserving group position
                for check_g in groups_dict.keys():
                    if check_g == g_name:
                        new_all_items.extend(updated_g_records)
                    else:
                        new_all_items.extend([it for it in other_items if it.get("group", "").strip() == check_g])
                
                # Re-index STT
                for i, it in enumerate(new_all_items, start=1):
                    it["stt"] = i
                st.session_state.work_items = new_all_items
                st.success(f"Đã lưu các thay đổi cho Nhóm {g_idx} thành công!")
                st.rerun()

# -----------------------------------------------------------------------------
# TAB 2: QUẢN LÝ ĐƠN GIÁ (VẬT LIỆU, NHÂN CÔNG, MÁY THI CÔNG & PHÁP LÝ)
# -----------------------------------------------------------------------------
with tab2:
    st.markdown("#### 💰 QUẢN LÝ ĐƠN GIÁ TÀI NGUYÊN & CƠ SỞ PHÁP LÝ")
    st.caption("Cột Khối lượng được tính toán tự động từ bảng Tiên lượng. Người dùng tùy chỉnh Đơn giá và Căn cứ pháp lý.")

    # Calculate real-time resource quantities from work_items
    mat_qty = {}
    lab_qty = {}
    mac_qty = {}
    for item in st.session_state.work_items:
        c = item["norm_code"]
        v = float(item.get("volume", 0.0))
        norm = mgr.get_norm(c)
        if not norm: continue
        u_str = norm.get("unit", "")
        div = 10.0 if "10" in u_str else (5.0 if "5" in u_str else 1.0)
        res = norm.get("resources", {})
        for m in res.get("materials", []):
            if m.get("code"): mat_qty[m["code"]] = mat_qty.get(m["code"], 0.0) + (v * float(m.get("rate", 0.0)) / div)
        for l in res.get("labor", []):
            if l.get("code"): lab_qty[l["code"]] = lab_qty.get(l["code"], 0.0) + (v * float(l.get("rate", 0.0)) / div)
        for mac in res.get("machines", []):
            if mac.get("code"): mac_qty[mac["code"]] = mac_qty.get(mac["code"], 0.0) + (v * float(mac.get("rate", 0.0)) / div)

    sub_t1, sub_t2, sub_t3 = st.tabs(["🧱 1. Vật Liệu", "👷 2. Nhân Công", "🚜 3. Máy Thi Công"])
    prices = mgr.get_resource_prices()

    with sub_t1:
        st.markdown("##### 1. BẢNG TỔNG HỢP VẬT LIỆU")
        df_vl = pd.DataFrame(prices.get("materials", []))
        if not df_vl.empty:
            df_vl["qty_calc"] = df_vl["code"].apply(lambda x: round(mat_qty.get(x, 0.0), 3))
            if "legal_basis" not in df_vl.columns: df_vl["legal_basis"] = df_vl.get("notes", "Công bố giá LS / Báo giá NSX")
            edited_vl = st.data_editor(
                df_vl[["code", "name", "unit", "qty_calc", "price_base", "price_notice", "legal_basis"]],
                column_config={
                    "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                    "name": st.column_config.TextColumn("Tên vật tư", disabled=True),
                    "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                    "qty_calc": st.column_config.NumberColumn("Khối lượng (Tiên lượng)", format="%.3f", disabled=True),
                    "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                    "price_notice": st.column_config.NumberColumn("Giá thông báo (đ)", format="%d"),
                    "legal_basis": st.column_config.TextColumn("Cơ sở đơn giá / Căn cứ pháp lý", width="large")
                },
                use_container_width=True,
                hide_index=True,
                key="editor_vl"
            )

    with sub_t2:
        st.markdown("##### 2. BẢNG TỔNG HỢP NHÂN CÔNG")
        df_nc = pd.DataFrame(prices.get("labor", []))
        if not df_nc.empty:
            df_nc["qty_calc"] = df_nc["code"].apply(lambda x: round(lab_qty.get(x, 0.0), 3))
            if "legal_basis" not in df_nc.columns: df_nc["legal_basis"] = "5898/SXD-QLXD ngày 02/7/2026 của SXD tỉnh Lâm Đồng"
            edited_nc = st.data_editor(
                df_nc[["code", "name", "unit", "qty_calc", "price_base", "price_current", "legal_basis"]],
                column_config={
                    "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                    "name": st.column_config.TextColumn("Tên nhân công", disabled=True),
                    "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                    "qty_calc": st.column_config.NumberColumn("Khối lượng (Tiên lượng)", format="%.3f", disabled=True),
                    "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                    "price_current": st.column_config.NumberColumn("Giá hiện tại (đ)", format="%d"),
                    "legal_basis": st.column_config.TextColumn("Cơ sở đơn giá / Căn cứ pháp lý", width="large")
                },
                use_container_width=True,
                hide_index=True,
                key="editor_nc"
            )

    with sub_t3:
        st.markdown("##### 3. BẢNG TỔNG HỢP MÁY THI CÔNG")
        df_m = pd.DataFrame(prices.get("machines", []))
        if not df_m.empty:
            df_m["qty_calc"] = df_m["code"].apply(lambda x: round(mac_qty.get(x, 0.0), 3))
            if "legal_basis" not in df_m.columns: df_m["legal_basis"] = "Thông tư 37/2026/TT-BXD của Bộ Xây dựng"
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
                    "legal_basis": st.column_config.TextColumn("Cơ sở đơn giá / Căn cứ pháp lý", width="large")
                },
                use_container_width=True,
                hide_index=True,
                key="editor_m"
            )

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

    T_sum = vl_sum + nc_sum + mac_sum
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
        {"STT": "I", "Nội dung chi phí": "Chi phí trực tiếp", "Cách tính": "=D7+D9+D11", "Giá trị (đồng)": T_sum, "Ký hiệu": "T", "Cơ sở hệ số / Ghi chú": "T = VL + NC + M"},
        {"STT": "1", "Nội dung chi phí": "Chi phí vật liệu", "Cách tính": "=D8", "Giá trị (đồng)": vl_sum, "Ký hiệu": "VL", "Cơ sở hệ số / Ghi chú": "VL = VLHT"},
        {"STT": "", "Nội dung chi phí": " - Đơn giá vật liệu", "Cách tính": "Theo bảng tổng hợp vật liệu", "Giá trị (đồng)": vl_sum, "Ký hiệu": "VLHT", "Cơ sở hệ số / Ghi chú": "Theo bảng TH_VatLieu"},
        {"STT": "2", "Nội dung chi phí": "Chi phí nhân công", "Cách tính": "=D10", "Giá trị (đồng)": nc_sum, "Ký hiệu": "NC", "Cơ sở hệ số / Ghi chú": "NC = NCHT"},
        {"STT": "", "Nội dung chi phí": " - Đơn giá nhân công", "Cách tính": "Theo bảng tổng hợp nhân công", "Giá trị (đồng)": nc_sum, "Ký hiệu": "NCHT", "Cơ sở hệ số / Ghi chú": "Theo bảng TH_NhanCong"},
        {"STT": "3", "Nội dung chi phí": "Chi phí máy và thiết bị thi công", "Cách tính": "=D12", "Giá trị (đồng)": mac_sum, "Ký hiệu": "M", "Cơ sở hệ số / Ghi chú": "M = MHT"},
        {"STT": "", "Nội dung chi phí": " - Đơn giá máy thi công", "Cách tính": "Theo bảng tổng hợp máy thi công", "Giá trị (đồng)": mac_sum, "Ký hiệu": "MHT", "Cơ sở hệ số / Ghi chú": "Theo bảng TH_MayThiCong"},
        {"STT": "II", "Nội dung chi phí": "Chi phí gián tiếp", "Cách tính": "=D14+D15", "Giá trị (đồng)": GT_val, "Ký hiệu": "GT", "Cơ sở hệ số / Ghi chú": "GT = C + TT"},
        {"STT": "1", "Nội dung chi phí": "Chi phí chung", "Cách tính": f"=T * {st.session_state.coefficients['c']['rate']}%", "Giá trị (đồng)": C_val, "Ký hiệu": "C", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['c']['basis']},
        {"STT": "2", "Nội dung chi phí": "Chi phí công việc không XĐ từ thiết kế", "Cách tính": f"=T * {st.session_state.coefficients['tt']['rate']}%", "Giá trị (đồng)": TT_val, "Ký hiệu": "TT", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['tt']['basis']},
        {"STT": "III", "Nội dung chi phí": "Thu nhập chịu thuế tính trước", "Cách tính": f"=(T+GT) * {st.session_state.coefficients['tl']['rate']}%", "Giá trị (đồng)": TL_val, "Ký hiệu": "TL", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['tl']['basis']},
        {"STT": "", "Nội dung chi phí": "Chi phí xây dựng trước thuế", "Cách tính": "=T+GT+TL", "Giá trị (đồng)": G_val, "Ký hiệu": "G", "Cơ sở hệ số / Ghi chú": "G = T + GT + TL"},
        {"STT": "IV", "Nội dung chi phí": "Thuế giá trị gia tăng", "Cách tính": f"=G * {st.session_state.coefficients['vat']['rate']}%", "Giá trị (đồng)": GTGT_val, "Ký hiệu": "GTGT", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['vat']['basis']},
        {"STT": "", "Nội dung chi phí": "CHI PHÍ XÂY DỰNG SAU THUẾ", "Cách tính": "=G+GTGT", "Giá trị (đồng)": Gxd_val, "Ký hiệu": "Gxd", "Cơ sở hệ số / Ghi chú": "Gxd = G + GTGT"},
        {"STT": "V", "Nội dung chi phí": "Chi phí nhà tạm để ở và điều hành thi công", "Cách tính": f"=G * {st.session_state.coefficients['lt']['rate']}% * (1+{st.session_state.coefficients['vat']['rate']}%)", "Giá trị (đồng)": LT_val, "Ký hiệu": "LT", "Cơ sở hệ số / Ghi chú": st.session_state.coefficients['lt']['basis']},
        {"STT": "", "Nội dung chi phí": "Tổng cộng", "Cách tính": "=Gxd+LT", "Giá trị (đồng)": TongCong, "Ký hiệu": "", "Cơ sở hệ số / Ghi chú": ""},
        {"STT": "", "Nội dung chi phí": "Làm tròn (đến hàng nghìn)", "Cách tính": "=ROUND(Tổng cộng, -3)", "Giá trị (đồng)": LamTron, "Ký hiệu": "", "Cơ sở hệ số / Ghi chú": ""}
    ]

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

        sel_stt = st.selectbox(
            "Chọn công việc cần kiểm tra chiết tính:",
            options=[item["stt"] for item in st.session_state.work_items],
            format_func=lambda x: f"Công tác {x}: {[it['job_name'] for it in st.session_state.work_items if it['stt'] == x][0]} ({[it['norm_code'] for it in st.session_state.work_items if it['stt'] == x][0]})"
        )
        it_sel = [it for it in st.session_state.work_items if it["stt"] == sel_stt][0]
        n_info = mgr.get_norm(it_sel["norm_code"])

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
