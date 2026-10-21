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
from core.excel_generator import ExcelGenerator

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

if "work_items" not in st.session_state:
    # Default 13 work items
    st.session_state.work_items = [
        {"stt": 1, "norm_code": "BD.41111", "job_name": "Thay mới lắp đặt thiết bị đầu báo khói địa chỉ", "description": "Dịch vụ Thay thiết bị đầu báo", "volume": 10.0, "unit": "Bộ"},
        {"stt": 2, "norm_code": "BD.41121", "job_name": "Lắp đặt đèn báo cháy", "description": "Lắp mới / thay thế đèn báo cháy vị trí khu vực phụ trợ", "volume": 5.0, "unit": "Bộ"},
        {"stt": 3, "norm_code": "BD.41131", "job_name": "Lắp đặt lại nút nhấn khẩn báo cháy", "description": "Lắp đặt lại nút nhấn sau khi bảo trì", "volume": 38.0, "unit": "Cái"},
        {"stt": 4, "norm_code": "BD.41141", "job_name": "Lắp đặt chuông báo cháy", "description": "Thay thế, lắp đặt chuông báo cháy tự động", "volume": 5.0, "unit": "Bộ"},
        {"stt": 5, "norm_code": "BD.42201", "job_name": "Lắp đặt bộ điều khiển, bộ chuyển mạch PCCC", "description": "Lắp đặt bộ chuyển mạch tín hiệu tủ trung gian", "volume": 1.0, "unit": "Bộ"},
        {"stt": 6, "norm_code": "E6C.3.01", "job_name": "Bảo trì, sửa chữa, thử nghiệm giả lập tác động nút nhấn khẩn", "description": "Kiểm tra, bảo dưỡng nút nhấn báo cháy chuyên dụng", "volume": 38.0, "unit": "Cái"},
        {"stt": 7, "norm_code": "E6C.3.02", "job_name": "Bảo trì, sửa chữa, hiệu chỉnh, thử nghiệm độ nhạy đầu báo khói/nhiệt", "description": "Dịch vụ Bảo trì đầu báo chuyên dụng", "volume": 890.0, "unit": "Bộ"},
        {"stt": 8, "norm_code": "E6C.3.03", "job_name": "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng dây báo cháy", "description": "Kiểm tra module dây giám sát nhiệt độ cao", "volume": 20.0, "unit": "Bộ"},
        {"stt": 9, "norm_code": "E6C.3.04", "job_name": "Bảo dưỡng bộ module giám sát tín hiệu cháy sử dụng beam quang", "description": "Bảo dưỡng beam khói khoảng cách xa", "volume": 15.0, "unit": "Bộ"},
        {"stt": 10, "norm_code": "E6C.3.05", "job_name": "Sửa chữa, bảo dưỡng trung tâm báo cháy FMZ 5000", "description": "Bảo dưỡng tủ điều khiển chữa cháy tại chỗ và tủ trung gian", "volume": 3.0, "unit": "trung tâm"},
        {"stt": 11, "norm_code": "E6C.3.06", "job_name": "Kiểm tra và bảo dưỡng phần điện van Deluge, Wet hệ PCCC", "description": "Bảo dưỡng cuộn hút, van xả tràn Deluge và đồng hồ áp", "volume": 12.0, "unit": "Cái"},
        {"stt": 12, "norm_code": "E6B.1.13.01", "job_name": "Kiểm tra bình accu chuyên dụng phục vụ hệ thống PCCC", "description": "Đo nội trở, dung lượng Ah, vô mỡ cọc bình và châm nước cất", "volume": 50.0, "unit": "Bình"},
        {"stt": 13, "norm_code": "E6C.1.31", "job_name": "Kiểm tra và bảo dưỡng thiết bị đo khí NH3, Clo, H2", "description": "Bơm khí chuẩn hiệu chuẩn điểm Zero & Span cho đầu đo khí", "volume": 4.0, "unit": "Bộ"}
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
    "💰 2. Quản Lý Đơn Giá",
    "🔍 3. Đơn Giá Chi Tiết Trực Quan",
    "🖼️ 4. Minh Chứng Scan Gốc",
    "🚀 5. Xuất File Excel Thẩm Định"
])

# -----------------------------------------------------------------------------
# TAB 1: KHỐI LƯỢNG TIÊN LƯỢNG
# -----------------------------------------------------------------------------
with tab1:
    st.markdown("#### BẢNG TỔNG HỢP KHỐI LƯỢNG CÔNG VIỆC (TIÊN LƯỢNG)")
    st.caption("Xem và điều chỉnh danh sách công tác, khối lượng, và mã định mức áp dụng cho dự án.")

    df_items = pd.DataFrame(st.session_state.work_items)
    
    # Configure data editor
    norm_codes_list = mgr.get_all_norm_codes()
    edited_df = st.data_editor(
        df_items,
        column_config={
            "stt": st.column_config.NumberColumn("STT", width="small", disabled=True),
            "norm_code": st.column_config.SelectboxColumn("Mã Định Mức", options=norm_codes_list, width="medium", required=True),
            "job_name": st.column_config.TextColumn("Danh mục công việc (Theo PAKT)", width="large", required=True),
            "description": st.column_config.TextColumn("Mô tả chi tiết", width="large"),
            "volume": st.column_config.NumberColumn("Khối lượng", format="%.2f", min_value=0.0, step=1.0, required=True),
            "unit": st.column_config.TextColumn("Đơn vị tính", width="small", required=True),
        },
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True
    )

    if st.button("💾 Cập nhật Bảng Khối Lượng"):
        st.session_state.work_items = edited_df.to_dict("records")
        st.success("Đã cập nhật bảng khối lượng công việc thành công!")

# -----------------------------------------------------------------------------
# TAB 2: QUẢN LÝ ĐƠN GIÁ (VẬT LIỆU, NHÂN CÔNG, MÁY THI CÔNG)
# -----------------------------------------------------------------------------
with tab2:
    st.markdown("#### QUẢN LÝ BẢNG GIÁ TÀI NGUYÊN (VẬT LIỆU, NHÂN CÔNG, MÁY THI CÔNG)")
    st.caption("Chỉnh sửa đơn giá tại đây. Khi thay đổi, Đơn giá chi tiết và File Excel xuất ra sẽ tự động nhận giá mới.")

    sub_t1, sub_t2, sub_t3 = st.tabs(["🧱 Vật Liệu", "👷 Nhân Công", "🚜 Máy Thi Công"])
    prices = mgr.get_resource_prices()

    with sub_t1:
        st.markdown("##### 1. ĐƠN GIÁ VẬT LIỆU (Theo giá thông báo địa phương / Liên sở)")
        df_vl = pd.DataFrame(prices.get("materials", []))
        if not df_vl.empty:
            edited_vl = st.data_editor(
                df_vl,
                column_config={
                    "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                    "name": st.column_config.TextColumn("Tên vật tư", disabled=True),
                    "unit": st.column_config.TextColumn("Đơn vị", disabled=True),
                    "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                    "price_notice": st.column_config.NumberColumn("Giá thông báo (đ)", format="%d"),
                    "notes": st.column_config.TextColumn("Ghi chú")
                },
                use_container_width=True,
                hide_index=True,
                key="editor_vl"
            )

    with sub_t2:
        st.markdown("##### 2. ĐƠN GIÁ NHÂN CÔNG (Căn cứ quyết định địa phương / EVN)")
        df_nc = pd.DataFrame(prices.get("labor", []))
        if not df_nc.empty:
            edited_nc = st.data_editor(
                df_nc,
                column_config={
                    "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                    "name": st.column_config.TextColumn("Tên nhân công", disabled=True),
                    "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                    "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                    "price_current": st.column_config.NumberColumn("Giá hiện tại (đ)", format="%d"),
                    "legal_basis": st.column_config.TextColumn("Căn cứ pháp lý")
                },
                use_container_width=True,
                hide_index=True,
                key="editor_nc"
            )

    with sub_t3:
        st.markdown("##### 3. ĐƠN GIÁ MÁY THI CÔNG (Theo định mức ca máy BXD TT 37/2026/TT-BXD)")
        df_m = pd.DataFrame(prices.get("machines", []))
        if not df_m.empty:
            edited_m = st.data_editor(
                df_m,
                column_config={
                    "code": st.column_config.TextColumn("Mã hiệu", disabled=True),
                    "name": st.column_config.TextColumn("Tên máy thi công", disabled=True),
                    "unit": st.column_config.TextColumn("ĐVT", disabled=True),
                    "price_base": st.column_config.NumberColumn("Giá gốc (đ)", format="%d"),
                    "price_notice": st.column_config.NumberColumn("Giá thông báo (đ)", format="%d"),
                    "price_diff": st.column_config.NumberColumn("Bù giá (đ)", format="%d"),
                    "price_current": st.column_config.NumberColumn("Giá hiện tại (đ)", format="%d"),
                    "legal_basis": st.column_config.TextColumn("Căn cứ pháp lý")
                },
                use_container_width=True,
                hide_index=True,
                key="editor_m"
            )

    col_btn1, col_btn2 = st.columns([2, 8])
    with col_btn1:
        if st.button("💾 Lưu Thay Đổi Bảng Giá", type="primary"):
            updated = {
                "materials": edited_vl.to_dict("records"),
                "labor": edited_nc.to_dict("records"),
                "machines": edited_m.to_dict("records")
            }
            mgr.update_resource_prices(updated)
            st.success("Đã lưu bảng giá tài nguyên thành công vào hệ thống!")

# -----------------------------------------------------------------------------
# TAB 3: ĐƠN GIÁ CHI TIẾT TRỰC QUAN (REALTIME EXPLORER)
# -----------------------------------------------------------------------------
with tab3:
    st.markdown("#### BẢNG ĐƠN GIÁ CHI TIẾT CỦA TỪNG CÔNG TÁC (REAL-TIME)")
    st.caption("Xem bảng phân tích chiết tính đơn giá cho 1 đơn vị công tác (Vật liệu + Nhân công + Máy thi công).")

    # Map prices
    vl_price_map = {m["code"]: m.get("price_notice", 0) for m in mgr.get_resource_prices().get("materials", [])}
    nc_price_map = {l["code"]: l.get("price_current", 0) for l in mgr.get_resource_prices().get("labor", [])}
    m_price_map = {mac["code"]: mac.get("price_current", 0) for mac in mgr.get_resource_prices().get("machines", [])}

    selected_stt = st.selectbox(
        "Chọn công việc để xem chiết tính chi tiết:",
        options=[item["stt"] for item in st.session_state.work_items],
        format_func=lambda x: f"Công tác {x}: {[it['job_name'] for it in st.session_state.work_items if it['stt'] == x][0]} ({[it['norm_code'] for it in st.session_state.work_items if it['stt'] == x][0]})"
    )

    selected_item = [it for it in st.session_state.work_items if it["stt"] == selected_stt][0]
    norm_info = mgr.get_norm(selected_item["norm_code"])

    if norm_info:
        col_c1, col_c2, col_c3 = st.columns([4, 2, 2])
        with col_c1:
            st.markdown(f"**Tên công việc:** {selected_item['job_name']}")
            st.caption(f"Căn cứ: {norm_info.get('legal_basis', '')}")
        with col_c2:
            st.metric("Khối lượng", f"{selected_item['volume']:,.2f} {selected_item['unit']}")
        with col_c3:
            st.metric("Định mức gốc tính cho", norm_info.get("unit", ""))

        # Build table rows for this norm
        rows = []
        tot_vl = 0.0
        tot_nc = 0.0
        tot_m = 0.0

        res = norm_info.get("resources", {})
        # Materials
        for m in res.get("materials", []):
            code = m["code"]
            dg = vl_price_map.get(code, 0)
            dm = m["rate"]
            tt = dm * dg if code != "V00750" else 0 # handled below
            rows.append({"Phân loại": "Vật liệu", "Mã hiệu": code, "Tên tài nguyên": m["name"], "ĐVT": m["unit"], "Định mức": dm, "Đơn giá (đ)": dg, "Thành tiền (đ)": tt})
            tot_vl += tt

        # Labor
        for l in res.get("labor", []):
            code = l["code"]
            dg = nc_price_map.get(code, 0)
            dm = l["rate"]
            tt = dm * dg
            rows.append({"Phân loại": "Nhân công", "Mã hiệu": code, "Tên tài nguyên": l["name"], "ĐVT": l["unit"], "Định mức": dm, "Đơn giá (đ)": dg, "Thành tiền (đ)": tt})
            tot_nc += tt

        # Machines
        for mac in res.get("machines", []):
            code = mac["code"]
            dg = m_price_map.get(code, 0)
            dm = mac["rate"]
            tt = dm * dg
            rows.append({"Phân loại": "Máy thi công", "Mã hiệu": code, "Tên tài nguyên": mac["name"], "ĐVT": mac["unit"], "Định mức": dm, "Đơn giá (đ)": dg, "Thành tiền (đ)": tt})
            tot_m += tt

        total_unit_rate = tot_vl + tot_nc + tot_m
        df_dt = pd.DataFrame(rows)
        st.dataframe(
            df_dt.style.format({
                "Định mức": "{:,.2f}",
                "Đơn giá (đ)": "{:,.0f}",
                "Thành tiền (đ)": "{:,.0f}"
            }),
            use_container_width=True,
            hide_index=True
        )

        st.markdown(f"""
        <div class="metric-card">
            <b>TỔNG ĐƠN GIÁ CHI TIẾT:</b> <span style="font-size:20px; color:#002060; font-weight:bold;">{total_unit_rate:,.0f} đ</span> 
            (Vật liệu: {tot_vl:,.0f} đ | Nhân công: {tot_nc:,.0f} đ | Máy: {tot_m:,.0f} đ)
        </div>
        """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# TAB 4: MINH CHỨNG SCAN GỐC (EVIDENCE BROWSER)
# -----------------------------------------------------------------------------
with tab4:
    st.markdown("#### KHO HỒ SƠ MINH CHỨNG SCAN ĐỊNH MỨC GỐC")
    st.caption("Xem trang PDF gốc và cơ cấu các bước thực hiện định mức (phục vụ thẩm định/thanh kiểm tra).")

    sel_code = st.selectbox("Chọn mã định mức cần tra cứu minh chứng:", options=mgr.get_all_norm_codes())
    norm_detail = mgr.get_norm(sel_code)

    if norm_detail:
        col_img, col_steps = st.columns([5, 5])
        with col_img:
            st.markdown(f"**Ảnh scan trích xuất từ tài liệu gốc ({sel_code}):**")
            img_p = mgr.get_evidence_image_path(sel_code)
            if img_p and img_p.exists():
                st.image(str(img_p), caption=f"Minh chứng pháp lý: {norm_detail.get('legal_basis', '')}", use_container_width=True)
            else:
                st.warning("Chưa có ảnh scan minh chứng cho mã này.")

        with col_steps:
            st.markdown("**Bảng phân tích tỷ lệ các bước thực hiện:**")
            st.caption(f"Đơn vị tính gốc: {norm_detail.get('unit', '')} | Căn cứ: {norm_detail.get('legal_basis', '')}")
            steps = norm_detail.get("work_steps", [])
            if steps:
                df_steps = pd.DataFrame(steps)
                st.dataframe(
                    df_steps[["step", "name", "dvt_step", "rate_base", "apply_flag", "rate_calculated"]].rename(columns={
                        "step": "Bước", "name": "Nội dung công việc", "dvt_step": "ĐVT",
                        "rate_base": "Hao phí gốc", "apply_flag": "Áp dụng (1/0)", "rate_calculated": "Tính toán"
                    }),
                    use_container_width=True,
                    hide_index=True
                )
                tot_step = sum(s.get("rate_calculated", 0) for s in steps)
                st.info(f"Tổng hao phí nhân công các bước: **{tot_step:.4f} công**")

# -----------------------------------------------------------------------------
# TAB 5: XUẤT BẢN FILE EXCEL DỰ TOÁN
# -----------------------------------------------------------------------------
with tab5:
    st.markdown("#### XUẤT BẢN FILE EXCEL DỰ TOÁN THẨM ĐỊNH")
    st.caption("Tạo file Excel đầy đủ 19 sheet với 100% công thức liên kết động (Dynamic Links).")

    st.write(f"📁 **Vị trí lưu trữ đầu ra:** `{OUTPUT_DIR}`")
    export_filename = st.text_input("Tên tệp Excel xuất ra:", value="DU_TOAN_PCCC_2026_THAM_DINH.xlsx")

    if st.button("🚀 BẮT ĐẦU TẠO FILE EXCEL", type="primary", use_container_width=True):
        out_target = OUTPUT_DIR / export_filename
        with st.spinner("Đang tổng hợp dữ liệu, chèn ảnh scan và thiết lập 100% công thức liên kết động..."):
            try:
                generator.generate(
                    project_name=project_name,
                    category_name=category_name,
                    work_items=st.session_state.work_items,
                    output_path=out_target
                )
                st.success(f"✅ ĐÃ TẠO FILE EXCEL THÀNH CÔNG: {out_target.name}")
                
                # Direct download button in browser
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
