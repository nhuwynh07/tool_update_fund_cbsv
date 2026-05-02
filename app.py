"""
app.py – Giao diện web Streamlit cho hệ thống đối soát thu chi bộ.

Khởi chạy: streamlit run app.py
"""
import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

# Thêm thư mục gốc vào sys.path để import các module backend
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import (
    SHEET_QUY, QUY_MONTH_COLS, QUY_STANDARD_MONTH_AMOUNT,
    DEFAULT_COLLECTION_QUARTER,
)
from text_utils import parse_money
from statement_reader import read_members
from matcher import months_of_quarter
from main import run_core

# ─────────────────────────────────────────────────────────────────────────────
# Cấu hình trang
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Đối Soát Thu Chi Bộ",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CSS tùy chỉnh
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Google Font */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Hero header */
.hero {
    background: linear-gradient(135deg, #1a237e 0%, #283593 50%, #3949ab 100%);
    border-radius: 16px;
    padding: 32px 40px;
    margin-bottom: 28px;
    color: white;
    box-shadow: 0 8px 32px rgba(26,35,126,0.3);
}
.hero h1 { font-size: 2rem; font-weight: 700; margin: 0 0 6px 0; letter-spacing: -0.5px; }
.hero p  { font-size: 1rem; opacity: 0.85; margin: 0; }

/* Metric cards */
.metric-row { display: flex; gap: 16px; margin-bottom: 24px; }
.metric-card {
    flex: 1;
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
    border-left: 4px solid;
    transition: transform .15s;
}
.metric-card:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(0,0,0,0.12); }
.metric-card.blue   { border-color: #3949ab; }
.metric-card.green  { border-color: #2e7d32; }
.metric-card.red    { border-color: #c62828; }
.metric-card.orange { border-color: #e65100; }
.metric-card .label { font-size: .8rem; color: #666; font-weight: 500; text-transform: uppercase; letter-spacing: .5px; }
.metric-card .value { font-size: 1.9rem; font-weight: 700; margin-top: 4px; }
.metric-card.blue   .value { color: #3949ab; }
.metric-card.green  .value { color: #2e7d32; }
.metric-card.red    .value { color: #c62828; }
.metric-card.orange .value { color: #e65100; }

/* Upload boxes */
.upload-section {
    background: #f8f9ff;
    border: 2px dashed #c5cae9;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 12px;
}

/* Status badges */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: .78rem;
    font-weight: 600;
}
.badge-success { background: #e8f5e9; color: #2e7d32; }
.badge-danger  { background: #ffebee; color: #c62828; }
.badge-warning { background: #fff8e1; color: #e65100; }
.badge-info    { background: #e3f2fd; color: #1565c0; }

/* Tab styling */
.stTabs [data-baseweb="tab-list"] { gap: 8px; }
.stTabs [data-baseweb="tab"] {
    border-radius: 8px 8px 0 0;
    font-weight: 500;
    padding: 8px 20px;
}

/* Run button */
# div.stButton > button {
#     background: linear-gradient(135deg, #1a237e, #3949ab);
#     color: white;
#     border: none;
#     border-radius: 10px;
#     padding: 12px 32px;
#     font-size: 1rem;
#     font-weight: 600;
#     letter-spacing: .3px;
#     width: 100%;
#     transition: all .2s;
# }
# div.stButton > button:hover {
#     background: linear-gradient(135deg, #283593, #5c6bc0);
#     box-shadow: 0 4px 16px rgba(26,35,126,0.35);
#     transform: translateY(-1px);
# }

div.stButton > button {
    background: linear-gradient(135deg, #283593 0%, #3949ab 55%, #5c6bc0 100%) !important;
    color: #ffffff !important;
    border: 1px solid rgba(255,255,255,0.08) !important;
    box-shadow: 0 6px 18px rgba(57,73,171,0.28);
}

div.stButton > button:disabled,
div.stButton > button[disabled] {
    background: #e8eaf6 !important;
    color: #3949ab !important;
    border: 1px solid #c5cae9 !important;
    opacity: 1 !important;
}


/* Download button */
div[data-testid="stDownloadButton"] > button {
    background: #2e7d32;
    color: white;
    border: none;
    border-radius: 10px;
    padding: 10px 28px;
    font-weight: 600;
    width: 100%;
}
div[data-testid="stDownloadButton"] > button:hover {
    background: #1b5e20;
}

/* Section title */
.section-title {
    font-size: 1.05rem;
    font-weight: 700;
    color: #1a237e;
    margin: 20px 0 12px 0;
    padding-bottom: 6px;
    border-bottom: 2px solid #e8eaf6;
}

/* Sidebar */
section[data-testid="stSidebar"] > div {
    background: linear-gradient(180deg, #1a237e 0%, #283593 100%);
    padding-top: 16px;
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stFileUploader label { color: #c5cae9 !important; }

/* Remove white background from file uploader in sidebar */
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section {
    background-color: transparent !important;
    border: 1px dashed rgba(255, 255, 255, 0.4) !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] section > div {
    background-color: transparent !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
    background-color: transparent !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div {
    background-color: transparent !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
    font-size: 0 !important;
    color: transparent !important;
    background-color: #ffffff !important;
    border: 1px solid #c5cae9 !important;
    border-radius: 8px !important;
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    padding: 0.5rem 1rem !important;
    transition: all 0.2s;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button * {
    display: none !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button::after {
    content: "Chọn file";
    font-size: 14px !important;
    color: #1a237e !important;
    font-weight: 600 !important;
    display: block !important;
}
section[data-testid="stSidebar"] [data-testid="stFileUploader"] button:hover {
    background-color: #e8eaf6 !important;
    border-color: #3949ab !important;
}


/* Translate Drag and drop text to Vietnamese */
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div > div > span {
    display: none;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div > div::before {
    content: "Kéo và thả file vào đây";
    color: white;
    font-size: 14px;
    display: block;
    margin-bottom: 2px;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div > div > small {
    display: none;
}
section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] > div > div::after {
    content: "Giới hạn 200MB/file • XLSX, XLS";
    color: #c5cae9;
    font-size: 12px;
    display: block;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# Helper: đọc trạng thái đảng viên từ workbook đã xử lý
# ─────────────────────────────────────────────────────────────────────────────
def build_member_status(wb, members_quy: list) -> pd.DataFrame:
    """
    Đọc sheet Quỹ sau khi xử lý, trả về DataFrame trạng thái đóng phí
    của từng đảng viên theo từng tháng trong quý thu mặc định.
    """
    ws = wb[SHEET_QUY]
    q_months = months_of_quarter(DEFAULT_COLLECTION_QUARTER)
    records  = []

    for m in members_quy:
        row_data = {
            "TT":     m["tt"],
            "Họ và tên": m["name"],
        }
        paid_count = 0
        for month in q_months:
            col = QUY_MONTH_COLS.get(month)
            val = parse_money(ws.cell(m["row"], col).value) if col else 0
            row_data[f"T{month}"] = "✅" if val == QUY_STANDARD_MONTH_AMOUNT else (
                f"{val:,}đ" if val > 0 else "❌"
            )
            if val == QUY_STANDARD_MONTH_AMOUNT:
                paid_count += 1

        if paid_count == len(q_months):
            row_data["Trạng thái Q" + str(DEFAULT_COLLECTION_QUARTER)] = "Đã đóng đủ"
        elif paid_count > 0:
            row_data["Trạng thái Q" + str(DEFAULT_COLLECTION_QUARTER)] = f"Còn thiếu {len(q_months)-paid_count} tháng"
        else:
            row_data["Trạng thái Q" + str(DEFAULT_COLLECTION_QUARTER)] = "Chưa đóng"

        records.append(row_data)

    return pd.DataFrame(records)


def color_status(val: str) -> str:
    """Tô màu cột trạng thái trong DataFrame."""
    if val == "Đã đóng đủ":
        return "background-color:#e8f5e9; color:#2e7d32; font-weight:600"
    if val == "Chưa đóng":
        return "background-color:#ffebee; color:#c62828; font-weight:600"
    return "background-color:#fff8e1; color:#e65100; font-weight:600"


def color_log_type(val: str) -> str:
    mapping = {
        "quy":     "background-color:#e3f2fd; color:#1565c0",
        "phat":    "background-color:#ffebee; color:#c62828",
        "ignore":  "background-color:#f3e5f5; color:#6a1b9a",
        "unknown": "background-color:#fff8e1; color:#e65100",
    }
    return mapping.get(str(val).lower(), "")


# ─────────────────────────────────────────────────────────────────────────────
# Sidebar – Upload files
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏛️ Quản lý quỹ CBSV1")
    st.markdown("---")
    st.markdown("### 📂 Tải lên file dữ liệu")

    st.markdown("**File sao kê ngân hàng**")
    sao_ke_file = st.file_uploader(
        "Sao kê (Excel)",
        type=["xlsx", "xls"],
        key="sao_ke",
        label_visibility="collapsed",
    )
    if sao_ke_file:
        st.success(f"✓ {sao_ke_file.name}")

    st.markdown("**File báo cáo Đảng phí**")
    tong_hop_file = st.file_uploader(
        "Báo cáo (Excel)",
        type=["xlsx", "xls"],
        key="tong_hop",
        label_visibility="collapsed",
    )
    if tong_hop_file:
        st.success(f"✓ {tong_hop_file.name}")

    st.markdown("---")
    st.markdown("#### Hướng dẫn")
    st.markdown("""
1. Tải lên **file sao kê** ngân hàng (Techcombank)
2. Tải lên **file báo cáo** tổng hợp đảng phí
3. Bấm **Chạy đối soát**
4. Xem kết quả và tải file đã cập nhật
""")


# ─────────────────────────────────────────────────────────────────────────────
# Main area
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>🏛️ Phần mềm quản lý quỹ CBSV1</h1>
    <p>Đối soát sao kê ngân hàng với báo cáo đảng phí & Cập nhật trạng thái đóng đảng phí và quỹ chi bộ</p>
</div>
""", unsafe_allow_html=True)

# Nút chạy
both_uploaded = sao_ke_file is not None and tong_hop_file is not None
if not both_uploaded:
    st.info("⬅️ Vui lòng tải lên đủ hai file ở thanh bên trái để bắt đầu.")

col_run, col_space = st.columns([1, 2])
with col_run:
    run_clicked = st.button(
        "⚡ Chạy đối soát",
        disabled=not both_uploaded,
        use_container_width=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# Xử lý khi bấm nút
# ─────────────────────────────────────────────────────────────────────────────
if run_clicked and both_uploaded:
    with st.spinner("Đang xử lý dữ liệu..."):
        try:
            result = run_core(
                io.BytesIO(sao_ke_file.getvalue()),
                io.BytesIO(tong_hop_file.getvalue()),
            )
            st.session_state["result"]        = result
            st.session_state["member_status"] = build_member_status(
                result["workbook"], result["members_quy"]
            )
            st.success("✅ Đối soát hoàn tất!")
        except Exception as e:
            st.error(f"❌ Lỗi khi xử lý: {e}")
            st.exception(e)

# ─────────────────────────────────────────────────────────────────────────────
# Hiển thị kết quả
# ─────────────────────────────────────────────────────────────────────────────
if "result" in st.session_state:
    result        = st.session_state["result"]
    member_status = st.session_state["member_status"]
    logs_df       = result["logs"]

    # ── Thống kê tổng quan ──────────────────────────────────────────────────
    status_col = f"Trạng thái Q{DEFAULT_COLLECTION_QUARTER}"
    n_total    = len(member_status)
    n_paid     = (member_status[status_col] == "Đã đóng đủ").sum()
    n_partial  = member_status[status_col].str.startswith("Còn thiếu").sum()
    n_unpaid   = (member_status[status_col] == "Chưa đóng").sum()

    # Tổng tiền quỹ thu được (chỉ tính giao dịch khớp được tên)
    total_quy = logs_df.loc[
        (logs_df["Loại"] == "quy") & (logs_df["Tên khớp"] != ""),
        "Số tiền"
    ].sum()

    st.markdown(f"""
    <div class="metric-row">
        <div class="metric-card blue">
            <div class="label">Tổng giao dịch</div>
            <div class="value">{result['n_transactions']}</div>
        </div>
        <div class="metric-card green">
            <div class="label">Đã đóng đủ Q{DEFAULT_COLLECTION_QUARTER}</div>
            <div class="value">{n_paid} / {n_total}</div>
        </div>
        <div class="metric-card red">
            <div class="label">Chưa đóng</div>
            <div class="value">{n_unpaid}</div>
        </div>
        <div class="metric-card orange">
            <div class="label">Tổng tiền quỹ thu được</div>
            <div class="value">{total_quy:,.0f}đ</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Tabs nội dung ───────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        "📋 Log giao dịch",
        "👥 Trạng thái đảng viên",
        "📥 Tải xuống",
    ])

    # ── Tab 1: Log giao dịch ────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-title">Nhật ký đối soát giao dịch</div>', unsafe_allow_html=True)

        c1, c2, c3 = st.columns([3, 2, 2])
        with c1:
            search_name = st.text_input("🔍 Tìm theo tên", placeholder="Nhập họ tên...")
        with c2:
            filter_type = st.selectbox("Lọc loại", ["Tất cả", "quy", "phat", "ignore", "unknown"])
        with c3:
            filter_matched = st.selectbox("Khớp tên", ["Tất cả", "Đã khớp", "Chưa khớp"])

        df_show = logs_df.copy()
        if search_name:
            df_show = df_show[df_show["Tên khớp"].str.contains(search_name, case=False, na=False)]
        if filter_type != "Tất cả":
            df_show = df_show[df_show["Loại"] == filter_type]
        if filter_matched == "Đã khớp":
            df_show = df_show[df_show["Tên khớp"] != ""]
        elif filter_matched == "Chưa khớp":
            df_show = df_show[df_show["Tên khớp"] == ""]

        st.caption(f"Hiển thị {len(df_show)} / {len(logs_df)} giao dịch")

        styled = (
            df_show.reset_index(drop=True)
            .style
            .applymap(color_log_type, subset=["Loại"])
            .format({"Số tiền": "{:,.0f}", "Điểm": lambda x: f"{x:.1f}" if x != "" else ""})
        )
        st.dataframe(styled, use_container_width=True, height=420)

    # ── Tab 2: Trạng thái đảng viên ─────────────────────────────────────────
    with tab2:
        st.markdown(
            f'<div class="section-title">Trạng thái đóng phí Quý {DEFAULT_COLLECTION_QUARTER}</div>',
            unsafe_allow_html=True,
        )

        c1, c2 = st.columns([3, 2])
        with c1:
            search_member = st.text_input("🔍 Tìm đảng viên", placeholder="Nhập họ tên...", key="s2")
        with c2:
            filter_status = st.selectbox(
                "Lọc trạng thái",
                ["Tất cả", "Đã đóng đủ", "Còn thiếu", "Chưa đóng"],
                key="fs2",
            )

        ms_show = member_status.copy()
        if search_member:
            ms_show = ms_show[ms_show["Họ và tên"].str.contains(search_member, case=False, na=False)]
        if filter_status == "Còn thiếu":
            ms_show = ms_show[ms_show[status_col].str.startswith("Còn thiếu")]
        elif filter_status != "Tất cả":
            ms_show = ms_show[ms_show[status_col] == filter_status]

        st.caption(
            f"Đã đóng đủ: **{n_paid}** · Còn thiếu: **{n_partial}** · Chưa đóng: **{n_unpaid}**"
        )

        styled_ms = (
            ms_show.reset_index(drop=True)
            .style
            .applymap(color_status, subset=[status_col])
        )
        st.dataframe(styled_ms, use_container_width=True, height=460)

        # Progress bar
        pct = n_paid / n_total * 100 if n_total > 0 else 0
        st.markdown(f"**Tỷ lệ đóng đủ Q{DEFAULT_COLLECTION_QUARTER}:** {pct:.1f}%")
        st.progress(int(pct))

    # ── Tab 3: Tải xuống ────────────────────────────────────────────────────
    with tab3:
        st.markdown('<div class="section-title">Tải xuống file Excel đã cập nhật</div>', unsafe_allow_html=True)

        out_name = tong_hop_file.name.replace(".xlsx", "_UPDATED.xlsx")

        col_dl, col_info = st.columns([1, 2])
        with col_dl:
            st.download_button(
                label="📥 Tải file Excel đã cập nhật",
                data=result["excel_bytes"],
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_info:
            st.markdown(f"""
            **File:** `{out_name}`

            File này bao gồm:
            - ✅ Sheet **{SHEET_QUY}** – đã cập nhật trạng thái nộp quỹ
            - ✅ Sheet **DoiSoat_Log** – nhật ký toàn bộ giao dịch đối soát
            """)

        st.markdown("---")
        st.markdown("**Tải xuống bảng log (CSV)**")
        csv_bytes = logs_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="📄 Tải log CSV",
            data=csv_bytes,
            file_name="DoiSoat_Log.csv",
            mime="text/csv",
        )
