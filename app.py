"""
app.py  –  Phiên bản Streamlit (web)
Deploy lên Streamlit Cloud:
  1. Push repo lên GitHub (gồm app.py, teacher_core.py, requirements.txt)
  2. Vào streamlit.io/cloud → New app → chọn repo → main file = app.py
  3. Thêm secret ANTHROPIC_API_KEY trong Settings → Secrets
"""

import io
import streamlit as st
import anthropic

from teacher_core import process_data

NIEN_KHOA_OPTIONS = ["2025-2026", "2026-2027", "2027-2028"]

# ── Cấu hình trang ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Convert PCCM",
    page_icon="🏫",
    layout="centered",
)

# ── CSS tùy chỉnh ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F4E79, #2E75B6);
        color: white;
        padding: 1.5rem 2rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        text-align: center;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; }
    .main-header p  { margin: 0.4rem 0 0; opacity: 0.85; font-size: 0.95rem; }
    .step-box {
        background: #f0f4fa;
        border-left: 4px solid #2E75B6;
        padding: 0.8rem 1rem;
        border-radius: 0 8px 8px 0;
        margin-bottom: 1rem;
    }
    .success-box {
        background: #e8f5e9;
        border-left: 4px solid #43a047;
        padding: 0.8rem 1rem;
        border-radius: 0 8px 8px 0;
    }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>🏫 CONVERT PCCM</h1>
  <p>File đầu vào cần có sheet Data</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar: API key ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Cài đặt")
    api_key_input = st.text_input(
        "Anthropic API Key",
        type="password",
        placeholder="sk-ant-...",
        help="Cần để nhận dạng tên môn học không chuẩn bằng AI. "
             "Trên Streamlit Cloud, đặt secret ANTHROPIC_API_KEY thay vì nhập ở đây.",
    )
    st.markdown("---")
    st.markdown("**Hướng dẫn deploy:**")
    st.markdown("""
1. Push code lên GitHub
2. Vào [streamlit.io/cloud](https://streamlit.io/cloud)
3. New app → chọn repo → `app.py`
4. Settings → Secrets → thêm `ANTHROPIC_API_KEY`
""")

# ── Lấy API key (ưu tiên secrets, sau đó input) ───────────────────────────────
try:
    effective_api_key = st.secrets["ANTHROPIC_API_KEY"]
except Exception:
    effective_api_key = api_key_input.strip() or None

# ── Form nhập liệu ────────────────────────────────────────────────────────────
st.markdown('<div class="step-box"><b>Bước 1:</b> Tải lên file Excel chứa sheet <code>Data</code></div>',
            unsafe_allow_html=True)
uploaded = st.file_uploader(
    "Chọn file Excel (.xlsx / .xls)",
    type=["xlsx", "xls", "xlsm"],
    label_visibility="collapsed",
)

st.markdown('<div class="step-box"><b>Bước 2:</b> Chọn niên khóa</div>',
            unsafe_allow_html=True)
nien_khoa = st.selectbox(
    "Niên khóa",
    options=NIEN_KHOA_OPTIONS,
    label_visibility="collapsed",
)

st.markdown('<div class="step-box"><b>Bước 3:</b> Nhấn nút để xử lý</div>',
            unsafe_allow_html=True)
run_btn = st.button("▶  Chuyển đổi", type="primary", use_container_width=True,
                    disabled=(uploaded is None or not effective_api_key))

if uploaded is None:
    st.info("📂 Vui lòng tải lên file Excel đầu vào.")
if not effective_api_key:
    st.warning("🔑 Vui lòng nhập Anthropic API Key (sidebar) hoặc đặt secret `ANTHROPIC_API_KEY`.")

# ── Xử lý ─────────────────────────────────────────────────────────────────────
if run_btn and uploaded and effective_api_key:
    log_area   = st.empty()
    prog_bar   = st.progress(0)
    log_lines  = []

    # Đếm tổng số GV để tính %
    import pandas as pd
    from teacher_core import detect_header_row, find_column
    raw_bytes = uploaded.read()
    try:
        _xl = pd.ExcelFile(io.BytesIO(raw_bytes))
        _sn = next((s for s in _xl.sheet_names if s.strip().lower() == "data"), _xl.sheet_names[0])
        _rdf = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=_sn, header=None)
        _hri = detect_header_row(_rdf)
        _df  = pd.read_excel(io.BytesIO(raw_bytes), sheet_name=_sn, header=_hri)
        _df.columns = [str(c).strip() for c in _df.columns]
        _ch  = find_column(_df, ["họ tên", "họ và tên", "tên", "giáo viên", "ho ten"])
        total_teachers = len(_df[_df[_ch].notna()]) if _ch else 1
    except Exception:
        total_teachers = 1

    processed = [0]

    def progress_cb(msg):
        log_lines.append(msg)
        log_area.code("\n".join(log_lines[-20:]), language=None)
        # Cập nhật progress bar khi xử lý từng GV
        if "Xử lý giáo viên" in msg:
            processed[0] += 1
            pct = min(int(processed[0] / total_teachers * 90), 90)
            prog_bar.progress(pct)

    try:
        client = anthropic.Anthropic(api_key=effective_api_key)
        result_bytes = process_data(
            io.BytesIO(raw_bytes), nien_khoa, client, progress_cb=progress_cb
        )
        prog_bar.progress(100)

        filename = uploaded.name.replace(".xlsx", "").replace(".xls", "")
        out_name = f"{filename}_output_{nien_khoa}.xlsx"

        st.markdown('<div class="success-box">✅ <b>Chuyển đổi thành công!</b> Nhấn nút bên dưới để tải về.</div>',
                    unsafe_allow_html=True)
        st.download_button(
            label="⬇️  Tải xuống file Excel",
            data=result_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    except Exception as e:
        prog_bar.empty()
        st.error(f"❌ Lỗi xử lý: {e}")
