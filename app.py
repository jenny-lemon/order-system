# -*- coding: utf-8 -*-
import streamlit as st
from 儲值金系統設定 import run_process_web

st.set_page_config(page_title="儲值金訂單系統", page_icon="💰", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;600&family=DM+Sans:opsz,wght@9..40,300;9..40,400;9..40,500;9..40,600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', 'Noto Sans TC', sans-serif;
}
.block-container {
    padding-top: 1.4rem !important;
    padding-bottom: 1rem !important;
    max-width: 1100px !important;
}

/* 標題 */
h1 {
    font-size: 19px !important;
    font-weight: 600 !important;
    margin: 0 0 4px 0 !important;
}
.page-sub {
    font-size: 12px;
    color: #aaa;
    margin-bottom: 14px;
}

/* section label */
.sec-label {
    font-size: 11px;
    font-weight: 600;
    color: #aaa;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin: 0 0 6px 0;
}

/* Input / Select */
[data-testid="stTextInput"] label,
[data-testid="stNumberInput"] label,
[data-testid="stSelectbox"] label,
[data-testid="stMultiSelect"] label {
    font-size: 11.5px !important;
    color: #777 !important;
    font-weight: 500 !important;
    margin-bottom: 2px !important;
}
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input {
    font-size: 13.5px !important;
    padding: 7px 10px !important;
    border-radius: 7px !important;
    border-color: #dde0e5 !important;
    background: #fff !important;
}
[data-testid="stTextInput"] input:focus {
    border-color: #111 !important;
    box-shadow: 0 0 0 2px rgba(0,0,0,0.07) !important;
}
[data-testid="stSelectbox"] > div > div {
    border-radius: 7px !important;
    border-color: #dde0e5 !important;
    font-size: 13.5px !important;
    background: #fff !important;
}

/* Multiselect */
[data-testid="stMultiSelect"] > div > div {
    border-radius: 7px !important;
    border-color: #dde0e5 !important;
    font-size: 13px !important;
}

/* hint */
.hint-box {
    background: #f9fafb;
    border: 1px solid #e8eaed;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 12px;
    color: #888;
    margin-top: 4px;
    margin-bottom: 2px;
}

/* 執行按鈕 */
[data-testid="stButton"] > button {
    background: #111 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    font-size: 14px !important;
    font-weight: 600 !important;
    padding: 9px 0 !important;
    letter-spacing: 0.02em !important;
    transition: background 0.15s !important;
}
[data-testid="stButton"] > button:hover { background: #333 !important; }

/* Log */
[data-testid="stCode"] {
    font-size: 11.5px !important;
    border-radius: 8px !important;
    max-height: 280px;
    overflow-y: auto;
    background: #13131f !important;
}

/* Metric */
[data-testid="stMetric"] {
    background: #fff !important;
    border: 1px solid #e4e6ea !important;
    border-radius: 10px !important;
    padding: 14px 16px 12px !important;
    text-align: center !important;
}
[data-testid="stMetricLabel"] {
    font-size: 11px !important;
    color: #999 !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.06em !important;
}
[data-testid="stMetricValue"] {
    font-size: 30px !important;
    font-weight: 700 !important;
    color: #111 !important;
}

hr { border-color: #ebebeb !important; margin: 12px 0 !important; }

[data-testid="stAlert"] {
    border-radius: 8px !important;
    font-size: 13px !important;
    margin-top: 8px !important;
}
</style>
""", unsafe_allow_html=True)

# ── helpers ─────────────────────────────────────────────
def sec(title):
    st.markdown(f'<p class="sec-label">{title}</p>', unsafe_allow_html=True)

def parse_row_input(row_text: str):
    if not row_text or not row_text.strip():
        raise ValueError("請輸入列號，例如：2,3,5-7")
    rows = set()
    for part in [p.strip() for p in row_text.split(",") if p.strip()]:
        if "-" in part:
            s, e = part.split("-", 1)
            s, e = int(s.strip()), int(e.strip())
            if s <= 0 or e <= 0: raise ValueError("列號必須大於 0")
            if s > e: raise ValueError(f"區間錯誤：{part}")
            rows.update(range(s, e + 1))
        else:
            n = int(part)
            if n <= 0: raise ValueError("列號必須大於 0")
            rows.add(n)
    return sorted(rows)

# ── 標題 ─────────────────────────────────────────────────
st.title("💰 儲值金訂單系統")
st.markdown('<p class="page-sub">支援建單、寄確認信、改 Google 日曆，可指定列號批次處理。</p>', unsafe_allow_html=True)

# ── 帳密 + 環境（同列，與 memo 一致）────────────────────
col_e, col_p, col_env = st.columns([3.2, 3.2, 1.2])
with col_e:
    backend_email    = st.text_input("後台帳號")
with col_p:
    backend_password = st.text_input("後台密碼", type="password")
with col_env:
    env = st.selectbox("環境", ["prod", "dev"], index=0)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行設定 ──────────────────────────────────────────────
sec("執行設定")
c1, c2, c3 = st.columns([1.5, 2, 2])
with c1:
    region = st.selectbox("執行區域", ["台北", "台中", "桃園", "新竹", "高雄"])
with c2:
    sheet_name = st.text_input("工作表名稱", value="202604")
with c3:
    row_input = st.text_input("執行列號", value="2,3,5-7")

st.markdown('<div class="hint-box">💡 列號支援：單列 <code>2</code>、逗號分隔 <code>2,3,5</code>、區間 <code>2,3,5-7</code></div>', unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行項目 ──────────────────────────────────────────────
sec("執行項目")
selected_actions = st.multiselect(
    "執行項目",
    options=["建單", "寄確認信", "改 Google 日曆"],
    default=["建單", "寄確認信", "改 Google 日曆"],
    label_visibility="collapsed",
)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行按鈕 ──────────────────────────────────────────────
run_clicked = st.button("🚀  開始執行", use_container_width=True)

# ── 執行過程 ──────────────────────────────────────────────
with st.expander("📄  執行過程", expanded=True):
    log_box = st.empty()
    log_box.code("尚未執行")

# ── 執行結果 ──────────────────────────────────────────────
result_container = st.container()

# ── 執行邏輯 ─────────────────────────────────────────────
if run_clicked:
    # 驗證
    if not backend_email.strip():
        st.error("請輸入後台帳號"); st.stop()
    if not backend_password.strip():
        st.error("請輸入後台密碼"); st.stop()
    if not sheet_name.strip():
        st.error("請輸入工作表名稱"); st.stop()
    if not selected_actions:
        st.error("請至少選擇一個執行項目"); st.stop()

    try:
        target_rows = parse_row_input(row_input)
    except Exception as e:
        st.error(f"列號格式錯誤：{e}"); st.stop()

    logs = []
    def ui_log(msg):
        logs.append(str(msg))
        log_box.code("\n".join(logs[-80:]))

    total_success = total_fail = total_processed = 0

    with st.spinner("執行中，請稍候…"):
        for row_no in target_rows:
            ui_log(f"▶ 開始執行第 {row_no} 列…")
            try:
                result = run_process_web(
                    env_name=env,
                    region=region,
                    backend_email=backend_email.strip(),
                    backend_password=backend_password.strip(),
                    sheet_name=sheet_name.strip(),
                    start_row=row_no,
                    end_row=row_no,
                    selected_actions=selected_actions,
                    logger=ui_log,
                )
                if isinstance(result, dict):
                    total_success   += result.get("success_count", 0)
                    total_fail      += result.get("fail_count", 0)
                    total_processed += result.get("total_processed", 0)
            except Exception as e:
                total_fail += 1
                ui_log(f"❌ 第 {row_no} 列失敗：{e}")

    ui_log("===== 執行完成 =====")

    with result_container:
        st.markdown("<hr>", unsafe_allow_html=True)
        sec("執行結果")
        c1, c2, c3 = st.columns(3)
        c1.metric("執行筆數", total_processed)
        c2.metric("成功",     total_success)
        c3.metric("失敗",     total_fail)

        if total_fail == 0 and total_processed > 0:
            st.success(f"✅ 全部完成，共處理 **{total_processed}** 筆，成功 **{total_success}** 筆。")
        elif total_fail > 0:
            st.warning(f"⚠️ 執行完成，但有 **{total_fail}** 筆失敗，請查看執行過程。")
        else:
            st.info("執行完成，無資料被處理。")
