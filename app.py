# -*- coding: utf-8 -*-
import streamlit as st
from 儲值金系統設定 import run_process_web

st.set_page_config(page_title="儲值金訂單系統", page_icon="💰", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Space+Grotesk:wght@500;700&display=swap');

:root {
    --lemon: #F5C518;
    --lemon-dark: #D4A017;
    --lemon-soft: #FFFBEA;
    --lemon-mid: #FFF3C4;
    --charcoal: #1C1C1E;
    --ink: #3A3A3C;
    --muted: #8E8E93;
    --border: #E5E5EA;
    --surface: #FFFFFF;
    --success: #34C759;
    --danger: #FF3B30;
    --radius: 14px;
}

html, body, [class*="css"] {
    font-family: 'Noto Sans TC', sans-serif;
    color: var(--charcoal);
}

#MainMenu, footer, header {
    visibility: hidden;
}

[data-testid="stAppViewContainer"] {
    background: #FAFAFA;
}

.block-container {
    padding-top: 2rem !important;
    padding-bottom: 2rem !important;
    max-width: 1180px !important;
}

/* Hero */

.hero {
    background: linear-gradient(135deg, #FFFDF0 0%, #FFFBEA 100%);
    border: 1.5px solid var(--lemon-mid);
    border-radius: var(--radius);
    padding: 2rem 2.5rem 1.6rem;
    margin-bottom: 2rem;
    display: flex;
    align-items: center;
    gap: 1.2rem;
    box-shadow: 0 2px 12px rgba(245,197,24,0.10);
}

.hero-emoji {
    font-size: 3rem;
    line-height: 1;
}

.hero-title {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.9rem;
    font-weight: 700;
    color: var(--charcoal);
    letter-spacing: -0.5px;
}

.hero-sub {
    color: var(--ink);
    font-size: 0.92rem;
    margin-top: 0.3rem;
    opacity: 0.78;
}

/* Step */

.step-pill {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    background: var(--lemon-mid);
    border: 1.5px solid var(--lemon);
    border-radius: 30px;
    padding: 0.28rem 0.9rem;
    font-size: 0.78rem;
    font-weight: 700;
    color: var(--charcoal);
    margin-bottom: 0.9rem;
    letter-spacing: 0.02em;
}

.step-num {
    background: var(--lemon);
    border-radius: 50%;
    width: 20px;
    height: 20px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 0.72rem;
    font-weight: 700;
}

.sec-label {
    font-size: 12px;
    font-weight: 700;
    color: var(--muted);
    letter-spacing: .04em;
    margin-bottom: 8px;
}

.hint-box {
    background: var(--lemon-soft);
    border-left: 4px solid var(--lemon);
    border-radius: 0 8px 8px 0;
    padding: 0.75rem 1rem;
    font-size: 0.9rem;
    color: var(--ink);
    margin-top: 0.6rem;
}

/* Labels */

[data-testid="stTextInput"] label,
[data-testid="stNumberInput"] label,
[data-testid="stSelectbox"] label,
[data-testid="stMultiSelect"] label {
    font-size: 13px !important;
    color: var(--ink) !important;
    font-weight: 700 !important;
}

/* Inputs */

[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stSelectbox"] > div > div,
[data-testid="stMultiSelect"] > div > div {
    border-radius: 10px !important;
    border: 1.5px solid var(--border) !important;
    background: white !important;
    font-size: 15px !important;
}

[data-testid="stTextInput"] input:focus {
    border-color: var(--lemon-dark) !important;
    box-shadow: 0 0 0 2px rgba(245,197,24,0.22) !important;
}

/* Buttons */

[data-testid="stButton"] > button {
    background: var(--lemon) !important;
    color: var(--charcoal) !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 15px !important;
    font-weight: 700 !important;
    padding: 0.55rem 1.2rem !important;
    box-shadow: 0 2px 10px rgba(245,197,24,0.28) !important;
}

[data-testid="stButton"] > button:hover {
    background: var(--lemon-dark) !important;
    transform: translateY(-1px) !important;
}

[data-testid="stButton"] > button:disabled {
    background: #D1D5DB !important;
    color: #777 !important;
}

/* Expander */

[data-testid="stExpander"] {
    border: 1px solid #ececec !important;
    border-radius: 14px !important;
    background: white !important;
    overflow: hidden !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.04);
}

[data-testid="stExpander"] summary {
    font-size: 14px !important;
    font-weight: 700 !important;
    color: var(--charcoal) !important;
    padding: 12px 16px !important;
}

/* Code */

[data-testid="stCode"] {
    font-size: 13px !important;
    min-height: 420px !important;
    max-height: 520px !important;
    overflow-y: auto !important;
    background: #1C1C1E !important;
    margin: 0 !important;
    white-space: pre-wrap !important;
}

/* Metrics */

[data-testid="stMetric"] {
    background: white !important;
    border: 1px solid #ececec !important;
    border-radius: 14px !important;
    padding: 14px 16px !important;
    text-align: center !important;
    box-shadow: 0 2px 12px rgba(0,0,0,0.04);
}

[data-testid="stMetricLabel"] {
    font-size: 12px !important;
    color: var(--muted) !important;
    font-weight: 700 !important;
}

[data-testid="stMetricValue"] {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 32px !important;
    font-weight: 700 !important;
    color: var(--charcoal) !important;
}

[data-testid="stAlert"] {
    border-radius: 10px !important;
    font-size: 14px !important;
}

hr {
    border-color: #e8e8e8 !important;
    margin: 1.4rem 0 !important;
}
</style>
""", unsafe_allow_html=True)


def sec(title):
    st.markdown(f'<p class="sec-label">{title}</p>', unsafe_allow_html=True)


def step(num, title):
    st.markdown(
        f'<div class="step-pill"><span class="step-num">{num}</span>{title}</div>',
        unsafe_allow_html=True
    )


def parse_row_input(row_text: str):
    if not row_text or not row_text.strip():
        raise ValueError("請輸入列號，例如：2,3,5-7")
    rows = set()
    for part in [p.strip() for p in row_text.split(",") if p.strip()]:
        if "-" in part:
            s, e = part.split("-", 1)
            s, e = int(s.strip()), int(e.strip())
            if s <= 0 or e <= 0:
                raise ValueError("列號必須大於 0")
            if s > e:
                raise ValueError(f"區間錯誤：{part}")
            rows.update(range(s, e + 1))
        else:
            n = int(part)
            if n <= 0:
                raise ValueError("列號必須大於 0")
            rows.add(n)
    return sorted(rows)


# ── 標題 ─────────────────────────────────────────────────
st.markdown("""
<div class="hero">
  <div class="hero-emoji">💰</div>
  <div>
    <div class="hero-title">儲值金訂單系統</div>
    <div class="hero-sub">支援建單、寄確認信、改 Google 日曆，可指定列號批次處理。</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ── 帳密 + 環境 ─────────────────────────────────────────
step("1", "登入與環境設定")

col_e, col_p, col_env = st.columns([3.2, 3.2, 1.2])
with col_e:
    backend_email = st.text_input("後台帳號")
with col_p:
    backend_password = st.text_input("後台密碼", type="password")
with col_env:
    env = st.selectbox("環境", ["prod", "dev"], index=0)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行設定 ──────────────────────────────────────────────
step("2", "執行設定")

c1, c2, c3 = st.columns(3)
with c1:
    region = st.selectbox("執行區域", ["台北", "台中", "桃園", "新竹", "高雄"])
with c2:
    sheet_name = st.text_input("工作表名稱", value="", placeholder="例：202604")
with c3:
    row_input = st.text_input("執行列號", value="", placeholder="例：2,3,5-7")

st.markdown(
    '<div class="hint-box">💡 列號支援：單列 <code>2</code>、逗號分隔 <code>2,3,5</code>、區間 <code>2,3,5-7</code></div>',
    unsafe_allow_html=True
)

st.markdown("<hr>", unsafe_allow_html=True)

# ── 執行項目 ──────────────────────────────────────────────
step("3", "執行項目")

# 根據環境決定預設執行項目
default_actions = (
    ["建單", "寄確認信", "改 Google 日曆"]
    if env == "prod"
    else ["建單"]
)

selected_actions = st.multiselect(
    "執行項目",
    options=["建單", "寄確認信", "改 Google 日曆"],
    default=default_actions,
    label_visibility="collapsed",
)

st.markdown(
    '<div class="hint-box">可自由組合，例如只寄確認信、只改日曆，或全流程一起跑。</div>',
    unsafe_allow_html=True
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
    if not backend_email.strip():
        st.error("請輸入後台帳號")
        st.stop()
    if not backend_password.strip():
        st.error("請輸入後台密碼")
        st.stop()
    if not sheet_name.strip():
        st.error("請輸入工作表名稱")
        st.stop()
    if not selected_actions:
        st.error("請至少選擇一個執行項目")
        st.stop()

    try:
        target_rows = parse_row_input(row_input)
    except Exception as e:
        st.error(f"列號格式錯誤：{e}")
        st.stop()

    logs = []

    def ui_log(msg):
        logs.append(str(msg))
        log_box.code("\\n".join(logs[-80:]))

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
                    total_success += result.get("success_count", 0)
                    total_fail += result.get("fail_count", 0)
                    total_processed += result.get("total_processed", 0)
            except Exception as e:
                total_fail += 1
                ui_log(f"❌ 第 {row_no} 列失敗：{e}")

    ui_log("===== 執行完成 =====")

    with result_container:
        st.markdown("<hr>", unsafe_allow_html=True)
        step("4", "執行結果")

        c1, c2, c3 = st.columns(3)
        c1.metric("執行筆數", total_processed)
        c2.metric("成功", total_success)
        c3.metric("失敗", total_fail)

        if total_fail == 0 and total_processed > 0:
            st.success(f"✅ 全部完成，共處理 **{total_processed}** 筆，成功 **{total_success}** 筆。")
        elif total_fail > 0:
            st.warning(f"⚠️ 執行完成，但有 **{total_fail}** 筆失敗，請查看執行過程。")
        else:
            st.info("執行完成，無資料被處理。")
