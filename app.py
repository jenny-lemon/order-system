import streamlit as st
from 儲值金系統設定 import run_process_web

st.set_page_config(page_title="儲值金系統", layout="wide")
st.title("儲值金訂單安全版網頁系統")

with st.form("run_form"):
    col1, col2 = st.columns(2)

    with col1:
        env = st.selectbox("執行環境", ["dev", "prod"], index=0)
        region = st.selectbox("執行區域", ["台北", "台中"])
        backend_email = st.text_input("後台帳號")
        backend_password = st.text_input("後台密碼", type="password")

    with col2:
        sheet_name = st.text_input("工作表名稱", value="202604")
        start_row = st.number_input("開始列", min_value=1, value=2, step=1)
        end_row = st.number_input("結束列", min_value=1, value=10, step=1)

    submitted = st.form_submit_button("開始執行")

if submitted:
    if not backend_email.strip():
        st.error("請輸入後台帳號")
        st.stop()

    if not backend_password.strip():
        st.error("請輸入後台密碼")
        st.stop()

    log_box = st.empty()
    logs = []

    def ui_log(msg: str):
        logs.append(msg)
        log_box.text("\n".join(logs[-200:]))

    try:
        result = run_process_web(
            env_name=env,
            region=region,
            backend_email=backend_email.strip(),
            backend_password=backend_password.strip(),
            sheet_name=sheet_name.strip(),
            start_row=int(start_row),
            end_row=int(end_row),
            logger=ui_log,
        )
        st.success("執行完成")
        st.json(result)
    except Exception as e:
        st.error(f"執行失敗：{e}")
