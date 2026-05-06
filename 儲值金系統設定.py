# -*- coding: utf-8 -*-
import os
import re
import json
import time
import html
from datetime import datetime, timedelta, timezone
from collections import defaultdict

import requests
import pandas as pd
from bs4 import BeautifulSoup

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from accounts import ACCOUNTS
from env import (
    ENV,
    BASE_URL_DEV,
    BASE_URL_PROD,
    GOOGLE_SHEET_ID,
    ENABLE_GCAL_COLOR_SYNC,
    GOOGLE_CALENDAR_MAP,
    GOOGLE_SERVICE_ACCOUNT_FILE,
    COLOR_PURPLE,
    COLOR_YELLOW,
    REQUEST_DELAY,
    ORDER_PREFIX_DEV,
    ORDER_PREFIX_PROD,
)

try:
    import streamlit as st
except Exception:
    st = None

try:
    from env import GOOGLE_MAPS_API_KEY
except Exception:
    GOOGLE_MAPS_API_KEY = ""


# =========================
# 執行細節 log：完整保留，但 Streamlit 畫面預設收合
# =========================
_EXECUTION_DETAIL_LOG = []
_EXECUTION_DETAIL_CONTEXT = ""


def reset_execution_detail_log():
    global _EXECUTION_DETAIL_LOG, _EXECUTION_DETAIL_CONTEXT
    _EXECUTION_DETAIL_LOG = []
    _EXECUTION_DETAIL_CONTEXT = ""


def set_execution_detail_context(row_num=None, name="", date_text=""):
    """設定後續 detail_log 所屬列別，讓執行細節可依第 X 列/客戶姓名分框顯示。"""
    global _EXECUTION_DETAIL_CONTEXT
    parts = []
    if row_num not in (None, ""):
        parts.append(f"第{row_num}列")
    if str(name or "").strip():
        parts.append(str(name).strip())
    if str(date_text or "").strip():
        parts.append(str(date_text).strip())
    _EXECUTION_DETAIL_CONTEXT = "｜".join(parts)


def _json_safe(value):
    try:
        json.dumps(value, ensure_ascii=False, default=str)
        return value
    except Exception:
        return str(value)


def detail_log(title, data=None):
    entry = {
        "context": _EXECUTION_DETAIL_CONTEXT or "未標示列別",
        "title": str(title),
        "data": _json_safe(data),
    }
    _EXECUTION_DETAIL_LOG.append(entry)
    prefix = f"[{entry['context']}] " if entry.get("context") else ""
    if data is None:
        print(f"[DETAIL] {prefix}{title}")
    else:
        try:
            print(f"[DETAIL] {prefix}{title} =", json.dumps(data, ensure_ascii=False, default=str))
        except Exception:
            print(f"[DETAIL] {prefix}{title} = {data}")


def render_execution_detail_log():
    if st is None or not _EXECUTION_DETAIL_LOG:
        return
    try:
        grouped = []
        context_index = {}
        for entry in _EXECUTION_DETAIL_LOG:
            ctx = entry.get("context") or "未標示列別"
            if ctx not in context_index:
                context_index[ctx] = len(grouped)
                grouped.append((ctx, []))
            grouped[context_index[ctx]][1].append(entry)

        for ctx, entries in grouped:
            with st.expander(f"🔍 執行細節｜{ctx}", expanded=False):
                for entry in entries:
                    st.markdown(f"**{entry['title']}**")
                    data = entry.get("data")
                    if isinstance(data, (dict, list)):
                        st.json(data)
                    elif data not in (None, ""):
                        st.code(str(data))
    except Exception:
        pass


# =========================
# 環境
# =========================
if ENV == "dev":
    BASE_URL = BASE_URL_DEV
    ORDER_PREFIX = ORDER_PREFIX_DEV
else:
    BASE_URL = BASE_URL_PROD
    ORDER_PREFIX = ORDER_PREFIX_PROD

LOGIN_URL = f"{BASE_URL}/login"
BOOKING_URL = f"{BASE_URL}/booking/stored_value_routine"
PURCHASE_URL = f"{BASE_URL}/purchase"
GET_MEMBER_URL = f"{BASE_URL}/ajax/get_member"
CHECK_CONTAIN_URL = f"{BASE_URL}/ajax/check_contain"
CALCULATE_HOUR_URL = f"{BASE_URL}/ajax/calculate_hour"
GET_SECTION_URL = f"{BASE_URL}/ajax/get_section"
MAIL_SUCCESS_URL = f"{BASE_URL}/purchase/mail_success/{{order_no}}"

HEADERS = {"User-Agent": "Mozilla/5.0"}
MAIL_HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "User-Agent": "Mozilla/5.0",
    "Referer": PURCHASE_URL,
}

CLEAN_TYPE_MAP = {
    "居家清潔": "1",
    "辦公室清潔": "2",
    "裝修細清": "3",
}

ORDER_NO_REGEX = r"(LC|TT)\d+"

# 保留舊版可穩定比對班表的系統時段
STANDARD_SLOTS = [
    "08:30-12:30",
    "09:00-11:00",
    "09:00-12:00",
    "14:00-16:00",
    "14:00-17:00",
    "14:00-18:00",
    "09:00-16:00",
    "09:00-18:00",
]

KNOWN_SERVICE_STATUS = [
    "已處理",
    "未處理",
    "處理中",
    "已完成",
    "已取消",
    "待處理",
]

print("=== 儲值金系統設定.py 版本：2026-05-06-final-balance-staff-log ===")


# =========================
# 基本工具
# =========================
def is_blank(value):
    return str(value).strip() in ("", "nan", "None")


def normalize_phone(phone_value):
    phone = str(phone_value).strip().replace(".0", "")
    phone = re.sub(r"\D", "", phone)
    if len(phone) == 9:
        phone = "0" + phone
    return phone


def normalize_text_for_parse(text):
    return re.sub(r"\s+", "", str(text or ""))


def normalize_addr_for_match(addr):
    return re.sub(r"\s+", "", str(addr or "")).strip()


def same_address(a, b):
    return normalize_addr_for_match(a) == normalize_addr_for_match(b)


def first_nonzero(*values, default="0"):
    for value in values:
        text = str(value if value is not None else "").strip()
        if text not in ("", "0", "0.0", "nan", "None"):
            return text
    return str(default)


def parse_money_value(value, default=0):
    text = str(value if value is not None else "").strip()
    if text in ("", "nan", "None", "NULL"):
        return default
    text = text.replace(",", "")
    m = re.search(r"-?\d+(?:\.\d+)?", text)
    if not m:
        return default
    try:
        return int(float(m.group(0)))
    except Exception:
        return default


def extract_nested_money(obj, keys, default=0):
    value = find_nested_value(obj, keys)
    return parse_money_value(value, default=default)


def find_nested_value(obj, keys):
    key_set = {str(k) for k in keys}

    if isinstance(obj, dict):
        for key, value in obj.items():
            if str(key) in key_set and value not in (None, ""):
                return value

        for value in obj.values():
            found = find_nested_value(value, key_set)
            if found not in (None, ""):
                return found

    elif isinstance(obj, list):
        for item in obj:
            found = find_nested_value(item, key_set)
            if found not in (None, ""):
                return found

    return ""


def parse_date_value(date_value):
    if isinstance(date_value, pd.Timestamp):
        return date_value.to_pydatetime()

    text = str(date_value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass

    raise Exception(f"無法解析日期: {date_value}")


def get_date_str(date_value):
    return parse_date_value(date_value).strftime("%Y-%m-%d")


def normalize_sheet_date(date_value):
    return get_date_str(date_value)


def is_weekend(date_value):
    return parse_date_value(date_value).weekday() >= 5


def get_unit_price_by_date(date_value):
    return 700 if is_weekend(date_value) else 600


def calc_expected_service_total(date_value, people, hours):
    """
    VIP 儲值金排程扣款規則：
    服務總額 = 人數 * 服務人時 * 單價。
    平日單價 600，週末單價 700。車馬費不列入扣款判斷。
    這個金額是含稅服務費，用來和「儲值金 + 購物金」比較。
    """
    unit_price = get_unit_price_by_date(date_value)
    try:
        total = float(people) * float(hours) * float(unit_price)
    except Exception:
        total = 0
    return int(round(total)), unit_price


def calc_pre_tax_price_from_tax_total(tax_total):
    """後台 price 欄位通常是不含稅服務費；tax_total 是含稅服務費。"""
    amount = parse_money_value(tax_total, default=0)
    if amount <= 0:
        return 0
    return int(round(amount / 1.05))


def parse_time_slot(start_time_str, end_time_str):
    if not str(start_time_str).strip() or not str(end_time_str).strip():
        raise Exception(f"開始時間或結束時間為空：{start_time_str} / {end_time_str}")

    def to_hm(t):
        text = str(t).strip()
        parts = text.split(":")
        if not parts or not parts[0].strip():
            raise Exception(f"時間格式錯誤：{t}")
        h = int(parts[0])
        m = int(parts[1]) if len(parts) > 1 and parts[1].strip() else 0
        return h, m

    sh, sm = to_hm(start_time_str)
    eh, em = to_hm(end_time_str)
    return sh, sm, eh, em


def calc_hours_from_time(start_time_str, end_time_str):
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)
    hours = (eh - sh) + (em - sm) / 60.0
    return hours if hours > 0 else None


def calc_effective_hours_from_time(start_time_str, end_time_str):
    hours = calc_hours_from_time(start_time_str, end_time_str)
    if hours is None:
        return None
    if hours >= 7:
        hours -= 1
    return hours


def normalize_period_text(start_time_str, end_time_str):
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)
    return f"{sh:02d}:{sm:02d}-{eh:02d}:{em:02d}"


def display_period_text(start_time_str, end_time_str):
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)
    return f"{sh:02d}:{sm:02d} - {eh:02d}:{em:02d}"


def normalize_sheet_period(start_time_str, end_time_str):
    return normalize_period_text(start_time_str, end_time_str)


def build_target_slot_from_row(row):
    date_part = normalize_sheet_date(row["日期"])
    period_part = normalize_sheet_period(row["開始時間"], row["結束時間"])
    return f"{date_part}_{period_part}"


def slot_duration_hours(slot_text):
    start_text, end_text = slot_text.split("-")
    return calc_effective_hours_from_time(start_text, end_text)


def slot_start_hour(slot_text):
    return int(slot_text.split("-")[0].split(":")[0])


def is_morning_slot(slot_text):
    return slot_start_hour(slot_text) < 12


def map_to_system_slot(start_time_str, end_time_str, service_text=None):
    """
    重要規則：
    1. Google Sheet 的開始/結束時間 = 客戶實際要約的服務時段，也用來查班表。
       例如 Sheet 是 09:00-12:00，就一定查 09:00-12:00。
    2. calculate_hour 回傳的 hour 只用來算價格，不用來反推班表時段。
    3. 只有特殊時段 10:00-12:00，要送系統 09:00-11:00，並在簡訊/客備註記原始時間。
    """
    original_slot = normalize_period_text(start_time_str, end_time_str)

    if original_slot == "10:00-12:00":
        return {
            "original_slot": original_slot,
            "system_slot": "09:00-11:00",
            "need_note": True,
            "sms_time": original_slot,
            "customer_time_note": f"服務時間：{original_slot}",
        }

    # 標準時段直接用 Sheet 原始時段，不用 hour 反推
    if original_slot in STANDARD_SLOTS:
        return {
            "original_slot": original_slot,
            "system_slot": original_slot,
            "need_note": False,
            "sms_time": "",
            "customer_time_note": "",
        }

    # 非標準時段才用服務時數對應系統可送時段
    actual_hours = None

    if service_text and str(service_text).strip():
        match = re.search(r"(\d+)\s*人\s*(\d+(?:\.\d+)?)\s*小時", str(service_text))
        if match:
            actual_hours = float(match.group(2))
        else:
            match = re.search(r"(\d+(?:\.\d+)?)\s*小時", str(service_text))
            if match:
                actual_hours = float(match.group(1))

    if actual_hours is None:
        actual_hours = calc_effective_hours_from_time(start_time_str, end_time_str)

    if actual_hours is None:
        raise Exception(f"無法解析服務時段: {start_time_str}-{end_time_str}")

    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)
    original_is_morning = sh < 12

    matched_slot = None
    for slot in STANDARD_SLOTS:
        if is_morning_slot(slot) == original_is_morning and abs(slot_duration_hours(slot) - actual_hours) < 1e-9:
            matched_slot = slot
            break

    if not matched_slot:
        raise Exception(f"找不到可對應的系統時段：原始時段 {original_slot}，時數 {actual_hours}")

    return {
        "original_slot": original_slot,
        "system_slot": matched_slot,
        "need_note": True,
        "sms_time": original_slot,
        "customer_time_note": f"服務時間：{original_slot}",
    }


def parse_service_human_hour(service_text, start_time, end_time):
    """
    最終規則：
    1. 預設 2 人。
    2. 預設時數 = Google Sheet 開始/結束時間換算。
    3. 若 A欄/服務人時 有明確寫「3人4小時」，則人數與時數都以 A欄為準。
    """
    people = 2
    hours = calc_effective_hours_from_time(start_time, end_time)

    if service_text and str(service_text).strip():
        text = str(service_text).strip()

        people_match = re.search(r"(\d+)\s*人", text)
        if people_match:
            people = int(people_match.group(1))

        hour_match = re.search(r"(\d+(?:\.\d+)?)\s*小時", text)
        if hour_match:
            hours = float(hour_match.group(1))

    if hours is None:
        return people, None

    return people, int(hours) if float(hours).is_integer() else hours


def normalize_hours_text(cell_value, start_time_str=None, end_time_str=None):
    people, hours = parse_service_human_hour(cell_value, start_time_str, end_time_str)
    if hours is None:
        return f"{people}人"
    htxt = f"{int(hours)}小時" if float(hours).is_integer() else f"{hours}小時"
    return f"{people}人{htxt}"


def build_group_key(row):
    normalized_human_hour = normalize_hours_text(
        row["服務人時"],
        row["開始時間"],
        row["結束時間"],
    )
    return (
        str(row["姓名"]).strip(),
        normalize_phone(row["電話"]),
        str(row["地址"]).strip(),
        str(row["購買項目"]).strip(),
        normalize_period_text(row["開始時間"], row["結束時間"]),
        normalized_human_hour,
        str(row["備註"]).strip(),
    )


def get_region_by_address(address, accounts_config):
    for region, config in accounts_config.items():
        keywords = config.get("address_keywords", [])
        if keywords:
            for kw in keywords:
                if kw in address:
                    return region
        else:
            if region == "台北" and ("台北市" in address or "新北市" in address):
                return region
            if region == "台中" and "台中市" in address:
                return region
            if region == "桃園" and "桃園" in address:
                return region
            if region == "新竹" and ("新竹市" in address or "新竹縣" in address):
                return region
            if region == "高雄" and ("高雄市" in address or "台南市" in address):
                return region
    return None


def should_process_row(row):
    return str(row.get("狀態", "")).strip() == "未安排" and is_blank(row.get("訂單編號", ""))


def should_create_order(row):
    return str(row.get("狀態", "")).strip() == "未安排" and is_blank(row.get("訂單編號", ""))


# =========================
# XYZ / 回填模板
# =========================
def finalize_xyz(meta=None, fallback_fare="0"):
    meta = meta or {}

    staff = str(meta.get("服務人員", "") or "").strip()
    status = str(meta.get("服務狀態", "") or "").strip()
    fare = str(meta.get("車馬費", "") or "").strip()

    if not staff:
        staff = "無人力"
    if not status:
        status = "未處理"
    if not fare:
        fare = str(fallback_fare or "0").strip() or "0"

    return {
        "服務人員": staff,
        "服務狀態": status,
        "車馬費": fare,
    }


def build_row_result(
    order_no="",
    result="失敗",
    reason="",
    no_slot_date="",
    insufficient_date="",
    sms_time="",
    customer_note="",
    service_notice="",
    confirm_mail="",
    calendar_result="",
    calendar_reason="",
    calendar_old="",
    calendar_new="",
    status_value="",
    staff="無人力",
    service_status="未處理",
    fare="0",
):
    xyz = finalize_xyz(
        {
            "服務人員": staff,
            "服務狀態": service_status,
            "車馬費": fare,
        },
        fallback_fare=fare or "0",
    )

    return {
        "訂單編號": order_no,
        "結果": result,
        "原因": reason,
        "沒班表日期": no_slot_date,
        "餘額不足未送": insufficient_date,
        "簡訊實際服務時間": sms_time,
        "客人備註": customer_note,
        "客服備註": service_notice,
        "確認信": confirm_mail,
        "日曆改色結果": calendar_result,
        "日曆改色原因": calendar_reason,
        "日曆原色": calendar_old,
        "日曆新色": calendar_new,
        "狀態": status_value,
        "服務人員": xyz["服務人員"],
        "服務狀態": xyz["服務狀態"],
        "車馬費": xyz["車馬費"],
    }


# =========================
# Google 憑證 / Sheet
# =========================
def get_service_account_info():
    if st is not None:
        try:
            if "gcp_service_account" in st.secrets:
                return dict(st.secrets["gcp_service_account"])
            if "GOOGLE_SERVICE_ACCOUNT" in st.secrets:
                return dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        except Exception:
            pass

    raw_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        try:
            return json.loads(raw_json)
        except Exception as e:
            raise Exception(f"GOOGLE_SERVICE_ACCOUNT_JSON 不是合法 JSON：{e}")

    candidate_files = []
    if GOOGLE_SERVICE_ACCOUNT_FILE:
        candidate_files.append(GOOGLE_SERVICE_ACCOUNT_FILE)
    candidate_files.append("google_service_account.json")

    for fp in candidate_files:
        if fp and os.path.exists(fp):
            with open(fp, "r", encoding="utf-8") as f:
                return json.load(f)

    raise FileNotFoundError(
        "找不到 Google 憑證。請在 Streamlit secrets 設定 gcp_service_account 或 GOOGLE_SERVICE_ACCOUNT，"
        "或提供 GOOGLE_SERVICE_ACCOUNT_JSON，或放置 google_service_account.json。"
    )


def build_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    service_account_info = get_service_account_info()
    creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    return gspread.authorize(creds)


def load_worksheet(sheet_name):
    client = build_gsheet_client()
    sh = client.open_by_key(GOOGLE_SHEET_ID)
    ws = sh.worksheet(sheet_name)

    values = ws.get_all_values()
    if not values:
        raise Exception(f"工作表 {sheet_name} 沒有資料")

    headers = values[0]
    rows = values[1:]
    df = pd.DataFrame(rows, columns=headers)
    df["__sheet_row__"] = range(2, len(df) + 2)
    return ws, df


def ensure_columns_in_sheet(ws):
    headers = ws.row_values(1)
    required = [
        "簡訊實際服務時間",
        "客人備註",
        "客服備註",
        "訂單編號",
        "結果",
        "原因",
        "沒班表日期",
        "餘額不足未送",
        "確認信",
        "日曆改色結果",
        "日曆改色原因",
        "日曆原色",
        "日曆新色",
        "狀態",
        "服務人員",
        "服務狀態",
        "車馬費",
    ]

    changed = False
    for col in required:
        if col not in headers:
            headers.append(col)
            changed = True

    if changed:
        ws.resize(rows=max(ws.row_count, 1), cols=len(headers))
        ws.update("A1", [headers])

    return headers


def set_customer_notice_clip_style(ws, headers=None, row_numbers=None):
    """
    Google Sheet 顯示規則：
    客服備註內容完整保留，但儲存格視覺上使用「自動裁剪 / CLIP」，
    避免長備註自動換行把列高撐高。
    """
    try:
        headers = headers or ws.row_values(1)
        if "客服備註" not in headers:
            return

        col_index = headers.index("客服備註")  # 0-based
        sheet_id = ws.id

        service_account_info = get_service_account_info()
        creds = Credentials.from_service_account_info(
            service_account_info,
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        service = build("sheets", "v4", credentials=creds)

        requests_body = [
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 1,
                        "startColumnIndex": col_index,
                        "endColumnIndex": col_index + 1,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "wrapStrategy": "CLIP"
                        }
                    },
                    "fields": "userEnteredFormat.wrapStrategy",
                }
            }
        ]

        # 只固定本次有寫入的資料列，避免長備註撐高列高。
        # row_numbers 是 Google Sheet 的 1-based row number；API 是 0-based index。
        if row_numbers:
            for row_num in sorted(set(int(x) for x in row_numbers if int(x) > 1)):
                requests_body.append(
                    {
                        "updateDimensionProperties": {
                            "range": {
                                "sheetId": sheet_id,
                                "dimension": "ROWS",
                                "startIndex": row_num - 1,
                                "endIndex": row_num,
                            },
                            "properties": {
                                "pixelSize": 21
                            },
                            "fields": "pixelSize",
                        }
                    }
                )

        service.spreadsheets().batchUpdate(
            spreadsheetId=GOOGLE_SHEET_ID,
            body={"requests": requests_body},
        ).execute()

    except Exception as e:
        print(f"設定客服備註欄位自動裁剪失敗: {e}")


def update_sheet_rows(ws, row_results):
    headers = ensure_columns_in_sheet(ws)
    header_index = {h: i + 1 for i, h in enumerate(headers)}
    updates = []

    for row_num, info in row_results.items():
        # 沒有成立訂單時，X/Y/Z/AA（服務人員、服務狀態、車馬費、客服備註）保持空白。
        if not str(info.get("訂單編號", "") or "").strip():
            info["服務人員"] = ""
            info["服務狀態"] = ""
            info["車馬費"] = ""
            info["客服備註"] = ""
        else:
            xyz = finalize_xyz(
                {
                    "服務人員": info.get("服務人員", ""),
                    "服務狀態": info.get("服務狀態", ""),
                    "車馬費": info.get("車馬費", ""),
                },
                fallback_fare=info.get("車馬費", "0"),
            )
            info["服務人員"] = xyz["服務人員"]
            info["服務狀態"] = xyz["服務狀態"]
            info["車馬費"] = xyz["車馬費"]

        for key, value in info.items():
            if key not in header_index:
                continue

            # I欄「狀態」只允許在成功完成流程時寫入「已安排」。
            # 其他空白或非已安排值都不覆蓋原本的「未安排」。
            if key == "狀態" and str(value).strip() != "已安排":
                continue

            updates.append({
                "range": gspread.utils.rowcol_to_a1(row_num, header_index[key]),
                "values": [[("" if value is None else str(value))]],
            })

    if updates:
        ws.batch_update(updates)
        set_customer_notice_clip_style(ws, headers=headers, row_numbers=row_results.keys())


# =========================
# 後台 API
# =========================
def login(session, email, password):
    resp = session.get(LOGIN_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return False

    soup = BeautifulSoup(resp.text, "html.parser")
    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        return False

    token = token_input.get("value", "").strip()
    if not token:
        return False

    resp = session.post(
        LOGIN_URL,
        data={"_token": token, "email": email, "password": password},
        headers=HEADERS,
        allow_redirects=True,
    )
    return resp.status_code == 200 and "login" not in resp.url.lower()


def get_csrf_token(session):
    resp = session.get(BOOKING_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        raise Exception(f"取得儲值金訂單頁失敗: {resp.status_code}")

    soup = BeautifulSoup(resp.text, "html.parser")
    token_input = soup.find("input", {"name": "_token"})
    if not token_input:
        raise Exception("無法從儲值金訂單頁提取 _token")

    token = token_input.get("value", "").strip()
    if not token:
        raise Exception("_token 為空")

    return token


def get_member(session, phone, token, clean_type_id):
    resp = session.post(
        GET_MEMBER_URL,
        data={"phone": phone, "_token": token, "clean_type_id": clean_type_id},
        headers=HEADERS,
        allow_redirects=True,
    )
    if resp.status_code != 200:
        return None

    try:
        result = resp.json()
    except Exception:
        return None

    return result if isinstance(result, dict) and result.get("return_code") == "0000" and result.get("member") else None


def flatten_keys(obj, prefix=""):
    keys = []
    if isinstance(obj, dict):
        for key, value in obj.items():
            path = f"{prefix}.{key}" if prefix else str(key)
            keys.append(path)
            keys.extend(flatten_keys(value, path))
    elif isinstance(obj, list):
        for idx, item in enumerate(obj):
            path = f"{prefix}[{idx}]" if prefix else f"[{idx}]"
            keys.extend(flatten_keys(item, path))
    return keys


def extract_address_notice_from_obj(obj):
    """
    客服備註跟著下拉地址。
    後台欄位名稱不固定，所以只從「選中的下拉地址原始物件」與其 purchase 內遞迴找。
    """
    preferred_keys = [
        "notice", "service_notice", "memo_notice", "customer_service_notice",
        "memoProcess", "memo_process", "customerNotice", "customer_notice",
        "notice_service", "serviceNotice",
    ]
    value = find_nested_value(obj, preferred_keys)
    if value not in (None, ""):
        return str(value)
    return ""


def pick_best_address_info(member_payload, target_address):
    """
    強制以真正下拉地址為主；沒有 addressId 視為沒選到下拉地址。
    客服備註只從選中的下拉地址原始物件 / 該地址 purchase 取，不用 lastPurchase。
    """
    member = member_payload.get("member", {}) if isinstance(member_payload, dict) else {}
    member_address_list = member.get("memberAddressList", []) if isinstance(member, dict) else []

    target_norm = normalize_addr_for_match(target_address)

    for item in member_address_list:
        item_addr = str(item.get("address", "")).strip()
        if normalize_addr_for_match(item_addr) == target_norm:
            purchase = item.get("purchase", {}) if isinstance(item.get("purchase"), dict) else {}
            raw_item = item.copy() if isinstance(item, dict) else {}
            notice = extract_address_notice_from_obj({"address": raw_item, "purchase": purchase})
            return {
                "addressId": str(item.get("id", "")).strip(),
                "country_id": item.get("countryId", ""),
                "area_id": item.get("areaId", ""),
                "address": item_addr,
                "lat": item.get("lat", ""),
                "lng": item.get("lng", ""),
                "company_id": item.get("companyId", 1),
                "purchase": purchase,
                "raw_address": raw_item,
                "notice": notice,
                "raw_address_keys": flatten_keys(raw_item)[:120],
                "address_purchase_keys": flatten_keys(purchase)[:120],
            }

    return {}


def geocode_address(address):
    """先用一般 Geocoding API；失敗時改用後台頁面相同的 Google JS GeocodeService.Search。"""
    debug = {
        "address": str(address or ""),
        "api_key_exists": bool(str(GOOGLE_MAPS_API_KEY or "").strip()),
        "normal_api": {},
        "js_api": {},
        "final_lat": "",
        "final_lng": "",
    }

    def finish(lat, lng):
        debug["final_lat"] = "" if lat is None else str(lat)
        debug["final_lng"] = "" if lng is None else str(lng)
        try:
            detail_log("🗺️ 地址定位 geocode", debug)
        except Exception:
            pass
        if lat is None or lng is None:
            return None, None
        return str(lat), str(lng)

    maps_key = str(GOOGLE_MAPS_API_KEY or "").strip()

    # 1) 標準 Geocoding API
    if maps_key:
        try:
            url = "https://maps.googleapis.com/maps/api/geocode/json"
            params = {
                "address": address,
                "language": "zh-TW",
                "key": maps_key,
            }
            resp = requests.get(url, params=params, timeout=15)
            debug["normal_api"] = {
                "url": url,
                "http_status": resp.status_code,
                "text_preview": resp.text[:500],
            }
            if resp.status_code == 200:
                data = resp.json()
                debug["normal_api"]["status"] = data.get("status")
                results = data.get("results", [])
                if results:
                    location = results[0].get("geometry", {}).get("location", {})
                    lat = location.get("lat")
                    lng = location.get("lng")
                    if lat is not None and lng is not None:
                        return finish(lat, lng)
        except Exception as e:
            debug["normal_api"]["error"] = str(e)

    # 2) 後台頁面使用的 Google JS GeocodeService.Search
    # 這支是瀏覽器 JS API 使用的 endpoint；若 env 沒設 key，就用後台頁面公開 key fallback。
    js_key = maps_key or "AIzaSyB_LWQGJzwV-rQfP0cfeGPl2PEyhk9pw5Y"
    try:
        url = "https://maps.googleapis.com/maps/api/js/GeocodeService.Search"
        params = {
            "4s": address,
            "9s": "zh-TW",
            "r_url": f"{BASE_URL}/booking/stored_value_routine",
            "callback": "_xdc_._lcgeo",
            "key": js_key,
            "token": str(int(time.time()))[-6:],
        }
        headers = {
            "accept": "*/*",
            "referer": BASE_URL + "/",
            "user-agent": HEADERS.get("User-Agent", "Mozilla/5.0"),
        }
        resp = requests.get(url, params=params, headers=headers, timeout=15)
        text = resp.text or ""
        debug["js_api"] = {
            "url": resp.url,
            "http_status": resp.status_code,
            "text_preview": text[:1000],
        }

        # 常見格式裡會出現 lat/lng 或 [lat,lng]。優先用 key-value regex。
        m = re.search(r'"lat"\s*:\s*(-?\d+(?:\.\d+)?).*?"lng"\s*:\s*(-?\d+(?:\.\d+)?)', text, flags=re.S)
        if m:
            return finish(m.group(1), m.group(2))

        # JS GeocodeService 有時用陣列存座標，找最像台灣座標的一組。
        nums = [float(x) for x in re.findall(r'-?\d+\.\d+', text)]
        for i in range(len(nums) - 1):
            lat, lng = nums[i], nums[i + 1]
            if 21 <= lat <= 26 and 119 <= lng <= 123:
                return finish(lat, lng)

    except Exception as e:
        debug["js_api"]["error"] = str(e)

    return finish(None, None)


def check_contain(session, member_id, address, lat, lng, token, clean_type_id):
    result, _debug = check_contain_with_debug(session, member_id, address, lat, lng, token, clean_type_id)
    return result


def check_contain_with_debug(session, member_id, address, lat, lng, token, clean_type_id):
    """
    模擬後台「查詢地區」。
    回傳 (json_result, debug_info)，debug_info 會完整保留 request / HTTP / raw response，方便查失敗原因。
    """
    data = {
        "memberId": str(member_id or ""),
        "cleanTypeId": str(clean_type_id or ""),
        "address": str(address or ""),
        "lat": str(lat or ""),
        "lng": str(lng or ""),
        "_token": str(token or ""),
    }
    debug = {
        "url": CHECK_CONTAIN_URL,
        "request": {
            "memberId": data["memberId"],
            "cleanTypeId": data["cleanTypeId"],
            "address": data["address"],
            "lat": data["lat"],
            "lng": data["lng"],
            "token_exists": bool(data["_token"]),
        },
        "http_status": None,
        "response_text": "",
        "json_parse_error": "",
        "json": None,
        "decision": "",
    }

    try:
        resp = session.post(
            CHECK_CONTAIN_URL,
            data=data,
            headers={**HEADERS, "X-Requested-With": "XMLHttpRequest"},
            allow_redirects=True,
            timeout=30,
        )
        debug["http_status"] = resp.status_code
        debug["final_url"] = getattr(resp, "url", "")
        debug["response_text"] = resp.text or ""
    except Exception as e:
        debug["decision"] = f"request_exception: {e}"
        return None, debug

    if resp.status_code != 200:
        debug["decision"] = f"http_status_not_200: {resp.status_code}"
        return None, debug

    try:
        result = resp.json()
        debug["json"] = result
    except Exception as e:
        debug["json_parse_error"] = str(e)
        debug["decision"] = "response_not_json"
        return None, debug

    if not isinstance(result, dict):
        debug["decision"] = "json_not_dict"
        return result, debug

    if str(result.get("return_code", "")) != "0000":
        debug["decision"] = f"return_code_not_0000: {result.get('return_code')} / {result.get('description')}"
    elif not isinstance(result.get("area"), dict) or not result.get("area"):
        debug["decision"] = "return_code_0000_but_area_empty"
    else:
        debug["decision"] = "success"

    return result, debug


def calculate_hour(session, order_data, token):
    data = order_data.copy()
    data["_token"] = token

    resp = session.post(CALCULATE_HOUR_URL, data=data, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return None

    try:
        return resp.json()
    except Exception:
        return None


def extract_calc_fields(calc_result, fallback_hours="", fallback_fare="0"):
    """
    calculate_hour 的回傳格式可能是 dict/list/html/string。
    手動流程是先送 hour/price/fare 空值，後台回傳後再填入：
    hour=4, price=4771, fare=200。
    這裡用遞迴 + 字串 regex 雙重解析。
    """
    def regex_find(text, names):
        text = str(text or "")
        for name in names:
            patterns = [
                rf'"{re.escape(name)}"\s*:\s*"?([0-9]+(?:\.[0-9]+)?)"?',
                rf"'{re.escape(name)}'\s*:\s*'?([0-9]+(?:\.[0-9]+)?)'?",
                rf'name=["\\\']{re.escape(name)}["\\\'][^>]*value=["\\\']?([0-9]+(?:\.[0-9]+)?)',
                rf'id=["\\\']{re.escape(name)}["\\\'][^>]*value=["\\\']?([0-9]+(?:\.[0-9]+)?)',
                rf'{re.escape(name)}=([0-9]+(?:\.[0-9]+)?)',
            ]
            for pat in patterns:
                m = re.search(pat, text)
                if m:
                    return m.group(1)
        return ""

    if isinstance(calc_result, (dict, list)):
        hour = find_nested_value(calc_result, [
            "hour", "clean_hour", "hours", "total_hour", "service_hour"
        ])
        price = find_nested_value(calc_result, [
            "price", "total_price", "service_price", "amount", "total", "money"
        ])
        price_vvip = find_nested_value(calc_result, [
            "price_vvip", "vvip_price", "vip_price"
        ])
        fare = find_nested_value(calc_result, [
            "fare", "car_fare", "traffic_fee", "trafficFee", "carFare", "車馬費"
        ])
        tax_total = find_nested_value(calc_result, [
            "tax_total", "taxTotal", "total_with_tax", "taxTotalPrice", "含稅總額"
        ])
        tax_price = find_nested_value(calc_result, [
            "tax_price", "taxPrice", "price_with_tax", "service_tax_price", "含稅服務費"
        ])
    else:
        hour = price = price_vvip = fare = tax_total = tax_price = ""

    raw_text = json.dumps(calc_result, ensure_ascii=False) if isinstance(calc_result, (dict, list)) else str(calc_result or "")

    if not hour:
        hour = regex_find(raw_text, ["hour", "clean_hour", "hours", "total_hour", "service_hour"])
    if not price:
        price = regex_find(raw_text, ["price", "total_price", "service_price", "amount", "total", "money"])
    if not price_vvip:
        price_vvip = regex_find(raw_text, ["price_vvip", "vvip_price", "vip_price"])
    if not fare:
        fare = regex_find(raw_text, ["fare", "car_fare", "traffic_fee", "trafficFee", "carFare"])
    if not tax_total:
        tax_total = regex_find(raw_text, ["tax_total", "taxTotal", "total_with_tax", "taxTotalPrice"])
    if not tax_price:
        tax_price = regex_find(raw_text, ["tax_price", "taxPrice", "price_with_tax", "service_tax_price"])

    required_amount = first_nonzero(tax_total, tax_price, price, default="0")

    return {
        "hour": str(hour or fallback_hours or ""),
        "price": first_nonzero(price, default="0"),
        "price_vvip": str(price_vvip or "0"),
        "fare": first_nonzero(fare, fallback_fare, default="0"),
        "tax_total": first_nonzero(tax_total, default="0"),
        "tax_price": first_nonzero(tax_price, default="0"),
        "required_amount": required_amount,
    }


def get_section_raw(session, order_data, token, date_slot):
    data = order_data.copy()
    data["_token"] = token
    data["date_list[]"] = date_slot

    resp = session.post(GET_SECTION_URL, data=data, headers=HEADERS, allow_redirects=True)
    return resp.text if resp.status_code == 200 else ""


def extract_cleaners_from_section_response(raw_text, date_slot):
    """
    從 get_section 回傳抓指定日期/時段的人員。
    支援 JSON list：
    [{"date":"2026-05-14","section":"14:00-18:00","cleaner":["胡偉勝"]}]
    """
    if not raw_text:
        return []

    date_part, period_part = date_slot.split("_", 1)
    raw = str(raw_text)

    try:
        data = json.loads(raw)
        if isinstance(data, dict):
            data = data.get("data") or data.get("result") or data.get("sections") or []
        if isinstance(data, list):
            for item in data:
                if not isinstance(item, dict):
                    continue
                item_date = str(item.get("date", "")).strip()
                item_section = str(item.get("section", "")).strip().replace(" ", "")
                if item_date == date_part and item_section == period_part.replace(" ", ""):
                    cleaners = item.get("cleaner") or item.get("cleaners") or []
                    if isinstance(cleaners, list):
                        return [str(x).strip().lstrip("＊*") for x in cleaners if str(x).strip()]
                    if isinstance(cleaners, str) and cleaners.strip():
                        return [x.strip().lstrip("＊*") for x in re.split(r"[,，、/]+", cleaners) if x.strip()]
    except Exception:
        pass

    text = html.unescape(raw)
    try:
        text = BeautifulSoup(text, "html.parser").get_text(" ", strip=True)
    except Exception:
        pass

    compact = re.sub(r"\s+", "", text)
    d = date_part
    p = period_part.replace(" ", "")
    idx = compact.find(d)
    if idx >= 0:
        nearby = compact[idx:idx + 600]
        pidx = nearby.find(p)
        if pidx >= 0:
            nearby = nearby[pidx:pidx + 500]
            m = re.search(r"[（(]([^）)]+)[）)]", nearby)
            if m:
                return [x.strip().lstrip("＊*") for x in re.split(r"[,，、/]+", m.group(1)) if x.strip()]

    return []


def clean_staff_name(name):
    """
    將後台/班表的人員文字整理成純姓名。
    例：
    - ＊黃惟芊 -> 黃惟芊
    - 00紀至聰(5) -> 紀至聰
    - 郭清松(1) -> 郭清松
    """
    text = str(name or "").strip()
    if not text:
        return ""

    text = text.replace("＊", "").replace("*", "").strip()
    text = re.sub(r"^\d+", "", text).strip()
    text = re.sub(r"[（(]\d+[）)]", "", text).strip()
    text = re.sub(r"\s+", "", text)
    return text


def format_staff_from_cleaners(cleaners, people=None):
    cleaned = []
    for name in cleaners or []:
        text = clean_staff_name(name)
        if text and text not in cleaned:
            cleaned.append(text)

    if not cleaned:
        return "無人力"

    try:
        limit = int(float(people)) if people not in (None, "") else 0
    except Exception:
        limit = 0

    if limit > 0:
        cleaned = cleaned[:limit]

    return " X ".join(cleaned) if cleaned else "無人力"


def slot_exists_in_section_response(raw_text, date_slot):
    """
    get_section 回傳可能是 HTML、JSON 包 HTML、escaped HTML。
    這裡不要只做單一 regex，改成多種格式都可比對。
    """
    if not raw_text:
        return False

    date_part, period_part = date_slot.split("_", 1)
    start_part, end_part = period_part.split("-", 1)

    raw = str(raw_text)
    unescaped = html.unescape(raw)

    try:
        soup_text = BeautifulSoup(unescaped, "html.parser").get_text(" ", strip=True)
    except Exception:
        soup_text = unescaped

    candidates = [raw, unescaped, soup_text]

    date_variants = list(dict.fromkeys([
        date_part,
        date_part.replace("-", "/"),
        date_part.replace("-", ""),
    ]))

    period_variants = list(dict.fromkeys([
        period_part,
        period_part.replace(" ", ""),
        f"{start_part} - {end_part}",
        f"{start_part}~{end_part}",
        f"{start_part}～{end_part}",
    ]))

    for text in candidates:
        compact = re.sub(r"\s+", "", text)

        for d in date_variants:
            for p in period_variants:
                dp = re.sub(r"\s+", "", d)
                pp = re.sub(r"\s+", "", p)
                if dp in compact and pp in compact:
                    date_idx = compact.find(dp)
                    period_idx = compact.find(pp)
                    if date_idx >= 0 and period_idx >= 0 and abs(period_idx - date_idx) < 500:
                        return True

        for d in date_variants:
            d_re = re.escape(d)
            s_re = re.escape(start_part)
            e_re = re.escape(end_part)
            patterns = [
                rf"{d_re}.{{0,500}}{s_re}\s*[-~～]\s*{e_re}",
                rf"{d_re}.{{0,500}}{re.escape(period_part)}",
            ]
            for pat in patterns:
                if re.search(pat, text, flags=re.S):
                    return True

    return False


# =========================
# Purchase 頁解析
# =========================
def extract_order_cards_from_purchase_html(html):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    blocks = []
    current = None

    for line in lines:
        if re.fullmatch(ORDER_NO_REGEX, line):
            if current:
                blocks.append(current)
            current = {"order_no": line, "lines": [line]}
        elif current:
            current["lines"].append(line)

    if current:
        blocks.append(current)

    return blocks


def match_order_from_purchase_page(html, target_date, target_period):
    for block in extract_order_cards_from_purchase_html(html):
        joined = "\n".join(block["lines"])
        if target_date in joined and target_period in joined:
            return block["order_no"]
    return None


def fetch_order_no_by_date_and_period(session, target_date, target_period):
    resp = session.get(PURCHASE_URL, headers=HEADERS, allow_redirects=True)
    return None if resp.status_code != 200 else match_order_from_purchase_page(resp.text, target_date, target_period)


def _extract_staff_line(lines):
    """
    從訂單列表卡片解析實際服務人員。
    只回傳純姓名，移除前綴數字與後綴括號編號。
    支援：
    - 紀至聰(5) X 郭清松(1) X 黃惟芊(2)
    - 00蔡立娟(5) X 楊超顯(5)
    """
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    # 優先抓「姓名(數字) X 姓名(數字) ...」這種實際已派工人員格式。
    candidates = re.findall(r'(?:\d{0,3})?[\u4e00-\u9fffA-Za-z]{2,12}[（(]\d+[）)]', normalized)
    names = []
    for item in candidates:
        name = clean_staff_name(item)
        if name and name not in names:
            names.append(name)

    if names:
        return " X ".join(names)

    return "無人力"


def _extract_status_line(lines):
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    for status in KNOWN_SERVICE_STATUS:
        if status in normalized:
            return status

    if "未處理" in normalized:
        return "未處理"
    if "已處理" in normalized:
        return "已處理"

    return "未處理"


def _extract_fare_line(lines):
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    m = re.search(r'車馬費[：:]?(\d+)', normalized)
    if m:
        return m.group(1)

    return "0"


def _extract_service_date_time(lines):
    service_date = ""
    service_time = ""

    for idx, line in enumerate(lines):
        text = line.strip()
        if re.match(r"\d{4}-\d{2}-\d{2}", text):
            service_date = text[:10]

            for j in range(idx + 1, min(idx + 5, len(lines))):
                nxt = lines[j].strip().replace(" ", "")
                if re.match(r"\d{2}:\d{2}-\d{2}:\d{2}", nxt):
                    service_time = nxt
                    break
            break

    return service_date, service_time


def fetch_order_meta_by_order_no(session, order_no):
    resp = session.get(PURCHASE_URL, headers=HEADERS, allow_redirects=True)
    if resp.status_code != 200:
        return {
            "服務人員": "無人力",
            "服務狀態": "未處理",
            "車馬費": "0",
            "服務日期": "",
            "服務時間": "",
        }

    blocks = extract_order_cards_from_purchase_html(resp.text)
    for block in blocks:
        if block["order_no"] == order_no:
            lines = block.get("lines", [])
            service_date, service_time = _extract_service_date_time(lines)
            staff = _extract_staff_line(lines)
            status = _extract_status_line(lines)
            fare = _extract_fare_line(lines)

            return {
                "服務人員": staff if staff else "無人力",
                "服務狀態": status if status else "未處理",
                "車馬費": fare if fare else "0",
                "服務日期": service_date,
                "服務時間": service_time,
            }

    return {
        "服務人員": "無人力",
        "服務狀態": "未處理",
        "車馬費": "0",
        "服務日期": "",
        "服務時間": "",
    }


def send_confirmation_mail(session, order_no):
    url = MAIL_SUCCESS_URL.format(order_no=order_no)
    resp = session.get(url, headers=MAIL_HEADERS, allow_redirects=True)

    if resp.status_code != 200:
        return False, f"HTTP {resp.status_code}"

    try:
        return True, str(resp.json())
    except Exception:
        return True, resp.text[:200]


# =========================
# Google Calendar
# =========================
def build_gcal_service():
    if not ENABLE_GCAL_COLOR_SYNC:
        return None

    scopes = ["https://www.googleapis.com/auth/calendar"]
    service_account_info = get_service_account_info()
    credentials = Credentials.from_service_account_info(service_account_info, scopes=scopes)
    return build("calendar", "v3", credentials=credentials)


def parse_event_time(dt_str):
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
    except Exception:
        try:
            return datetime.strptime(dt_str, "%Y-%m-%d")
        except Exception:
            return None


def color_name_from_id(color_id):
    mapping = {
        "1": "薰衣草紫",
        "2": "鼠尾草綠",
        "3": "葡萄紫",
        "4": "火鶴紅",
        "5": "香蕉黃",
        "6": "橘子橙",
        "7": "孔雀藍",
        "8": "石墨灰",
        "9": "藍莓藍",
        "10": "羅勒綠",
        "11": "番茄紅",
    }
    return mapping.get(str(color_id), f"未知({color_id})")


def find_matching_calendar_event(service, calendar_id, address, target_date, start_time_str, end_time_str):
    target_date_obj = parse_date_value(target_date)
    sh, sm, eh, em = parse_time_slot(start_time_str, end_time_str)

    tz = timezone(timedelta(hours=8))
    day_start = datetime(target_date_obj.year, target_date_obj.month, target_date_obj.day, 0, 0, 0, tzinfo=tz)
    day_end = day_start + timedelta(days=1)

    events = service.events().list(
        calendarId=calendar_id,
        timeMin=day_start.isoformat(),
        timeMax=day_end.isoformat(),
        singleEvents=True,
        orderBy="startTime",
    ).execute().get("items", [])

    target_addr = normalize_addr_for_match(address)

    for event in events:
        start_raw = event.get("start", {}).get("dateTime") or event.get("start", {}).get("date")
        end_raw = event.get("end", {}).get("dateTime") or event.get("end", {}).get("date")
        start_dt = parse_event_time(start_raw)
        end_dt = parse_event_time(end_raw)
        if not start_dt or not end_dt:
            continue

        location = event.get("location", "") or ""
        description = event.get("description", "") or ""
        summary = event.get("summary", "") or ""
        text_blob = normalize_addr_for_match(location + " " + description + " " + summary)

        if (
            start_dt.date() == target_date_obj.date()
            and (start_dt.hour, start_dt.minute) == (sh, sm)
            and (end_dt.hour, end_dt.minute) == (eh, em)
            and target_addr
            and target_addr in text_blob
        ):
            return event

    return None


def sync_calendar_color_for_row(service, calendar_id, address, date_value, start_time_str, end_time_str):
    if not ENABLE_GCAL_COLOR_SYNC or service is None:
        return {
            "日曆改色結果": "未執行",
            "日曆改色原因": "未啟用日曆改色",
            "日曆原色": "",
            "日曆新色": "",
        }

    try:
        event = find_matching_calendar_event(service, calendar_id, address, date_value, start_time_str, end_time_str)
    except HttpError as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"Calendar API 錯誤: {e}",
            "日曆原色": "",
            "日曆新色": "",
        }
    except Exception as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"Calendar 例外: {e}",
            "日曆原色": "",
            "日曆新色": "",
        }

    if not event:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": "找不到對應日曆事件",
            "日曆原色": "",
            "日曆新色": "",
        }

    event_id = event.get("id")
    old_color = str(event.get("colorId", ""))
    old_color_name = color_name_from_id(old_color)

    if old_color != COLOR_PURPLE:
        return {
            "日曆改色結果": "未改",
            "日曆改色原因": f"需求有異動（原色：{old_color_name}）",
            "日曆原色": old_color_name,
            "日曆新色": old_color_name,
        }

    try:
        service.events().patch(
            calendarId=calendar_id,
            eventId=event_id,
            body={"colorId": COLOR_YELLOW},
        ).execute()
    except HttpError as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"改色 API 錯誤: {e}",
            "日曆原色": old_color_name,
            "日曆新色": old_color_name,
        }
    except Exception as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": f"改色例外: {e}",
            "日曆原色": old_color_name,
            "日曆新色": old_color_name,
        }

    return {
        "日曆改色結果": "成功",
        "日曆改色原因": "葡萄紫 → 香蕉黃",
        "日曆原色": old_color_name,
        "日曆新色": color_name_from_id(COLOR_YELLOW),
    }


# =========================
# 各階段
# =========================
def prepare_base_order_data(row, member_payload, address_info, clean_type_id, people, hours, system_period, note_info):
    member = member_payload.get("member", {}) if isinstance(member_payload, dict) else {}
    old_purchase = address_info.get("purchase", {}) if isinstance(address_info, dict) else {}

    def pick(key, default=""):
        if old_purchase.get(key) not in (None, ""):
            return old_purchase.get(key)
        return default

    def pick_address_notice(default=""):
        # 客服備註必須以「該下拉地址系統預設帶出的備註」為準。
        # 不使用 member_payload.lastPurchase.notice，避免抓到會員其他地址或最後一筆訂單的備註。
        value = address_info.get("notice")
        if value not in (None, ""):
            return value
        value = extract_address_notice_from_obj({
            "address": address_info.get("raw_address", {}),
            "purchase": old_purchase,
        })
        if value not in (None, ""):
            return value
        return default

    base_memo = ""
    if note_info["need_note"]:
        base_memo = note_info["customer_time_note"] if not base_memo else f"{base_memo}；{note_info['customer_time_note']}"

    return {
        "clean_type_id": clean_type_id,
        "phone": normalize_phone(row["電話"]),
        "name": str(member.get("name") or row["姓名"]).strip(),
        "email": str(member.get("email") or "").strip(),
        "tel": str(member.get("tel") or normalize_phone(row["電話"])),
        "line": str(member.get("line") or ""),
        "fbName": str(member.get("fb_name") or ""),
        "fb": str(member.get("fb") or ""),
        "memoProcess": str(member.get("memo_process") or ""),
        "memoFinance": str(member.get("memo_finance") or ""),
        "addressId": str(address_info.get("addressId") or ""),
        "country_id": str(address_info.get("country_id") or pick("country_id", "12")),
        "address": str(row["地址"]).strip(),
        "ping": str(pick("ping", "4")),
        "room": str(pick("room", "0")),
        "bathroom": str(pick("bathroom", "0")),
        "balcony": str(pick("balcony", "0")),
        "livingroom": str(pick("livingroom", "0")),
        "kitchen": str(pick("kitchen", "0")),
        "window": str(pick("window", "")),
        "shutter": str(pick("shutter", "")),
        "clothes": str(pick("clothes", "0")),
        "dyson": str(pick("dyson", "0")),
        "refrigerator": str(pick("refrigerator", "0")),
        "disinfection": str(pick("disinfection", "0")),
        "go_abord": str(pick("go_abord", "0")),
        "home_move": str(pick("home_move", "0")),
        "storage": str(pick("storage", "0")),
        "cabinet": str(pick("cabinet", "0")),
        "quintuple": str(pick("quintuple", "0")),
        "hour": str(int(float(hours))),
        "price": "0",
        "price_vvip": "0",
        "person": str(int(people)),
        "date_s": "",
        "period_s": system_period,
        "period": note_info["sms_time"] if note_info["need_note"] else "",
        "cycle": "1",
        "fare": str(address_info.get("fare") or pick("fare", "0") or "0"),
        "memo": base_memo,
        "notice": str(pick_address_notice("")),
        "discount_code": "",
        "payway": "4",
        "is_backend": "477",
        "member_id": str(member.get("member_id") or ""),
        "company_id": str(address_info.get("company_id") or pick("company_id", "1")),
        "area_id": str(address_info.get("area_id") or pick("area_id", "25")),
        "lat": str(address_info.get("lat") or pick("lat", "")),
        "lng": str(address_info.get("lng") or pick("lng", "")),
    }


def filter_dates_by_balance(date_slots, date_prices, available_balance):
    # 用 Google Sheet 規則計算的含稅服務總額判斷；車馬費不列入。
    # VIP 可扣款餘額 = 儲值金 + 購物金。
    selected_slots, selected_prices, total = [], [], 0
    for slot, price in zip(date_slots, date_prices):
        price = parse_money_value(price, default=0)
        if total + price <= available_balance:
            selected_slots.append(slot)
            selected_prices.append(price)
            total += price
    return selected_slots, selected_prices, total


def fetch_vip_balance_from_page(session, member_id):
    """從儲值金歷程頁解析儲值金餘額與購物金餘額。"""
    if not member_id:
        return 0, 0, "no member_id"

    urls = [
        f"{BASE_URL}/member/{member_id}/stored_value_histories",
        f"{BASE_URL}/member/{member_id}/stored_value",
        f"{BASE_URL}/member/{member_id}",
    ]
    for url in urls:
        try:
            resp = session.get(url, headers=HEADERS, allow_redirects=True, timeout=20)
        except Exception:
            continue
        if resp.status_code != 200 or not resp.text:
            continue
        text = BeautifulSoup(resp.text, "html.parser").get_text("\n", strip=True)
        m_stored = re.search(r"儲值金餘額[：:]\s*([\d,]+)", text)
        m_shopping = re.search(r"購物金餘額[：:]\s*([\d,]+)", text)
        if m_stored or m_shopping:
            return (
                parse_money_value(m_stored.group(1) if m_stored else 0),
                parse_money_value(m_shopping.group(1) if m_shopping else 0),
                url,
            )
    return 0, 0, "not found"


def get_available_vip_balance(session, member_payload, member_id):
    # 先吃 get_member 可能已回傳的欄位；若沒有購物金，再抓儲值金歷程頁。
    stored = extract_nested_money(member_payload, ["storedValue", "stored_value", "儲值金餘額"], default=0)
    shopping = extract_nested_money(member_payload, [
        "shoppingValue", "shopping_value", "shoppingMoney", "shopping_money",
        "bonus", "bonusValue", "point", "points", "購物金餘額", "購物金",
    ], default=0)

    source = "get_member"
    page_stored, page_shopping, page_source = fetch_vip_balance_from_page(session, member_id)
    if page_stored or page_shopping:
        stored = page_stored
        shopping = page_shopping
        source = page_source

    total = stored + shopping
    detail_log("VIP balance", {
        "stored_value": stored,
        "shopping_value": shopping,
        "available_balance": total,
        "source": source,
        "note": "車馬費不列入扣款判斷；用 Google Sheet 規則：人數 * 小時 * 平日600/週末700 比對",
    })
    return total, stored, shopping


def stage_send_confirmation(order_no, session):
    if not order_no:
        return {"確認信": ""}
    try:
        ok, mail_msg = send_confirmation_mail(session, order_no)
        return {"確認信": "已發送" if ok else f"發送失敗: {mail_msg}"}
    except Exception as e:
        return {"確認信": f"發送失敗: {e}"}


def stage_calendar_color(row, gcal_service, region):
    calendar_id = GOOGLE_CALENDAR_MAP.get(region)
    if not calendar_id:
        return {
            "日曆改色結果": "未執行",
            "日曆改色原因": f"找不到區域 {region} 的日曆設定",
            "日曆原色": "",
            "日曆新色": "",
        }

    try:
        return sync_calendar_color_for_row(
            gcal_service,
            calendar_id,
            str(row["地址"]).strip(),
            row["日期"],
            str(row["開始時間"]).strip(),
            str(row["結束時間"]).strip(),
        )
    except Exception as e:
        return {
            "日曆改色結果": "失敗",
            "日曆改色原因": str(e),
            "日曆原色": "",
            "日曆新色": "",
        }


def stage_update_status(order_no, confirm_info, calendar_info, row_result=None):
    confirm_ok = str(confirm_info.get("確認信", "")).strip() == "已發送"
    calendar_ok = str(calendar_info.get("日曆改色結果", "")).strip() == "成功"

    row_result = row_result or {}
    staff_ok = str(row_result.get("服務人員", "")).strip() not in ("", "無人力")
    service_status_ok = str(row_result.get("服務狀態", "")).strip() != ""
    fare_ok = str(row_result.get("車馬費", "")).strip() != ""

    if order_no and confirm_ok and calendar_ok and staff_ok and service_status_ok and fare_ok:
        return {"狀態": "已安排"}

    return {}


def has_action(selected_actions, action_name):
    return True if not selected_actions else action_name in selected_actions


def process_existing_order_only(row, gcal_service, region, session, selected_actions=None):
    order_no = str(row.get("訂單編號", "")).strip()

    if not order_no:
        return build_row_result(
            result="失敗",
            reason="無訂單編號",
            status_value="",
            staff="無人力",
            service_status="未處理",
            fare="0",
        )

    meta = fetch_order_meta_by_order_no(session, order_no)

    result = build_row_result(
        order_no=order_no,
        result="跳過",
        reason="",
        status_value="",
        staff=meta.get("服務人員", "無人力"),
        service_status=meta.get("服務狀態", "未處理"),
        fare=meta.get("車馬費", "0"),
    )

    did_anything = False
    confirm_info = {}
    calendar_info = {}

    if has_action(selected_actions, "寄確認信"):
        confirm_info = stage_send_confirmation(order_no, session)
        result.update(confirm_info)
        did_anything = True

    if has_action(selected_actions, "改 Google 日曆"):
        calendar_info = stage_calendar_color(row, gcal_service, region)
        result.update(calendar_info)
        did_anything = True

    result.update(stage_update_status(order_no, confirm_info, calendar_info, result))

    if did_anything:
        result["結果"] = "成功"

    return result


def process_one_group(session, rows_with_idx, token, gcal_service, region, backend_user_id=None, selected_actions=None):
    first_row_num, row0 = rows_with_idx[0]
    set_execution_detail_context(first_row_num, str(row0.get("姓名", "")).strip(), get_date_str(row0["日期"]))

    purchase_item = str(row0["購買項目"]).strip()
    clean_type_id = CLEAN_TYPE_MAP.get(purchase_item)
    if not clean_type_id:
        raise Exception(f"未知購買項目: {purchase_item}")

    mapped = map_to_system_slot(row0["開始時間"], row0["結束時間"], row0["服務人時"])
    system_period = mapped["system_slot"]
    system_display_period = display_period_text(system_period.split("-")[0], system_period.split("-")[1])

    people, hours = parse_service_human_hour(row0["服務人時"], row0["開始時間"], row0["結束時間"])
    if hours is None:
        raise Exception("無法判斷服務時數")

    detail_log("👥 parsed person/hour", {
        "服務人時": str(row0["服務人時"]),
        "sheet_time": normalize_period_text(row0["開始時間"], row0["結束時間"]),
        "person": people,
        "hour": hours,
    })

    phone = normalize_phone(row0["電話"])
    member_payload = get_member(session, phone, token, clean_type_id)
    if not member_payload:
        raise Exception(f"會員不存在: {phone}")

    member = member_payload.get("member", {})
    member_id = str(member.get("member_id") or member.get("id") or "")
    available_balance, stored_value, shopping_value = get_available_vip_balance(session, member_payload, member_id)

    target_address = str(row0["地址"]).strip().split(",")[0]
    member_address_list = member.get("memberAddressList", []) if isinstance(member, dict) else []
    address_candidates = []
    for item in member_address_list[:20]:
        if isinstance(item, dict):
            address_candidates.append({
                "id": item.get("id"),
                "address": item.get("address"),
                "lat": item.get("lat"),
                "lng": item.get("lng"),
                "match": normalize_addr_for_match(item.get("address", "")) == normalize_addr_for_match(target_address),
            })

    detail_log("🏠 下拉地址比對", {
        "target_address": target_address,
        "member_id": member_id,
        "member_address_count": len(member_address_list),
        "candidates_preview": address_candidates,
        "rule": "必須從 memberAddressList 完全符合的下拉地址取 addressId 與客服備註",
    })

    best_addr = pick_best_address_info(member_payload, target_address)
    if not best_addr:
        detail_log("❌ 下拉地址失敗", {
            "target_address": target_address,
            "member_id": member_id,
            "member_address_count": len(member_address_list),
            "candidates_preview": address_candidates,
        })
        raise Exception(f"找不到對應下拉地址：{target_address}")
    if not str(best_addr.get("addressId", "")).strip():
        detail_log("❌ 下拉地址缺少 addressId", {
            "target_address": target_address,
            "selected_address": best_addr.get("address"),
            "best_addr": best_addr,
        })
        raise Exception(f"地址存在但未選到下拉地址，缺少 addressId：{target_address}")

    selected_address = str(best_addr.get("address") or target_address).strip()
    before_lat = str(best_addr.get("lat") or "").strip()
    before_lng = str(best_addr.get("lng") or "").strip()

    geo_lat, geo_lng = geocode_address(selected_address)
    if geo_lat and geo_lng:
        best_addr["lat"] = geo_lat
        best_addr["lng"] = geo_lng

    final_lat = str(best_addr.get("lat") or "").strip()
    final_lng = str(best_addr.get("lng") or "").strip()
    detail_log("📌 地址定位檢查", {
        "selected_address": selected_address,
        "addressId": best_addr.get("addressId"),
        "dropdown_lat": before_lat,
        "dropdown_lng": before_lng,
        "geocode_lat": geo_lat,
        "geocode_lng": geo_lng,
        "final_lat": final_lat,
        "final_lng": final_lng,
        "will_call_check_contain": bool(final_lat and final_lng),
    })

    if not final_lat or not final_lng:
        detail_log("❌ 地址定位失敗", {
            "selected_address": selected_address,
            "addressId": best_addr.get("addressId"),
            "dropdown_lat": before_lat,
            "dropdown_lng": before_lng,
            "geocode_lat": geo_lat,
            "geocode_lng": geo_lng,
        })
        raise Exception(f"地址定位失敗，無法查詢地區：{selected_address}")

    addr_check, addr_check_debug = check_contain_with_debug(
        session,
        member.get("member_id", ""),
        selected_address,
        best_addr.get("lat", ""),
        best_addr.get("lng", ""),
        token,
        clean_type_id,
    )
    detail_log("📍 查詢地區 request/response", addr_check_debug)
    detail_log("check_contain raw", addr_check)

    if not isinstance(addr_check, dict):
        raise Exception(
            f"查詢地區失敗：{selected_address}，lat={best_addr.get('lat')}，lng={best_addr.get('lng')}，"
            f"原因={addr_check_debug.get('decision')}，HTTP={addr_check_debug.get('http_status')}"
        )
    if str(addr_check.get("return_code", "")) != "0000":
        raise Exception(
            f"查詢地區失敗：{selected_address}，lat={best_addr.get('lat')}，lng={best_addr.get('lng')}，"
            f"return_code={addr_check.get('return_code')}，description={addr_check.get('description')}，"
            f"HTTP={addr_check_debug.get('http_status')}"
        )

    area_info = addr_check.get("area") if isinstance(addr_check.get("area"), dict) else {}
    purchase_info = addr_check.get("purchase") if isinstance(addr_check.get("purchase"), dict) else {}
    if not area_info:
        raise Exception(
            f"查詢地區成功但沒有 area：{selected_address}，lat={best_addr.get('lat')}，lng={best_addr.get('lng')}，"
            f"decision={addr_check_debug.get('decision')}，raw={addr_check}"
        )

    if area_info:
        best_addr["area_id"] = area_info.get("area_id", best_addr.get("area_id"))
        best_addr["company_id"] = area_info.get("company_id", best_addr.get("company_id"))
        best_addr["country_id"] = area_info.get("country_id", best_addr.get("country_id"))

    # 注意：purchase_info 是查詢地區回傳的付款/發票資訊，不是下拉地址 purchase。
    # 不覆蓋 best_addr["purchase"]，避免把下拉地址的客服備註與服務資料洗掉。

    # 模擬後台「查詢地區」後的資料補齊：
    # 車馬費可能在 purchase、area 或巢狀欄位中，需全部掃描。
    fare_from_check = first_nonzero(
        purchase_info.get("fare") if purchase_info else "",
        purchase_info.get("car_fare") if purchase_info else "",
        purchase_info.get("traffic_fee") if purchase_info else "",
        area_info.get("fare") if area_info else "",
        area_info.get("car_fare") if area_info else "",
        area_info.get("traffic_fee") if area_info else "",
        find_nested_value(addr_check, ["fare", "car_fare", "traffic_fee", "trafficFee", "車馬費"]),
        best_addr.get("fare", ""),
        default="0",
    )
    best_addr["fare"] = fare_from_check

    # 客服備註來源修正：
    # 後台在選定會員地址 / 查詢地區後，應帶出「該地址前一次訂單」的預設備註。
    # 這裡只接受 check_contain 的 purchase / 該地址 address_info 回傳值，
    # 不使用 area_info.notice，也不使用 member_payload.lastPurchase.notice，
    # 避免抓到區域備註、會員其他地址或最後一筆訂單的備註。
    notice_from_address = (
        best_addr.get("notice", "")
        or extract_address_notice_from_obj({
            "address": best_addr.get("raw_address", {}),
            "purchase": best_addr.get("purchase", {}),
        })
        or ""
    )
    best_addr["notice"] = notice_from_address
    detail_log("📝 selected address notice", {
        "addressId": best_addr.get("addressId"),
        "notice_len": len(str(notice_from_address or "")),
        "raw_address_keys": best_addr.get("raw_address_keys", []),
        "address_purchase_keys": best_addr.get("address_purchase_keys", []),
        "check_contain_purchase_keys": flatten_keys(purchase_info)[:80],
    })

    base_data = prepare_base_order_data(
        row0,
        member_payload,
        best_addr,
        clean_type_id,
        people,
        hours,
        system_period,
        mapped,
    )

    # 強制套用查詢地址後取得的區域/車馬費資料
    base_data["fare"] = first_nonzero(best_addr.get("fare"), base_data.get("fare"), default="0")
    base_data["notice"] = str(best_addr.get("notice") or base_data.get("notice") or "")
    base_data["area_id"] = str(best_addr.get("area_id") or base_data.get("area_id") or "")
    base_data["company_id"] = str(best_addr.get("company_id") or base_data.get("company_id") or "")
    base_data["country_id"] = str(best_addr.get("country_id") or base_data.get("country_id") or "")
    base_data["addressId"] = str(best_addr.get("addressId") or base_data.get("addressId") or "")
    base_data["lat"] = str(best_addr.get("lat") or base_data.get("lat") or "")
    base_data["lng"] = str(best_addr.get("lng") or base_data.get("lng") or "")

    print("[DEBUG] address check result =", {
        "addressId": base_data.get("addressId"),
        "area_id": base_data.get("area_id"),
        "company_id": base_data.get("company_id"),
        "fare": base_data.get("fare"),
        "lat": base_data.get("lat"),
        "lng": base_data.get("lng"),
    })

    def build_time_fields():
        sms_time = base_data.get("period", "")
        customer_note = base_data.get("memo", "")
        if mapped["need_note"]:
            sms_time = mapped["original_slot"]
            customer_note = f"服務時間：{mapped['original_slot']}"
        return sms_time, customer_note

    def build_priced_payload_for_date(date_s):
        calc_data = base_data.copy()

        # 重要：完全模擬手動「計算時數」流程。
        # 手動 request 會送 date_s/hour/price/price_vvip/fare 空值，
        # 讓後台自行計算 hour/price/fare；若先帶 0，後台可能不會重算。
        # 查詢班表/計算時數前，先把人數與時數改成 Google Sheet/A欄規則後的值。
        # 不採用後台自動推回來的 hour 來決定班表。
        calc_data["date_s"] = date_s
        calc_data["hour"] = str(base_data.get("hour") or "")
        calc_data["person"] = str(base_data.get("person") or "")
        calc_data["price"] = ""
        calc_data["price_vvip"] = ""
        calc_data["fare"] = ""

        calc_result = calculate_hour(session, calc_data, token)
        if not calc_result:
            raise Exception(f"計算時數失敗：{date_s}")

        detail_log("🟠 calculate_hour raw", calc_result)

        calc_fields = extract_calc_fields(
            calc_result,
            fallback_hours=base_data.get("hour", ""),
            fallback_fare=best_addr.get("fare", "0"),
        )

        expected_total, unit_price = calc_expected_service_total(date_s, people, hours)
        expected_price = calc_pre_tax_price_from_tax_total(expected_total)

        calc_fields["sheet_expected_total"] = str(expected_total)
        calc_fields["sheet_unit_price"] = str(unit_price)
        calc_fields["sheet_expected_price"] = str(expected_price)
        calc_fields["balance_rule"] = "儲值金 + 購物金 >= Google Sheet 規則含稅服務總額；車馬費不列入"
        calc_fields["calculate_hour_tax_total_note"] = (
            "僅供 debug。若後台 calculate_hour 回傳與 Sheet 規則不同，"
            "扣款判斷與送單 price 以 Sheet 規則為準。"
        )

        payload = base_data.copy()
        payload["date_s"] = date_s
        payload["hour"] = str(base_data.get("hour") or calc_fields.get("hour") or "")
        payload["person"] = str(base_data.get("person") or payload.get("person") or "")
        payload["price"] = str(expected_price or calc_fields.get("price") or "0")
        payload["price_vvip"] = str(calc_fields.get("price_vvip") or "0")

        detail_log("🟣 calc_fields", calc_fields)
        payload["fare"] = first_nonzero(calc_fields.get("fare"), best_addr.get("fare"), base_data.get("fare"), default="0")
        payload["_required_balance_amount"] = str(expected_total)
        payload["_backend_required_balance_amount"] = str(calc_fields.get("required_amount") or calc_fields.get("tax_total") or calc_fields.get("tax_price") or calc_fields.get("price") or "0")

        if str(payload.get("price", "")).strip() in ("", "0", "0.0"):
            raise Exception(f"計算服務費後 price 仍為 0，請貼 🟠 calculate_hour raw 與 🟣 calc_fields：{date_s}")
        if str(payload.get("_required_balance_amount", "")).strip() in ("", "0", "0.0"):
            raise Exception(f"Sheet 規則服務總額仍為 0，請貼 👥 parsed person/hour 與 🟣 calc_fields：{date_s}")

        payload["notice"] = str(base_data.get("notice") or best_addr.get("notice") or "")
        payload["area_id"] = str(base_data.get("area_id") or best_addr.get("area_id") or "")
        payload["company_id"] = str(base_data.get("company_id") or best_addr.get("company_id") or "")
        payload["addressId"] = str(base_data.get("addressId") or best_addr.get("addressId") or "")
        return payload

    row_details = []
    for row_num, row in rows_with_idx:
        date_s = get_date_str(row["日期"])
        set_execution_detail_context(row_num, str(row.get("姓名", "")).strip(), date_s)
        priced_payload = build_priced_payload_for_date(date_s)
        required_amount = parse_money_value(priced_payload.get("_required_balance_amount"), default=0)

        row_details.append({
            "row_num": row_num,
            "date": date_s,
            "slot": f"{date_s}_{system_period}",
            "price": required_amount,  # 用 Sheet 規則含稅服務總額比對儲值金+購物金；車馬費不列入
            "display_period": system_display_period,
            "row": row,
            "payload": priced_payload,
        })

        detail_log("🧭 row slot", {
            "row_num": row_num,
            "sheet_time": normalize_period_text(row["開始時間"], row["結束時間"]),
            "system_period": system_period,
            "slot": f"{date_s}_{system_period}",
            "price": priced_payload.get("price"),
            "sheet_required_amount": priced_payload.get("_required_balance_amount"),
            "backend_calculate_hour_required_amount": priced_payload.get("_backend_required_balance_amount"),
            "fare": priced_payload.get("fare"),
            "balance_rule": "儲值金 + 購物金 >= sheet_required_amount；車馬費不列入",
        })

    need_create_order = has_action(selected_actions, "建單")
    row_results = {}

    if not need_create_order:
        for detail in row_details:
            existing_order_no = str(detail["row"].get("訂單編號", "")).strip()
            sms_time, customer_note = build_time_fields()
            service_notice = str(detail["payload"].get("notice") or "")

            meta = fetch_order_meta_by_order_no(session, existing_order_no) if existing_order_no else {
                "服務人員": "無人力",
                "服務狀態": "未處理",
                "車馬費": "0",
            }

            result = build_row_result(
                order_no=existing_order_no,
                result="成功" if existing_order_no else "失敗",
                reason="" if existing_order_no else "無訂單編號，無法寄信或改日曆",
                sms_time=sms_time,
                customer_note=customer_note,
                service_notice=service_notice,
                status_value="",
                staff=meta.get("服務人員", "無人力"),
                service_status=meta.get("服務狀態", "未處理"),
                fare=meta.get("車馬費", "0"),
            )

            if existing_order_no and has_action(selected_actions, "寄確認信"):
                result.update(stage_send_confirmation(existing_order_no, session))

            if has_action(selected_actions, "改 Google 日曆"):
                calendar_info = stage_calendar_color(detail["row"], gcal_service, region)
                result.update(calendar_info)
                if existing_order_no:
                    result.update(stage_update_status(existing_order_no, result, calendar_info, result))

            row_results[detail["row_num"]] = result

        return row_results

    no_slot_dates = []
    valid_details = []

    for detail in row_details:
        set_execution_detail_context(detail["row_num"], str(detail["row"].get("姓名", "")).strip(), detail["date"])
        raw = get_section_raw(session, detail["payload"], token, detail["slot"])
        slot_ok = slot_exists_in_section_response(raw, detail["slot"])
        cleaners = extract_cleaners_from_section_response(raw, detail["slot"])
        detail["section_cleaners"] = cleaners
        detail["section_staff"] = format_staff_from_cleaners(cleaners, people=people)

        detail_log("🧩 section match", {
            "slot": detail["slot"],
            "matched": slot_ok,
            "staff": detail.get("section_staff"),
            "raw_preview": str(raw)[:500],
        })

        if slot_ok:
            valid_details.append(detail)
        else:
            no_slot_dates.append(detail["date"])

    if not valid_details:
        for detail in row_details:
            sms_time, customer_note = build_time_fields()
            service_notice = str(detail["payload"].get("notice") or "")
            row_results[detail["row_num"]] = build_row_result(
                result="失敗",
                reason="無班表",
                no_slot_date=detail["date"],
                sms_time=sms_time,
                customer_note=customer_note,
                service_notice=service_notice,
                status_value="",
                staff="無人力",
                service_status="未處理",
                fare="0",
            )
        return row_results

    valid_slots_for_balance = [x["slot"] for x in valid_details]
    valid_prices_for_balance = [x["price"] for x in valid_details]
    send_slots, _, _ = filter_dates_by_balance(valid_slots_for_balance, valid_prices_for_balance, available_balance)

    insufficient_dates = []
    send_details = []

    for detail in valid_details:
        if detail["slot"] in send_slots:
            send_details.append(detail)
        else:
            insufficient_dates.append(detail["date"])

    for detail in row_details:
        set_execution_detail_context(detail["row_num"], str(detail["row"].get("姓名", "")).strip(), detail["date"])
        sms_time, customer_note = build_time_fields()
        service_notice = str(detail["payload"].get("notice") or "")

        if detail["date"] in no_slot_dates:
            row_results[detail["row_num"]] = build_row_result(
                result="失敗",
                reason="無班表",
                no_slot_date=detail["date"],
                sms_time=sms_time,
                customer_note=customer_note,
                service_notice=service_notice,
                status_value="",
                staff="無人力",
                service_status="未處理",
                fare="0",
            )
        elif detail["date"] in insufficient_dates:
            row_results[detail["row_num"]] = build_row_result(
                result="未送",
                reason="餘額不足",
                insufficient_date=detail["date"],
                sms_time=sms_time,
                customer_note=customer_note,
                service_notice=service_notice,
                status_value="",
                staff=detail.get("section_staff") or "無人力",
                service_status="未處理",
                fare=str(detail["payload"].get("fare") or "0"),
            )

    if not send_details:
        return row_results

    # 每筆單獨送出，避免日期互相污染
    for detail in send_details:
        set_execution_detail_context(detail["row_num"], str(detail["row"].get("姓名", "")).strip(), detail["date"])
        payload = detail["payload"].copy()
        slots = [detail["slot"]]

        detail_log("📦 booking payload",
              {
                  "date": detail["date"],
                  "slot": detail["slot"],
                  "price": payload.get("price"),
                  "sheet_required_amount": payload.get("_required_balance_amount"),
                  "backend_calculate_hour_required_amount": payload.get("_backend_required_balance_amount"),
                  "fare": payload.get("fare"),
                  "addressId": payload.get("addressId"),
                  "area_id": payload.get("area_id"),
                  "company_id": payload.get("company_id"),
                  "notice_len": len(str(payload.get("notice") or "")),
              })

        session.post(
            BOOKING_URL,
            data={**payload, "_token": token, "date_list[]": slots},
            headers=HEADERS,
            allow_redirects=True,
        )

        time.sleep(1)

        order_no = fetch_order_no_by_date_and_period(session, detail["date"], detail["display_period"])
        sms_time, customer_note = build_time_fields()
        service_notice = str(payload.get("notice") or "")

        if not order_no:
            row_results[detail["row_num"]] = build_row_result(
                result="已送出",
                reason="抓不到訂單編號",
                sms_time=sms_time,
                customer_note=customer_note,
                service_notice=service_notice,
                status_value="",
                staff=detail.get("section_staff") or "無人力",
                service_status="未處理",
                fare=str(detail["payload"].get("fare") or "0"),
            )
            continue

        meta = fetch_order_meta_by_order_no(session, order_no)

        staff_value = meta.get("服務人員", "")
        if not staff_value or staff_value == "無人力":
            staff_value = detail.get("section_staff") or "無人力"

        stage_result = build_row_result(
            order_no=order_no,
            result="成功",
            reason="",
            sms_time=sms_time,
            customer_note=customer_note,
            service_notice=service_notice,
            status_value="",
            staff=staff_value,
            service_status=meta.get("服務狀態", "未處理"),
            fare=meta.get("車馬費", "0") or str(detail["payload"].get("fare") or "0"),
        )

        confirm_info = {}
        calendar_info = {}

        if has_action(selected_actions, "寄確認信"):
            confirm_info = stage_send_confirmation(order_no, session)
            stage_result.update(confirm_info)

        if has_action(selected_actions, "改 Google 日曆"):
            calendar_info = stage_calendar_color(detail["row"], gcal_service, region)
            stage_result.update(calendar_info)

        stage_result.update(stage_update_status(order_no, confirm_info, calendar_info, stage_result))

        row_results[detail["row_num"]] = stage_result

    return row_results


# =========================
# 主執行
# =========================
def run_process(sheet_name, start_row, end_row, env_name_from_ui=None):
    print(f"目前環境：{ENV}")
    print(f"BASE_URL：{BASE_URL}")
    print(f"執行工作表：{sheet_name}")
    print(f"執行列範圍：{start_row} ~ {end_row}")

    ws, df = load_worksheet(sheet_name)

    required_cols = [
        "服務人時",
        "備註",
        "姓名",
        "電話",
        "地址",
        "日期",
        "開始時間",
        "結束時間",
        "狀態",
        "購買項目",
        "訂單編號",
    ]
    for col in required_cols:
        if col not in df.columns:
            raise Exception(f"工作表缺少必要欄位: {col}")

    df = df[(df["__sheet_row__"] >= start_row) & (df["__sheet_row__"] <= end_row)]
    df = df[df.apply(should_process_row, axis=1)]

    if df.empty:
        print("沒有符合條件的資料可執行。")
        return

    gcal_service = None
    if ENABLE_GCAL_COLOR_SYNC:
        try:
            gcal_service = build_gcal_service()
            print("Google Calendar 已啟用")
        except Exception as e:
            print(f"Google Calendar 初始化失敗：{e}")
            gcal_service = None

    grouped_orders = defaultdict(list)

    for _, row in df.iterrows():
        region = get_region_by_address(str(row["地址"]), ACCOUNTS)
        if not region:
            continue
        if not should_create_order(row):
            continue

        key = (region, build_group_key(row))
        grouped_orders[key].append((int(row["__sheet_row__"]), row))

    all_row_results = {}

    region_groups = defaultdict(list)
    for (region, group_key), items in grouped_orders.items():
        region_groups[region].append((group_key, items))

    for region, group_items in region_groups.items():
        config = ACCOUNTS.get(region)
        if not config:
            continue

        email = config["email"]
        password = config["password"]

        print(f"\n===== 開始處理區域：{region} ({email}) =====")

        session = requests.Session()
        if not login(session, email, password):
            print("登入失敗，略過該區域")
            continue

        for group_no, (_, rows_with_idx) in enumerate(group_items, start=1):
            _, first_row = rows_with_idx[0]
            print(f"\n--- 處理第 {group_no} 組：{first_row['姓名']}，共 {len(rows_with_idx)} 筆 ---")

            try:
                token = get_csrf_token(session)
                row_results = process_one_group(
                    session,
                    rows_with_idx,
                    token,
                    gcal_service,
                    region,
                    None,
                    ["建單", "寄確認信", "改 Google 日曆"],
                )
                all_row_results.update(row_results)
            except Exception as e:
                print(f"❌ 整組失敗：{e}")
                for row_num, _ in rows_with_idx:
                    all_row_results[row_num] = build_row_result(
                        result="失敗",
                        reason=str(e),
                        status_value="",
                        staff="無人力",
                        service_status="未處理",
                        fare="0",
                    )

            time.sleep(REQUEST_DELAY)

    update_sheet_rows(ws, all_row_results)
    print("已回填 Google Sheet。")


def get_runtime_config(env_name: str):
    if env_name == "dev":
        return {
            "BASE_URL": BASE_URL_DEV,
            "ORDER_PREFIX": ORDER_PREFIX_DEV,
        }
    return {
        "BASE_URL": BASE_URL_PROD,
        "ORDER_PREFIX": ORDER_PREFIX_PROD,
    }


def run_process_web(env_name, region, backend_email, backend_password, sheet_name, start_row, end_row, selected_actions=None, logger=print):
    global BASE_URL, ORDER_PREFIX
    if env_name == "dev":
        BASE_URL = BASE_URL_DEV
        ORDER_PREFIX = ORDER_PREFIX_DEV
    else:
        BASE_URL = BASE_URL_PROD
        ORDER_PREFIX = ORDER_PREFIX_PROD

    global LOGIN_URL, BOOKING_URL, PURCHASE_URL, GET_MEMBER_URL
    global CHECK_CONTAIN_URL, CALCULATE_HOUR_URL, GET_SECTION_URL, MAIL_SUCCESS_URL

    LOGIN_URL = f"{BASE_URL}/login"
    BOOKING_URL = f"{BASE_URL}/booking/stored_value_routine"
    PURCHASE_URL = f"{BASE_URL}/purchase"
    GET_MEMBER_URL = f"{BASE_URL}/ajax/get_member"
    CHECK_CONTAIN_URL = f"{BASE_URL}/ajax/check_contain"
    CALCULATE_HOUR_URL = f"{BASE_URL}/ajax/calculate_hour"
    GET_SECTION_URL = f"{BASE_URL}/ajax/get_section"
    MAIL_SUCCESS_URL = f"{BASE_URL}/purchase/mail_success/{{order_no}}"

    reset_execution_detail_log()

    logger(f"目前環境：{env_name}")
    logger(f"BASE_URL：{BASE_URL}")
    logger(f"執行區域：{region}")
    logger(f"執行工作表：{sheet_name}")
    logger(f"執行列範圍：{start_row} ~ {end_row}")

    if selected_actions is None:
        selected_actions = ["建單", "寄確認信", "改 Google 日曆"]

    ws, df = load_worksheet(sheet_name)

    required_cols = [
        "服務人時",
        "備註",
        "姓名",
        "電話",
        "地址",
        "日期",
        "開始時間",
        "結束時間",
        "狀態",
        "購買項目",
        "訂單編號",
    ]
    for col in required_cols:
        if col not in df.columns:
            raise Exception(f"工作表缺少必要欄位: {col}")

    df = df[(df["__sheet_row__"] >= start_row) & (df["__sheet_row__"] <= end_row)]
    df = df[df.apply(should_process_row, axis=1)]

    if df.empty:
        logger("沒有符合條件的資料可執行。")
        return {
            "success": True,
            "message": "沒有符合條件的資料",
            "failed_records": [],
        }

    filtered_rows = [row for _, row in df.iterrows() if get_region_by_address(str(row["地址"]), ACCOUNTS) == region]
    if not filtered_rows:
        logger(f"沒有 {region} 區域的資料可執行。")
        return {
            "success": True,
            "message": f"沒有 {region} 區域資料",
            "failed_records": [],
        }

    df = pd.DataFrame(filtered_rows)
    if "__sheet_row__" not in df.columns:
        raise Exception("資料缺少 __sheet_row__")

    gcal_service = None
    if ENABLE_GCAL_COLOR_SYNC:
        try:
            gcal_service = build_gcal_service()
            logger("Google Calendar 已啟用")
        except Exception as e:
            logger(f"Google Calendar 初始化失敗：{e}")
            gcal_service = None

    session = requests.Session()
    if not login(session, backend_email, backend_password):
        raise Exception("後台登入失敗，請確認帳號密碼")

    grouped_orders = defaultdict(list)
    existing_order_rows = []

    for _, row in df.iterrows():
        row_num = int(row["__sheet_row__"])

        if not has_action(selected_actions, "建單") or not should_create_order(row):
            existing_order_rows.append((row_num, row))
            continue

        grouped_orders[build_group_key(row)].append((row_num, row))

    all_row_results = {}
    failed_records = []

    total_groups = len(grouped_orders) + len(existing_order_rows)
    logger(f"本次共 {total_groups} 組/筆待處理")

    for row_num, row in existing_order_rows:
        try:
            result = process_existing_order_only(row, gcal_service, region, session, selected_actions)
            all_row_results[row_num] = result
            if result.get("結果") == "失敗":
                failed_records.append({
                    "row": row_num,
                    "name": str(row.get("姓名", "未知")).strip(),
                    "error": str(result.get("原因", "")),
                })
        except Exception as e:
            all_row_results[row_num] = build_row_result(
                result="失敗",
                reason=f"補處理失敗: {e}",
                status_value="",
                staff="無人力",
                service_status="未處理",
                fare="0",
            )
            failed_records.append({
                "row": row_num,
                "name": str(row.get("姓名", "未知")).strip(),
                "error": f"補處理失敗: {e}",
            })

    for group_no, (_, rows_with_idx) in enumerate(grouped_orders.items(), start=1):
        _, first_row = rows_with_idx[0]
        date_preview = ", ".join(get_date_str(r["日期"]) for _, r in rows_with_idx[:3])
        more = "..." if len(rows_with_idx) > 3 else ""
        logger(f"處理第 {group_no}/{len(grouped_orders)} 組：{first_row['姓名']}，共 {len(rows_with_idx)} 筆（{date_preview}{more}）")

        try:
            token = get_csrf_token(session)
            row_results = process_one_group(session, rows_with_idx, token, gcal_service, region, None, selected_actions)
            all_row_results.update(row_results)

            for row_num, row in rows_with_idx:
                result = row_results.get(row_num, {})
                if result.get("結果") == "失敗":
                    failed_records.append({
                        "row": row_num,
                        "name": str(row.get("姓名", "未知")).strip(),
                        "error": str(result.get("原因", "")),
                    })
        except Exception as e:
            logger(f"整組失敗：{e}")
            for row_num, row in rows_with_idx:
                failed_records.append({
                    "row": row_num,
                    "name": str(row.get("姓名", "未知")).strip(),
                    "error": str(e),
                })
                all_row_results[row_num] = build_row_result(
                    result="失敗",
                    reason=str(e),
                    status_value="",
                    staff="無人力",
                    service_status="未處理",
                    fare="0",
                )

        time.sleep(REQUEST_DELAY)

    update_sheet_rows(ws, all_row_results)
    logger("已回填 Google Sheet。")

    success_count = sum(1 for v in all_row_results.values() if v.get("結果") == "成功")
    fail_count = sum(1 for v in all_row_results.values() if v.get("結果") == "失敗")

    logger(f"執行摘要：成功 {success_count} 筆｜失敗 {fail_count} 筆｜總處理 {len(all_row_results)} 筆")
    render_execution_detail_log()

    return {
        "success": True,
        "sheet_name": sheet_name,
        "region": region,
        "env": env_name,
        "success_count": success_count,
        "fail_count": fail_count,
        "total_processed": len(all_row_results),
        "failed_records": failed_records,
    }
