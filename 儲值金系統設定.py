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

print("=== 儲值金系統設定.py 版本：2026-05-03-final-staff-notice-aa ===")


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

    staff_raw = str(meta.get("服務人員", "") or "").strip()
    staff = normalize_staff_display(staff_raw) if staff_raw else ""
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


def pick_best_address_info(member_payload, target_address):
    """
    強制以真正下拉地址為主；沒有 addressId 視為沒選到下拉地址
    """
    member = member_payload.get("member", {}) if isinstance(member_payload, dict) else {}
    member_address_list = member.get("memberAddressList", []) if isinstance(member, dict) else []

    target_norm = normalize_addr_for_match(target_address)

    for item in member_address_list:
        item_addr = str(item.get("address", "")).strip()
        if normalize_addr_for_match(item_addr) == target_norm:
            return {
                "addressId": str(item.get("id", "")).strip(),
                "country_id": item.get("countryId", ""),
                "area_id": item.get("areaId", ""),
                "address": item_addr,
                "lat": item.get("lat", ""),
                "lng": item.get("lng", ""),
                "company_id": item.get("companyId", 1),
                "purchase": item.get("purchase", {}) if isinstance(item.get("purchase"), dict) else {},
            }

    return {}


def geocode_address(address):
    if not GOOGLE_MAPS_API_KEY:
        return None, None

    try:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            "address": address,
            "language": "zh-TW",
            "key": GOOGLE_MAPS_API_KEY,
        }
        resp = requests.get(url, params=params, timeout=15)
        if resp.status_code != 200:
            return None, None

        data = resp.json()
        results = data.get("results", [])
        if not results:
            return None, None

        location = results[0].get("geometry", {}).get("location", {})
        lat = location.get("lat")
        lng = location.get("lng")
        if lat is None or lng is None:
            return None, None

        return str(lat), str(lng)
    except Exception:
        return None, None


def check_contain(session, member_id, address, lat, lng, token, clean_type_id):
    resp = session.post(
        CHECK_CONTAIN_URL,
        data={
            "memberId": member_id,
            "cleanTypeId": clean_type_id,
            "address": address,
            "lat": lat or "",
            "lng": lng or "",
            "_token": token,
        },
        headers=HEADERS,
        allow_redirects=True,
    )
    if resp.status_code != 200:
        return None

    try:
        return resp.json()
    except Exception:
        return None


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
    else:
        hour = price = price_vvip = fare = ""

    raw_text = json.dumps(calc_result, ensure_ascii=False) if isinstance(calc_result, (dict, list)) else str(calc_result or "")

    if not hour:
        hour = regex_find(raw_text, ["hour", "clean_hour", "hours", "total_hour", "service_hour"])
    if not price:
        price = regex_find(raw_text, ["price", "total_price", "service_price", "amount", "total", "money"])
    if not price_vvip:
        price_vvip = regex_find(raw_text, ["price_vvip", "vvip_price", "vip_price"])
    if not fare:
        fare = regex_find(raw_text, ["fare", "car_fare", "traffic_fee", "trafficFee", "carFare"])

    return {
        "hour": str(hour or fallback_hours or ""),
        "price": first_nonzero(price, default="0"),
        "price_vvip": str(price_vvip or "0"),
        "fare": first_nonzero(fare, fallback_fare, default="0"),
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
    將班表/訂單頁的人名清成純姓名。
    例：
    - 00洪暐智(4) -> 洪暐智
    - X蔡佩玲(1) -> 蔡佩玲
    - ＊黃惟芊 -> 黃惟芊
    - 吳豐閔 X X蔡佩玲 -> 會在 normalize_staff_display 再統一成 吳豐閔 X 蔡佩玲
    """
    text = html.unescape(str(name or "")).strip()
    if not text:
        return ""

    text = text.strip().lstrip("＊*").strip()
    text = re.sub(r"^[Xx×＊*\s]+", "", text).strip()
    text = re.sub(r"^\d+", "", text).strip()
    text = re.sub(r"[（(]\d+[）)]", "", text).strip()
    text = re.sub(r"^[Xx×\s]+", "", text).strip()
    text = re.sub(r"[Xx×\s]+$", "", text).strip()
    return text


def normalize_staff_display(value, limit=None):
    """
    X欄顯示規則：名字和名字中間只保留一個「 X 」。
    不管來源是 list、已經串好的字串、或含有 X姓名，都先拆開、清洗、再重組。
    """
    if value in (None, ""):
        return ""

    if isinstance(value, (list, tuple)):
        raw_parts = []
        for item in value:
            raw_parts.extend(re.split(r"\s*[Xx×]\s*|[,，、/]+", str(item or "")))
    else:
        raw_parts = re.split(r"\s*[Xx×]\s*|[,，、/]+", str(value or ""))

    cleaned = []
    seen = set()
    for part in raw_parts:
        name = clean_staff_name(part)
        if not name or name in seen:
            continue
        cleaned.append(name)
        seen.add(name)
        if limit and len(cleaned) >= int(limit):
            break

    return " X ".join(cleaned)


def format_staff_from_cleaners(cleaners, people=None):
    try:
        limit = int(people) if people not in (None, "") else None
    except Exception:
        limit = None

    staff = normalize_staff_display(cleaners or [], limit=limit)
    return staff if staff else "無人力"


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
    joined = "\n".join(lines)
    normalized = normalize_text_for_parse(joined)

    m = re.search(r'([\u4e00-\u9fffA-Za-z0-9]+\(\d+\))X([\u4e00-\u9fffA-Za-z0-9]+\(\d+\))', normalized)
    if m:
        return normalize_staff_display([m.group(1), m.group(2)])

    for i, line in enumerate(lines):
        t = normalize_text_for_parse(line)
        if re.search(r'\(\d+\)X$', t):
            first = re.sub(r'X$', '', t)
            if i + 1 < len(lines):
                second = normalize_text_for_parse(lines[i + 1])
                if re.search(r'\(\d+\)$', second):
                    return normalize_staff_display([first, second])

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
    last_purchase = member_payload.get("lastPurchase", {}) if isinstance(member_payload, dict) else {}
    old_purchase = address_info.get("purchase", {}) if isinstance(address_info, dict) else {}

    def pick(key, default=""):
        if old_purchase.get(key) not in (None, ""):
            return old_purchase.get(key)
        if last_purchase.get(key) not in (None, ""):
            return last_purchase.get(key)
        return default

    def pick_address_notice(default=""):
        # 客服備註必須以「該地址系統預設帶出的備註」為準。
        # 不使用 lastPurchase.notice，避免抓到會員其他地址或最後一筆訂單的備註。
        if address_info.get("notice") not in (None, ""):
            return address_info.get("notice")
        if old_purchase.get("notice") not in (None, ""):
            return old_purchase.get("notice")
        if old_purchase.get("service_notice") not in (None, ""):
            return old_purchase.get("service_notice")
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


def filter_dates_by_balance(date_slots, date_prices, stored_value):
    # 只用服務費 price 判斷，車馬費不算在儲值金
    selected_slots, selected_prices, total = [], [], 0
    for slot, price in zip(date_slots, date_prices):
        if total + price <= stored_value:
            selected_slots.append(slot)
            selected_prices.append(price)
            total += price
    return selected_slots, selected_prices, total


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
    _, row0 = rows_with_idx[0]

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

    print("[DEBUG] parsed person/hour =", {
        "服務人時": str(row0["服務人時"]),
        "sheet_time": normalize_period_text(row0["開始時間"], row0["結束時間"]),
        "person": people,
        "hour": hours,
    })
    try:
        if st is not None:
            st.write("👥 parsed person/hour =", {
                "服務人時": str(row0["服務人時"]),
                "sheet_time": normalize_period_text(row0["開始時間"], row0["結束時間"]),
                "person": people,
                "hour": hours,
            })
    except Exception:
        pass

    phone = normalize_phone(row0["電話"])
    member_payload = get_member(session, phone, token, clean_type_id)
    if not member_payload:
        raise Exception(f"會員不存在: {phone}")

    member = member_payload.get("member", {})
    stored_value = int(float(member_payload.get("storedValue", 0) or 0))

    target_address = str(row0["地址"]).strip().split(",")[0]
    best_addr = pick_best_address_info(member_payload, target_address)
    if not best_addr:
        raise Exception("找不到對應地址資料")
    if not str(best_addr.get("addressId", "")).strip():
        raise Exception(f"地址存在但未選到下拉地址，缺少 addressId：{target_address}")

    selected_address = str(best_addr.get("address") or target_address).strip()

    geo_lat, geo_lng = geocode_address(selected_address)
    if geo_lat and geo_lng:
        best_addr["lat"] = geo_lat
        best_addr["lng"] = geo_lng

    addr_check = check_contain(
        session,
        member.get("member_id", ""),
        selected_address,
        best_addr.get("lat", ""),
        best_addr.get("lng", ""),
        token,
        clean_type_id,
    )
    if not addr_check:
        raise Exception(f"查詢地址/地區失敗：{selected_address}")

    # 確認是否真的有模擬按下「查詢地址」
    print("[DEBUG] check_contain raw =", addr_check)
    try:
        if st is not None:
            st.write("check_contain raw =", addr_check)
    except Exception:
        pass

    area_info = addr_check.get("area") if isinstance(addr_check.get("area"), dict) else {}
    purchase_info = addr_check.get("purchase") if isinstance(addr_check.get("purchase"), dict) else {}

    if area_info:
        best_addr["area_id"] = area_info.get("area_id", best_addr.get("area_id"))
        best_addr["company_id"] = area_info.get("company_id", best_addr.get("company_id"))
        best_addr["country_id"] = area_info.get("country_id", best_addr.get("country_id"))

    if purchase_info:
        best_addr["purchase"] = purchase_info

    # 模擬後台「查詢地址」後的資料補齊：
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
    notice_from_check = (
        (purchase_info.get("notice") if purchase_info else "")
        or (purchase_info.get("service_notice") if purchase_info else "")
        or find_nested_value(purchase_info, ["notice", "service_notice", "memo_notice", "customer_service_notice"])
        or best_addr.get("notice", "")
        or ""
    )
    best_addr["notice"] = notice_from_check

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

        print("[DEBUG] calculate_hour raw =", calc_result)
        try:
            if st is not None:
                st.write("🟠 calculate_hour raw =", calc_result)
        except Exception:
            pass

        calc_fields = extract_calc_fields(
            calc_result,
            fallback_hours=base_data.get("hour", ""),
            fallback_fare=best_addr.get("fare", "0"),
        )

        payload = base_data.copy()
        payload["date_s"] = date_s
        payload["hour"] = str(base_data.get("hour") or calc_fields.get("hour") or "")
        payload["person"] = str(base_data.get("person") or payload.get("person") or "")
        payload["price"] = str(calc_fields.get("price") or "0")
        payload["price_vvip"] = str(calc_fields.get("price_vvip") or "0")

        print("[DEBUG] calc_fields =", calc_fields)
        try:
            if st is not None:
                st.write("🟣 calc_fields =", calc_fields)
        except Exception:
            pass
        payload["fare"] = first_nonzero(calc_fields.get("fare"), best_addr.get("fare"), base_data.get("fare"), default="0")

        if str(payload.get("price", "")).strip() in ("", "0", "0.0"):
            raise Exception(f"計算時數後 price 仍為 0，請貼 🟠 calculate_hour raw 與 🟣 calc_fields：{date_s}")

        payload["notice"] = str(base_data.get("notice") or best_addr.get("notice") or "")
        payload["area_id"] = str(base_data.get("area_id") or best_addr.get("area_id") or "")
        payload["company_id"] = str(base_data.get("company_id") or best_addr.get("company_id") or "")
        payload["addressId"] = str(base_data.get("addressId") or best_addr.get("addressId") or "")
        return payload

    row_details = []
    for row_num, row in rows_with_idx:
        date_s = get_date_str(row["日期"])
        priced_payload = build_priced_payload_for_date(date_s)

        row_details.append({
            "row_num": row_num,
            "date": date_s,
            "slot": f"{date_s}_{system_period}",
            "price": int(float(priced_payload.get("price") or 0)),  # 只拿服務費比對儲值金
            "display_period": system_display_period,
            "row": row,
            "payload": priced_payload,
        })

        print("[DEBUG] row slot =", {
            "row_num": row_num,
            "sheet_time": normalize_period_text(row["開始時間"], row["結束時間"]),
            "system_period": system_period,
            "slot": f"{date_s}_{system_period}",
            "price": priced_payload.get("price"),
            "fare": priced_payload.get("fare"),
        })
        try:
            if st is not None:
                st.write("🧭 row slot =", {
                    "row_num": row_num,
                    "sheet_time": normalize_period_text(row["開始時間"], row["結束時間"]),
                    "system_period": system_period,
                    "slot": f"{date_s}_{system_period}",
                    "price": priced_payload.get("price"),
                    "fare": priced_payload.get("fare"),
                })
        except Exception:
            pass

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
        raw = get_section_raw(session, detail["payload"], token, detail["slot"])
        slot_ok = slot_exists_in_section_response(raw, detail["slot"])
        cleaners = extract_cleaners_from_section_response(raw, detail["slot"])
        detail["section_cleaners"] = cleaners
        detail["section_staff"] = format_staff_from_cleaners(cleaners, people=people)

        print("[DEBUG] section match =", {
            "slot": detail["slot"],
            "matched": slot_ok,
            "staff": detail.get("section_staff"),
            "raw_preview": str(raw)[:500],
        })
        try:
            if st is not None:
                st.write("🧩 section match =", {
                    "slot": detail["slot"],
                    "matched": slot_ok,
                    "staff": detail.get("section_staff"),
                    "raw_preview": str(raw)[:500],
                })
        except Exception:
            pass

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
    send_slots, _, _ = filter_dates_by_balance(valid_slots_for_balance, valid_prices_for_balance, stored_value)

    insufficient_dates = []
    send_details = []

    for detail in valid_details:
        if detail["slot"] in send_slots:
            send_details.append(detail)
        else:
            insufficient_dates.append(detail["date"])

    for detail in row_details:
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
        payload = detail["payload"].copy()
        slots = [detail["slot"]]

        print("[DEBUG] booking payload =",
              {
                  "date": detail["date"],
                  "slot": detail["slot"],
                  "price": payload.get("price"),
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
        logger(f"處理第 {group_no} 組：{first_row['姓名']}，共 {len(rows_with_idx)} 筆")

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
